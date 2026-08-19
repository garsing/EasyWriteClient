using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;

namespace WordAddIn1.BrowserHost
{
    /// <summary>
    /// AX 树漏掉真实 input 时，用 DOM.querySelectorAll 补进 snapshot / ref 表。
    /// 可由 F_browser_snapshot 的 dom_supplement 控制范围。
    /// </summary>
    internal static class BrowserDomInputSupplement
    {
        public const string DefaultSpec = "input,textarea,select,contenteditable";

        private const string InputSelector =
            "input:not([type=hidden]):not([type=submit]):not([type=button]):not([type=image]):not([type=checkbox]):not([type=radio]):not([type=file])";

        private const string ContenteditableSelector =
            "[contenteditable=true],[contenteditable=\"\"],[contenteditable=plaintext-only]";

        /// <summary>
        /// 解析 dom_supplement：省略/空/default/all → 默认集合；
        /// off/none/false → 关闭；否则按 , | ; 、 分隔的白名单项。
        /// </summary>
        public static bool TryResolveSelector(string spec, out string cssSelector, out string normalizedSpec, out string error)
        {
            cssSelector = null;
            normalizedSpec = null;
            error = null;

            if (string.IsNullOrWhiteSpace(spec)
                || string.Equals(spec.Trim(), "default", StringComparison.OrdinalIgnoreCase)
                || string.Equals(spec.Trim(), "all", StringComparison.OrdinalIgnoreCase))
            {
                normalizedSpec = DefaultSpec;
                cssSelector = BuildSelectorFromKinds(ParseKinds(DefaultSpec));
                return true;
            }

            string raw = spec.Trim();
            if (string.Equals(raw, "off", StringComparison.OrdinalIgnoreCase)
                || string.Equals(raw, "none", StringComparison.OrdinalIgnoreCase)
                || string.Equals(raw, "false", StringComparison.OrdinalIgnoreCase)
                || raw == "0")
            {
                normalizedSpec = "off";
                cssSelector = null;
                return true;
            }

            var kinds = ParseKinds(raw);
            if (kinds.Count == 0)
            {
                error = "dom_supplement 无有效项；可用 input,textarea,select,contenteditable 或 off";
                return false;
            }

            foreach (string k in kinds)
            {
                if (!IsAllowedKind(k))
                {
                    error = "dom_supplement 含未知项: " + k
                        + "；允许 input|textarea|select|contenteditable，或 off";
                    return false;
                }
            }

            normalizedSpec = string.Join(",", kinds);
            cssSelector = BuildSelectorFromKinds(kinds);
            return true;
        }

        public static async Task MergeAsync(YiWriteBrowserForm form, BrowserAxBuildResult built)
        {
            await MergeAsync(form, built, DefaultSpec).ConfigureAwait(true);
        }

        public static async Task MergeAsync(
            YiWriteBrowserForm form,
            BrowserAxBuildResult built,
            string domSupplementSpec)
        {
            if (form == null || built == null || !built.Success || built.Refs == null)
            {
                return;
            }

            if (!TryResolveSelector(domSupplementSpec, out string selector, out _, out _))
            {
                return;
            }

            if (string.IsNullOrEmpty(selector))
            {
                return;
            }

            try
            {
                string docJson = await form.CallCdpAsync(
                    "DOM.getDocument",
                    "{\"depth\":0,\"pierce\":true}").ConfigureAwait(true);
                var doc = JObject.Parse(docJson);
                if (doc["root"]?["nodeId"] == null)
                {
                    return;
                }

                int rootId = doc["root"]["nodeId"].Value<int>();
                var qsPayload = new JObject
                {
                    ["nodeId"] = rootId,
                    ["selector"] = selector
                };
                string qsJson = await form.CallCdpAsync(
                    "DOM.querySelectorAll",
                    qsPayload.ToString(Newtonsoft.Json.Formatting.None)).ConfigureAwait(true);
                var qs = JObject.Parse(qsJson);
                var nodeIds = qs["nodeIds"] as JArray;
                if (nodeIds == null || nodeIds.Count == 0)
                {
                    return;
                }

                var already = new HashSet<int>();
                int maxRef = 0;
                foreach (var kv in built.Refs)
                {
                    if (kv.Value?.BackendDomNodeId != null)
                    {
                        already.Add(kv.Value.BackendDomNodeId.Value);
                    }

                    if (kv.Key != null
                        && kv.Key.Length > 1
                        && kv.Key[0] == 'e'
                        && int.TryParse(kv.Key.Substring(1), out int n)
                        && n > maxRef)
                    {
                        maxRef = n;
                    }
                }

                var sb = new StringBuilder(built.TreeText ?? "");
                bool header = false;
                int nextRef = maxRef + 1;

                foreach (var idTok in nodeIds)
                {
                    if (built.Refs.Count >= BrowserAxTreeBuilder.MaxRefNodes)
                    {
                        built.Truncated = true;
                        built.TruncatedReason = "ref_nodes>" + BrowserAxTreeBuilder.MaxRefNodes;
                        break;
                    }

                    if (sb.Length >= BrowserAxTreeBuilder.MaxTreeChars)
                    {
                        built.Truncated = true;
                        built.TruncatedReason = "tree_chars>" + BrowserAxTreeBuilder.MaxTreeChars;
                        break;
                    }

                    if (idTok == null || idTok.Type == JTokenType.Null)
                    {
                        continue;
                    }

                    int nodeId = idTok.Value<int>();
                    string descJson = await form.CallCdpAsync(
                        "DOM.describeNode",
                        "{\"nodeId\":" + nodeId.ToString(CultureInfo.InvariantCulture) + ",\"pierce\":true}")
                        .ConfigureAwait(true);
                    var desc = JObject.Parse(descJson);
                    var node = desc["node"] as JObject;
                    if (node == null)
                    {
                        continue;
                    }

                    if (node["backendNodeId"] == null || node["backendNodeId"].Type == JTokenType.Null)
                    {
                        continue;
                    }

                    int backendId = node["backendNodeId"].Value<int>();
                    if (already.Contains(backendId))
                    {
                        continue;
                    }

                    string tag = (Convert.ToString(node["nodeName"]) ?? "INPUT").ToUpperInvariant();
                    var attrs = ParseAttributes(node["attributes"] as JArray);
                    string type = GetAttr(attrs, "type");

                    string role = InferRole(tag, type, attrs);
                    string name = FirstNonEmpty(
                        GetAttr(attrs, "aria-label"),
                        GetAttr(attrs, "placeholder"),
                        GetAttr(attrs, "title"),
                        GetAttr(attrs, "name"),
                        GetAttr(attrs, "id"));

                    if (!header)
                    {
                        if (sb.Length > 0 && !sb.ToString().EndsWith("\n"))
                        {
                            sb.AppendLine();
                        }

                        sb.AppendLine("- (DOM补全) 可输入控件");
                        header = true;
                    }

                    string refId = "e" + nextRef;
                    nextRef++;
                    built.Refs[refId] = new BrowserRefEntry
                    {
                        AxNodeId = null,
                        BackendDomNodeId = backendId,
                        Role = role,
                        Name = name ?? ""
                    };
                    already.Add(backendId);

                    sb.Append("  - ").Append(role).Append(" ref=").Append(refId);
                    if (!string.IsNullOrWhiteSpace(name))
                    {
                        sb.Append(" \"").Append(Escape(name.Trim())).Append('"');
                    }

                    sb.AppendLine();
                }

                if (header)
                {
                    built.TreeText = sb.ToString().TrimEnd();
                }
            }
            catch
            {
                // DOM 补全失败不阻断 AX snapshot
            }
        }

        private static List<string> ParseKinds(string raw)
        {
            var list = new List<string>();
            if (string.IsNullOrWhiteSpace(raw))
            {
                return list;
            }

            string[] parts = raw.Split(new[] { ',', '|', ';', '、' }, StringSplitOptions.RemoveEmptyEntries);
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (string p in parts)
            {
                string k = (p ?? "").Trim().ToLowerInvariant();
                if (k.Length == 0 || !seen.Add(k))
                {
                    continue;
                }

                list.Add(k);
            }

            return list;
        }

        private static bool IsAllowedKind(string kind)
        {
            return string.Equals(kind, "input", StringComparison.OrdinalIgnoreCase)
                || string.Equals(kind, "textarea", StringComparison.OrdinalIgnoreCase)
                || string.Equals(kind, "select", StringComparison.OrdinalIgnoreCase)
                || string.Equals(kind, "contenteditable", StringComparison.OrdinalIgnoreCase);
        }

        private static string BuildSelectorFromKinds(List<string> kinds)
        {
            var parts = new List<string>();
            foreach (string k in kinds)
            {
                if (string.Equals(k, "input", StringComparison.OrdinalIgnoreCase))
                {
                    parts.Add(InputSelector);
                }
                else if (string.Equals(k, "textarea", StringComparison.OrdinalIgnoreCase))
                {
                    parts.Add("textarea");
                }
                else if (string.Equals(k, "select", StringComparison.OrdinalIgnoreCase))
                {
                    parts.Add("select");
                }
                else if (string.Equals(k, "contenteditable", StringComparison.OrdinalIgnoreCase))
                {
                    parts.Add(ContenteditableSelector);
                }
            }

            return parts.Count == 0 ? null : string.Join(",", parts);
        }

        private static Dictionary<string, string> ParseAttributes(JArray attrs)
        {
            var map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            if (attrs == null)
            {
                return map;
            }

            for (int i = 0; i + 1 < attrs.Count; i += 2)
            {
                string k = Convert.ToString(attrs[i]) ?? "";
                string v = Convert.ToString(attrs[i + 1]) ?? "";
                if (!string.IsNullOrEmpty(k))
                {
                    map[k] = v;
                }
            }

            return map;
        }

        private static string GetAttr(Dictionary<string, string> attrs, string key)
        {
            if (attrs == null || !attrs.TryGetValue(key, out string v))
            {
                return null;
            }

            return string.IsNullOrWhiteSpace(v) ? null : v;
        }

        private static string InferRole(string tag, string type, Dictionary<string, string> attrs)
        {
            string ariaRole = GetAttr(attrs, "role");
            if (!string.IsNullOrWhiteSpace(ariaRole))
            {
                return ariaRole.Trim().ToLowerInvariant();
            }

            if (string.Equals(tag, "TEXTAREA", StringComparison.OrdinalIgnoreCase))
            {
                return "textbox";
            }

            if (string.Equals(tag, "SELECT", StringComparison.OrdinalIgnoreCase))
            {
                return "combobox";
            }

            if (string.Equals(type, "search", StringComparison.OrdinalIgnoreCase))
            {
                return "searchbox";
            }

            return "textbox";
        }

        private static string FirstNonEmpty(params string[] parts)
        {
            foreach (string p in parts)
            {
                if (!string.IsNullOrWhiteSpace(p))
                {
                    return p;
                }
            }

            return "";
        }

        private static string Escape(string s)
        {
            return s.Replace("\\", "\\\\").Replace("\"", "\\\"").Replace("\r", " ").Replace("\n", " ");
        }
    }
}
