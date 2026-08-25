using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Newtonsoft.Json.Linq;

namespace WordAddIn1.BrowserHost
{
    /// <summary>
    /// CDP Accessibility.getFullAXTree JSON → 缩进文本树 + ref 表（S6～S11）。
    /// </summary>
    internal static class BrowserAxTreeBuilder
    {
        public const int MaxRefNodes = 300;
        public const int MaxTreeChars = 36000;
        public const int DefaultDepth = 12;

        private static readonly HashSet<string> InteractiveRoles = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "button", "link", "textbox", "searchbox", "combobox", "checkbox", "radio",
            "switch", "tab", "menuitem", "menuitemcheckbox", "menuitemradio", "option",
            "slider", "spinbutton", "treeitem", "listbox", "checkable",
            "text field", "search box", "combo box"
        };

        /// <summary>可输入控件：百度等站点常标 ignored，overview 仍须发 ref。</summary>
        private static readonly HashSet<string> InputRoles = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "textbox", "searchbox", "combobox", "text field", "search box", "combo box"
        };

        private static readonly HashSet<string> StructureRoles = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "heading", "RootWebArea", "WebArea", "banner", "navigation", "main", "contentinfo"
        };

        private static readonly HashSet<string> FrameRoles = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "Iframe", "iframe", "Frame", "frame"
        };

        /// <summary>hover 才显现、AX 常标 ignored：仍发 ref。不含 display:none 模板。</summary>
        private static readonly HashSet<string> HoverRevealRoles = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "button", "link", "menuitem", "menuitemcheckbox", "menuitemradio", "tab"
        };

        public static BrowserAxBuildResult BuildOverview(string axTreeJson)
        {
            return Build(axTreeJson, detailRootRef: null, priorEntry: null);
        }

        public static BrowserAxBuildResult BuildDetail(
            string axTreeJson,
            string detailRef,
            BrowserRefEntry priorEntry)
        {
            return Build(axTreeJson, detailRootRef: detailRef, priorEntry: priorEntry);
        }

        private static BrowserAxBuildResult Build(
            string axTreeJson,
            string detailRootRef,
            BrowserRefEntry priorEntry)
        {
            var result = new BrowserAxBuildResult();
            if (string.IsNullOrWhiteSpace(axTreeJson))
            {
                result.Error = "无障碍树为空";
                return result;
            }

            JObject root;
            try
            {
                root = JObject.Parse(axTreeJson);
            }
            catch (Exception ex)
            {
                result.Error = "解析无障碍树失败: " + ex.Message;
                return result;
            }

            var nodesArr = root["nodes"] as JArray;
            if (nodesArr == null || nodesArr.Count == 0)
            {
                result.Error = "无障碍树无节点";
                return result;
            }

            var byId = new Dictionary<string, AxNode>(StringComparer.Ordinal);
            foreach (var token in nodesArr)
            {
                if (!(token is JObject jo))
                {
                    continue;
                }

                var node = AxNode.FromJson(jo);
                if (node == null || string.IsNullOrEmpty(node.NodeId))
                {
                    continue;
                }

                byId[node.NodeId] = node;
            }

            if (byId.Count == 0)
            {
                result.Error = "无障碍树无有效节点";
                return result;
            }

            string startId;
            if (!string.IsNullOrWhiteSpace(detailRootRef) && priorEntry != null)
            {
                startId = FindNodeId(byId, priorEntry);
                if (string.IsNullOrEmpty(startId))
                {
                    result.Error = "ref 无效或已过期，请重新 F_browser_snapshot";
                    return result;
                }
            }
            else
            {
                startId = FindRootId(byId);
            }

            if (string.IsNullOrEmpty(startId) || !byId.ContainsKey(startId))
            {
                result.Error = "找不到页面根节点";
                return result;
            }

            var sb = new StringBuilder();
            var refs = new Dictionary<string, BrowserRefEntry>(StringComparer.OrdinalIgnoreCase);
            int nextRef = 1;
            bool truncated = false;
            string truncatedReason = null;
            var visited = new HashSet<string>(StringComparer.Ordinal);

            void Walk(string nodeId, int depth)
            {
                if (truncated || string.IsNullOrEmpty(nodeId) || !byId.TryGetValue(nodeId, out AxNode node))
                {
                    return;
                }

                if (!visited.Add(nodeId))
                {
                    return;
                }

                if (sb.Length >= MaxTreeChars)
                {
                    truncated = true;
                    truncatedReason = "tree_chars>" + MaxTreeChars;
                    return;
                }

                // 临时关闭「AX ignored 可输入仍发 ref」：验证仅 DOM 补查是否足够
                // bool inputLike = IsInputLike(node);
                // if (node.Ignored && !IsRootLike(node.Role) && !inputLike)
                string peekRole = string.IsNullOrWhiteSpace(node.Role) ? "generic" : node.Role;
                bool hoverReveal = HoverRevealRoles.Contains(peekRole);
                if (node.Ignored && !IsRootLike(node.Role) && !FrameRoles.Contains(peekRole) && !hoverReveal)
                {
                    foreach (string child in node.ChildIds)
                    {
                        Walk(child, depth);
                    }

                    return;
                }

                bool inputLike = IsInputLike(node);

                string role = peekRole;
                string name = ResolveFrameName(node, role);

                if (FrameRoles.Contains(role))
                {
                    string frameRef = null;
                    if (refs.Count < MaxRefNodes)
                    {
                        frameRef = "e" + nextRef;
                        nextRef++;
                        refs[frameRef] = new BrowserRefEntry
                        {
                            AxNodeId = node.NodeId,
                            BackendDomNodeId = node.BackendDomNodeId,
                            Role = role,
                            Name = name
                        };
                    }
                    else
                    {
                        truncated = true;
                        truncatedReason = "ref_nodes>" + MaxRefNodes;
                    }

                    AppendLine(sb, depth, role, name, frameRef, note: null);
                    if (sb.Length >= MaxTreeChars)
                    {
                        truncated = true;
                        truncatedReason = "tree_chars>" + MaxTreeChars;
                    }

                    return;
                }

                bool wantRef = ShouldAssignRef(node, refs.Count);
                string refId = null;
                if (wantRef && refs.Count < MaxRefNodes && node.BackendDomNodeId.HasValue)
                {
                    refId = "e" + nextRef;
                    nextRef++;
                    refs[refId] = new BrowserRefEntry
                    {
                        AxNodeId = node.NodeId,
                        BackendDomNodeId = node.BackendDomNodeId,
                        Role = role,
                        Name = name
                    };
                }
                else if (wantRef && !node.BackendDomNodeId.HasValue)
                {
                    // 无可操作 backend id：仍打印，不发 ref
                    wantRef = false;
                }
                else if (wantRef && refs.Count >= MaxRefNodes)
                {
                    truncated = true;
                    truncatedReason = "ref_nodes>" + MaxRefNodes;
                }

                bool print = wantRef
                    || inputLike
                    || StructureRoles.Contains(role)
                    || (!string.IsNullOrWhiteSpace(name) && !IsNoiseRole(role))
                    || depth == 0;

                if (print)
                {
                    string note = (hoverReveal && node.Ignored) ? "隐藏" : null;
                    AppendLine(sb, depth, role, name, refId, note);
                    if (sb.Length >= MaxTreeChars)
                    {
                        truncated = true;
                        truncatedReason = "tree_chars>" + MaxTreeChars;
                        return;
                    }
                }

                int childDepth = print ? depth + 1 : depth;
                foreach (string child in node.ChildIds)
                {
                    if (truncated)
                    {
                        break;
                    }

                    Walk(child, childDepth);
                }
            }

            Walk(startId, 0);

            // 临时关闭：AX 树第二遍补全可输入控件（与 DOM 补查对照实验）
            // if (!truncated)
            // {
            //     AppendMissingInputControls(byId, refs, sb, ref nextRef, ref truncated, ref truncatedReason);
            // }

            result.Success = true;
            result.TreeText = sb.ToString().TrimEnd();
            result.Refs = refs;
            result.Truncated = truncated;
            result.TruncatedReason = truncatedReason;
            return result;
        }

        /// <summary>补全遍历仍漏掉的可输入节点（断链 / ignored 父级未走到等）。</summary>
        private static void AppendMissingInputControls(
            Dictionary<string, AxNode> byId,
            Dictionary<string, BrowserRefEntry> refs,
            StringBuilder sb,
            ref int nextRef,
            ref bool truncated,
            ref string truncatedReason)
        {
            var already = new HashSet<int>();
            foreach (var e in refs.Values)
            {
                if (e.BackendDomNodeId.HasValue)
                {
                    already.Add(e.BackendDomNodeId.Value);
                }
            }

            bool header = false;
            foreach (var node in byId.Values.OrderBy(n => n.NodeId, StringComparer.Ordinal))
            {
                if (truncated || sb.Length >= MaxTreeChars)
                {
                    truncated = true;
                    truncatedReason = truncatedReason ?? ("tree_chars>" + MaxTreeChars);
                    return;
                }

                if (!IsInputLike(node) || !node.BackendDomNodeId.HasValue)
                {
                    continue;
                }

                if (already.Contains(node.BackendDomNodeId.Value))
                {
                    continue;
                }

                if (refs.Count >= MaxRefNodes)
                {
                    truncated = true;
                    truncatedReason = "ref_nodes>" + MaxRefNodes;
                    return;
                }

                if (!header)
                {
                    sb.AppendLine("- (补全) 可输入控件");
                    header = true;
                }

                string role = string.IsNullOrWhiteSpace(node.Role) ? "textbox" : node.Role;
                string name = node.Name ?? "";
                string refId = "e" + nextRef;
                nextRef++;
                refs[refId] = new BrowserRefEntry
                {
                    AxNodeId = node.NodeId,
                    BackendDomNodeId = node.BackendDomNodeId,
                    Role = role,
                    Name = name
                };
                already.Add(node.BackendDomNodeId.Value);
                AppendLine(sb, 1, role, name, refId, note: null);
            }
        }

        private static bool IsInputLike(AxNode node)
        {
            if (node == null)
            {
                return false;
            }

            if (InputRoles.Contains(node.Role ?? ""))
            {
                return true;
            }

            return node.Editable;
        }

        private static bool ShouldAssignRef(AxNode node, int currentRefCount)
        {
            if (node == null || currentRefCount >= MaxRefNodes)
            {
                return false;
            }

            string role = node.Role ?? "";
            if (FrameRoles.Contains(role))
            {
                return true;
            }

            if (IsInputLike(node))
            {
                return true;
            }

            if (InteractiveRoles.Contains(role))
            {
                return true;
            }

            if (string.Equals(role, "heading", StringComparison.OrdinalIgnoreCase)
                && !string.IsNullOrWhiteSpace(node.Name))
            {
                return true;
            }

            return false;
        }

        private static bool IsNoiseRole(string role)
        {
            return string.Equals(role, "InlineTextBox", StringComparison.OrdinalIgnoreCase)
                || string.Equals(role, "LineBreak", StringComparison.OrdinalIgnoreCase)
                || string.Equals(role, "generic", StringComparison.OrdinalIgnoreCase)
                || string.Equals(role, "Group", StringComparison.OrdinalIgnoreCase)
                || string.Equals(role, "LayoutTable", StringComparison.OrdinalIgnoreCase)
                || string.Equals(role, "LayoutTableRow", StringComparison.OrdinalIgnoreCase)
                || string.Equals(role, "LayoutTableCell", StringComparison.OrdinalIgnoreCase);
        }

        private static bool IsRootLike(string role)
        {
            return string.Equals(role, "RootWebArea", StringComparison.OrdinalIgnoreCase)
                || string.Equals(role, "WebArea", StringComparison.OrdinalIgnoreCase);
        }

        private static string ResolveFrameName(AxNode node, string role)
        {
            string name = node?.Name ?? "";
            if (!FrameRoles.Contains(role ?? ""))
            {
                return name;
            }

            if (!string.IsNullOrWhiteSpace(name))
            {
                return name.Trim();
            }

            if (node != null && !string.IsNullOrWhiteSpace(node.UrlHint))
            {
                string tail = SrcTail(node.UrlHint);
                if (!string.IsNullOrWhiteSpace(tail))
                {
                    return tail;
                }
            }

            return "iframe";
        }

        private static string SrcTail(string url)
        {
            if (string.IsNullOrWhiteSpace(url))
            {
                return "";
            }

            try
            {
                var u = new Uri(url.Trim(), UriKind.Absolute);
                if (u.Segments != null && u.Segments.Length > 0)
                {
                    string last = (u.Segments[u.Segments.Length - 1] ?? "").Trim('/');
                    if (!string.IsNullOrWhiteSpace(last))
                    {
                        return last;
                    }
                }

                return u.Host ?? "";
            }
            catch
            {
                string s = url.Trim();
                int slash = s.LastIndexOf('/');
                return slash >= 0 && slash < s.Length - 1 ? s.Substring(slash + 1) : s;
            }
        }

        private static void AppendLine(
            StringBuilder sb,
            int depth,
            string role,
            string name,
            string refId,
            string note)
        {
            for (int i = 0; i < depth; i++)
            {
                sb.Append("  ");
            }

            sb.Append("- ").Append(role);
            if (!string.IsNullOrEmpty(refId))
            {
                sb.Append(" ref=").Append(refId);
            }

            if (!string.IsNullOrWhiteSpace(name))
            {
                sb.Append(" \"").Append(Escape(name.Trim())).Append('"');
            }

            if (!string.IsNullOrWhiteSpace(note))
            {
                sb.Append(" (").Append(note).Append(')');
            }

            sb.AppendLine();
        }

        private static string Escape(string s)
        {
            return s.Replace("\\", "\\\\").Replace("\"", "\\\"").Replace("\r", " ").Replace("\n", " ");
        }

        private static string FindRootId(Dictionary<string, AxNode> byId)
        {
            var childOfSomeone = new HashSet<string>(StringComparer.Ordinal);
            foreach (var n in byId.Values)
            {
                foreach (string c in n.ChildIds)
                {
                    childOfSomeone.Add(c);
                }
            }

            AxNode rootWeb = byId.Values.FirstOrDefault(n =>
                string.Equals(n.Role, "RootWebArea", StringComparison.OrdinalIgnoreCase));
            if (rootWeb != null)
            {
                return rootWeb.NodeId;
            }

            foreach (var n in byId.Values)
            {
                if (!childOfSomeone.Contains(n.NodeId))
                {
                    return n.NodeId;
                }
            }

            return byId.Keys.FirstOrDefault();
        }

        private static string FindNodeId(Dictionary<string, AxNode> byId, BrowserRefEntry prior)
        {
            if (prior == null)
            {
                return null;
            }

            if (!string.IsNullOrEmpty(prior.AxNodeId)
                && byId.ContainsKey(prior.AxNodeId))
            {
                return prior.AxNodeId;
            }

            if (prior.BackendDomNodeId.HasValue)
            {
                foreach (var n in byId.Values)
                {
                    if (n.BackendDomNodeId == prior.BackendDomNodeId)
                    {
                        return n.NodeId;
                    }
                }
            }

            return null;
        }

        private sealed class AxNode
        {
            public string NodeId { get; set; }
            public string Role { get; set; }
            public string Name { get; set; }
            public string UrlHint { get; set; }
            public bool Ignored { get; set; }
            public bool Editable { get; set; }
            public int? BackendDomNodeId { get; set; }
            public List<string> ChildIds { get; set; }

            public static AxNode FromJson(JObject jo)
            {
                if (jo == null)
                {
                    return null;
                }

                string id = TokenToId(jo["nodeId"]);
                if (string.IsNullOrEmpty(id))
                {
                    return null;
                }

                var childIds = new List<string>();
                if (jo["childIds"] is JArray kids)
                {
                    foreach (var k in kids)
                    {
                        string cid = TokenToId(k);
                        if (!string.IsNullOrEmpty(cid))
                        {
                            childIds.Add(cid);
                        }
                    }
                }

                int? backend = null;
                if (jo["backendDOMNodeId"] != null
                    && jo["backendDOMNodeId"].Type != JTokenType.Null
                    && int.TryParse(jo["backendDOMNodeId"].ToString(), out int bid))
                {
                    backend = bid;
                }

                return new AxNode
                {
                    NodeId = id,
                    Role = ReadAxValue(jo["role"]),
                    Name = ReadAxValue(jo["name"]),
                    UrlHint = ReadUrlProperty(jo["properties"]),
                    Ignored = jo["ignored"] != null && jo["ignored"].Type == JTokenType.Boolean && (bool)jo["ignored"],
                    Editable = ReadEditable(jo["properties"]),
                    BackendDomNodeId = backend,
                    ChildIds = childIds
                };
            }

            private static string ReadUrlProperty(JToken propertiesToken)
            {
                if (!(propertiesToken is JArray props))
                {
                    return "";
                }

                foreach (var p in props)
                {
                    if (!(p is JObject po))
                    {
                        continue;
                    }

                    string pname = Convert.ToString(po["name"]) ?? "";
                    if (!string.Equals(pname, "url", StringComparison.OrdinalIgnoreCase))
                    {
                        continue;
                    }

                    return ReadAxValue(po["value"]);
                }

                return "";
            }

            private static bool ReadEditable(JToken propertiesToken)
            {
                if (!(propertiesToken is JArray props))
                {
                    return false;
                }

                foreach (var p in props)
                {
                    if (!(p is JObject po))
                    {
                        continue;
                    }

                    string pname = Convert.ToString(po["name"]) ?? "";
                    if (!string.Equals(pname, "editable", StringComparison.OrdinalIgnoreCase))
                    {
                        continue;
                    }

                    // editable: boolean true，或 token "plaintext" / "richtext" 等
                    var val = po["value"];
                    if (val == null || val.Type == JTokenType.Null)
                    {
                        return true;
                    }

                    if (val.Type == JTokenType.Boolean)
                    {
                        return (bool)val;
                    }

                    if (val is JObject vo)
                    {
                        if (vo["value"] != null && vo["value"].Type == JTokenType.Boolean)
                        {
                            return (bool)vo["value"];
                        }

                        string s = Convert.ToString(vo["value"]) ?? "";
                        if (string.IsNullOrWhiteSpace(s))
                        {
                            return false;
                        }

                        if (string.Equals(s, "false", StringComparison.OrdinalIgnoreCase)
                            || string.Equals(s, "none", StringComparison.OrdinalIgnoreCase))
                        {
                            return false;
                        }

                        return true;
                    }

                    string raw = Convert.ToString(val) ?? "";
                    return !string.Equals(raw, "false", StringComparison.OrdinalIgnoreCase)
                        && !string.Equals(raw, "none", StringComparison.OrdinalIgnoreCase)
                        && !string.IsNullOrWhiteSpace(raw);
                }

                return false;
            }

            private static string TokenToId(JToken t)
            {
                if (t == null || t.Type == JTokenType.Null)
                {
                    return null;
                }

                return Convert.ToString(t)?.Trim();
            }

            private static string ReadAxValue(JToken t)
            {
                if (t == null || t.Type == JTokenType.Null)
                {
                    return "";
                }

                if (t is JObject o)
                {
                    if (o["value"] != null && o["value"].Type != JTokenType.Null)
                    {
                        return Convert.ToString(o["value"]) ?? "";
                    }
                }

                return Convert.ToString(t) ?? "";
            }
        }
    }

    internal sealed class BrowserAxBuildResult
    {
        public bool Success { get; set; }
        public string Error { get; set; }
        public string TreeText { get; set; } = "";
        public bool Truncated { get; set; }
        public string TruncatedReason { get; set; }
        public Dictionary<string, BrowserRefEntry> Refs { get; set; }
            = new Dictionary<string, BrowserRefEntry>(StringComparer.OrdinalIgnoreCase);
    }
}
