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
        public const int MaxRefNodes = 200;
        public const int MaxTreeChars = 24000;
        public const int DefaultDepth = 12;

        private static readonly HashSet<string> InteractiveRoles = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "button", "link", "textbox", "searchbox", "combobox", "checkbox", "radio",
            "switch", "tab", "menuitem", "menuitemcheckbox", "menuitemradio", "option",
            "slider", "spinbutton", "treeitem", "listbox", "checkable"
        };

        private static readonly HashSet<string> StructureRoles = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "heading", "RootWebArea", "WebArea", "banner", "navigation", "main", "contentinfo"
        };

        private static readonly HashSet<string> FrameRoles = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "Iframe", "iframe", "Frame", "frame"
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

            string startId = null;
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

                if (node.Ignored && !IsRootLike(node.Role))
                {
                    foreach (string child in node.ChildIds)
                    {
                        Walk(child, depth);
                    }

                    return;
                }

                string role = string.IsNullOrWhiteSpace(node.Role) ? "generic" : node.Role;
                string name = node.Name ?? "";

                if (FrameRoles.Contains(role))
                {
                    AppendLine(sb, depth, role, name, refId: null, note: "未展开 iframe");
                    if (sb.Length >= MaxTreeChars)
                    {
                        truncated = true;
                        truncatedReason = "tree_chars>" + MaxTreeChars;
                    }

                    return;
                }

                bool wantRef = ShouldAssignRef(role, name, refs.Count);
                string refId = null;
                if (wantRef && refs.Count < MaxRefNodes)
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
                else if (wantRef && refs.Count >= MaxRefNodes)
                {
                    truncated = true;
                    truncatedReason = "ref_nodes>" + MaxRefNodes;
                }

                bool print = wantRef
                    || StructureRoles.Contains(role)
                    || (!string.IsNullOrWhiteSpace(name) && !IsNoiseRole(role))
                    || depth == 0;

                if (print)
                {
                    AppendLine(sb, depth, role, name, refId, note: null);
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

            result.Success = true;
            result.TreeText = sb.ToString().TrimEnd();
            result.Refs = refs;
            result.Truncated = truncated;
            result.TruncatedReason = truncatedReason;
            return result;
        }

        private static bool ShouldAssignRef(string role, string name, int currentRefCount)
        {
            if (currentRefCount >= MaxRefNodes)
            {
                return false;
            }

            if (FrameRoles.Contains(role))
            {
                return false;
            }

            if (InteractiveRoles.Contains(role))
            {
                // 可交互：有名字优先；无名字仍给 ref（否则点不到无名按钮）
                return true;
            }

            if (string.Equals(role, "heading", StringComparison.OrdinalIgnoreCase)
                && !string.IsNullOrWhiteSpace(name))
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
            public bool Ignored { get; set; }
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
                    Ignored = jo["ignored"] != null && jo["ignored"].Type == JTokenType.Boolean && (bool)jo["ignored"],
                    BackendDomNodeId = backend,
                    ChildIds = childIds
                };
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
