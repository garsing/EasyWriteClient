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
    internal static partial class BrowserAxTreeBuilder
    {
        public const int MaxRefNodes = 300;
        public const int MaxTreeChars = 36000;
        public const int DefaultDepth = 12;

        private static readonly HashSet<string> InteractiveRoles = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "button", "link", "textbox", "searchbox", "combobox", "checkbox", "radio",
            "switch", "tab", "menuitem", "menuitemcheckbox", "menuitemradio", "option",
            "slider", "spinbutton", "treeitem", "listbox", "checkable",
            "text field", "search box", "combo box",
            "date", "time", "datetime", "date time"
        };

        /// <summary>可输入控件：百度等站点常标 ignored，overview 仍须发 ref。</summary>
        private static readonly HashSet<string> InputRoles = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "textbox", "searchbox", "combobox", "text field", "search box", "combo box",
            "date", "time", "datetime", "date time"
        };

        private static readonly HashSet<string> StructureRoles = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "heading", "RootWebArea", "WebArea", "banner", "navigation", "main", "contentinfo",
            "row", "cell", "gridcell", "columnheader", "rowheader"
        };

        private static readonly HashSet<string> RowRoles = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "row", "LayoutTableRow"
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
            return Build(axTreeJson, detailRootRef: null, priorEntry: null, window: null);
        }

        public static BrowserAxBuildResult BuildOverview(string axTreeJson, RepeatWindowRequest window)
        {
            return Build(axTreeJson, detailRootRef: null, priorEntry: null, window: window);
        }

        public static BrowserAxBuildResult BuildDetail(
            string axTreeJson,
            string detailRef,
            BrowserRefEntry priorEntry)
        {
            return Build(axTreeJson, detailRootRef: detailRef, priorEntry: priorEntry, window: null);
        }

        private static BrowserAxBuildResult Build(
            string axTreeJson,
            string detailRootRef,
            BrowserRefEntry priorEntry,
            RepeatWindowRequest window)
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
            RepeatAnalysis analysis = AnalyzeRepeats(byId);
            int chromeLeft = analysis.ChromeSlots;
            bool pageHasRuns = analysis.RunsByParent.Count > 0;
            bool windowApplied = false;
            bool windowInvalid = false;

            string TakeRef(bool chrome)
            {
                if (chrome && chromeLeft > 0)
                {
                    chromeLeft--;
                }

                int limit = (!chrome && pageHasRuns)
                    ? MaxRefNodes - Math.Max(0, chromeLeft)
                    : MaxRefNodes;
                if (refs.Count >= limit)
                {
                    return null;
                }

                string id = "e" + nextRef;
                nextRef++;
                return id;
            }

            bool EmitRepeat(AxNode parent, RepeatRun run, int depth)
            {
                int start = 1;
                int end = Math.Min(RepeatWindowSize, run.Count);
                bool matched = window != null
                    && parent.BackendDomNodeId.HasValue
                    && parent.BackendDomNodeId.Value == window.ParentBackendId
                    && window.Ordinal == run.Ordinal
                    && string.Equals(window.Signature, run.Signature, StringComparison.Ordinal);
                if (matched)
                {
                    windowApplied = true;
                    if (window.Start > run.Count)
                    {
                        windowInvalid = true;
                        return false;
                    }

                    start = window.Start;
                    end = Math.Min(window.End, run.Count);
                    if (end - start + 1 > RepeatWindowSize)
                    {
                        end = start + RepeatWindowSize - 1;
                    }
                }

                string planned = "共 " + run.Count + " 项，当前 " + start + "-" + end;
                string headerRef = TakeRef(chrome: true);
                if (headerRef != null)
                {
                    refs[headerRef] = new BrowserRefEntry
                    {
                        AxNodeId = parent.NodeId,
                        Role = "repeat",
                        Name = planned,
                        IsRepeat = true,
                        RepeatParentBackendId = parent.BackendDomNodeId,
                        RepeatSignature = run.Signature,
                        RepeatOrdinal = run.Ordinal,
                        RepeatCount = run.Count
                    };
                    if (matched)
                    {
                        result.FocusRef = headerRef;
                    }
                }

                AppendLine(sb, depth, "repeat", planned, headerRef, note: null);

                bool prevTrunc = truncated;
                string prevReason = truncatedReason;
                Walk(run.MemberIds[start - 1], depth + 1, null, false);
                bool childTrunc = truncated;
                string childReason = childTrunc ? truncatedReason : null;
                truncated = prevTrunc;
                truncatedReason = prevReason;

                int printedEnd = start;
                bool missed = false;
                if (!(childTrunc && sb.Length >= MaxTreeChars))
                {
                    for (int n = start + 1; n <= end; n++)
                    {
                        if (sb.Length >= MaxTreeChars)
                        {
                            childTrunc = true;
                            childReason = "tree_chars>" + MaxTreeChars;
                            break;
                        }

                        if (!byId.TryGetValue(run.MemberIds[n - 1], out AxNode item))
                        {
                            continue;
                        }

                        visited.Add(item.NodeId);
                        string itemRole = string.IsNullOrWhiteSpace(item.Role) ? "generic" : item.Role;
                        string itemName = SanitizeAccessibleName(item.Name ?? "").Trim();
                        if (itemName.Length > RepeatNameChars)
                        {
                            itemName = itemName.Substring(0, RepeatNameChars) + "…";
                        }

                        string oneRef = null;
                        if (item.BackendDomNodeId.HasValue)
                        {
                            oneRef = TakeRef(chrome: false);
                            if (oneRef == null)
                            {
                                missed = true;
                            }
                            else
                            {
                                refs[oneRef] = new BrowserRefEntry
                                {
                                    AxNodeId = item.NodeId,
                                    BackendDomNodeId = item.BackendDomNodeId,
                                    Role = itemRole,
                                    Name = itemName
                                };
                            }
                        }

                        AppendLine(sb, depth + 1, itemRole, itemName, oneRef, note: null);
                        printedEnd = n;
                    }
                }

                if (printedEnd < end)
                {
                    string actual = "共 " + run.Count + " 项，当前 " + start + "-" + printedEnd;
                    sb.Replace(planned, actual);
                    if (headerRef != null && refs.TryGetValue(headerRef, out BrowserRefEntry header))
                    {
                        header.Name = actual;
                    }
                }

                if (childTrunc)
                {
                    truncated = true;
                    truncatedReason = childReason ?? truncatedReason;
                }

                return missed || childTrunc;
            }

            void VisitChildren(AxNode parent, int childDepth, string childInherit, bool childInCard)
            {
                analysis.RunsByParent.TryGetValue(parent.NodeId, out List<RepeatRun> childRuns);
                bool deferTrunc = false;
                string deferReason = null;
                for (int ci = 0; ci < parent.ChildIds.Count; ci++)
                {
                    if (truncated)
                    {
                        break;
                    }

                    RepeatRun runHere = null;
                    if (childRuns != null)
                    {
                        for (int r = 0; r < childRuns.Count; r++)
                        {
                            if (childRuns[r].StartIndex == ci)
                            {
                                runHere = childRuns[r];
                                break;
                            }
                        }
                    }

                    if (runHere != null)
                    {
                        bool prevTrunc = truncated;
                        string prevReason = truncatedReason;
                        bool missed = EmitRepeat(parent, runHere, childDepth);
                        bool childTrunc = truncated;
                        string childReason = childTrunc ? truncatedReason : null;
                        truncated = prevTrunc;
                        truncatedReason = prevReason;
                        if (missed || childTrunc)
                        {
                            deferTrunc = true;
                            deferReason = childReason;
                        }

                        ci += runHere.Count - 1;
                        continue;
                    }

                    Walk(parent.ChildIds[ci], childDepth, childInherit, childInCard);
                }

                if (deferTrunc)
                {
                    truncated = true;
                    if (!string.IsNullOrEmpty(deferReason))
                    {
                        truncatedReason = deferReason;
                    }
                    else if (string.IsNullOrEmpty(truncatedReason))
                    {
                        truncatedReason = "ref_nodes>" + MaxRefNodes;
                    }
                }
            }

            void Walk(string nodeId, int depth, string inheritName, bool inCardGroup)
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
                bool hoverReveal = HoverRevealRoles.Contains(peekRole)
                    || IsDetailRevealName(node.Name);
                bool ignoredImageInCard = inCardGroup && IsImageRole(peekRole);
                if (node.Ignored && !IsRootLike(node.Role) && !FrameRoles.Contains(peekRole)
                    && !hoverReveal && !ignoredImageInCard)
                {
                    string ignName = SanitizeAccessibleName(node.Name ?? "");
                    if (LooksLikeCardGroup(ignName, node, byId))
                    {
                        AppendLine(sb, depth, peekRole, ignName, null, note: null);
                        VisitChildren(node, depth + 1, ignName, true);
                        return;
                    }

                    VisitChildren(node, depth, inheritName, inCardGroup);
                    return;
                }

                bool inputLike = IsInputLike(node);

                string role = peekRole;
                string name = SanitizeAccessibleName(ResolveFrameName(node, role));
                if (string.IsNullOrWhiteSpace(name)
                    && string.Equals(role, "link", StringComparison.OrdinalIgnoreCase)
                    && !string.IsNullOrWhiteSpace(inheritName)
                    && HasImageDescendant(node, byId))
                {
                    name = inheritName;
                }

                if (string.IsNullOrWhiteSpace(name)
                    && inCardGroup
                    && IsImageRole(role)
                    && !string.IsNullOrWhiteSpace(inheritName))
                {
                    name = inheritName;
                }

                node.Name = name;

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

                bool namedGroup = LooksLikeCardGroup(name, node, byId) && IsNoiseRole(role);
                bool cardPic = inCardGroup && IsImageRole(role);
                bool cardText = inCardGroup
                    && !namedGroup
                    && !string.IsNullOrWhiteSpace(name)
                    && name.Trim().Length >= 2
                    && name.Trim().Length <= 40
                    && (string.Equals(role, "StaticText", StringComparison.OrdinalIgnoreCase)
                        || string.Equals(role, "text", StringComparison.OrdinalIgnoreCase));
                bool reserve = pageHasRuns && chromeLeft > 0;
                bool wantRef = (ShouldAssignRef(node, reserve ? 0 : refs.Count) || cardPic || cardText) && !namedGroup;
                string refId = null;
                if (wantRef && node.BackendDomNodeId.HasValue)
                {
                    bool chrome = IsChromeRole(role) && !analysis.IsUnderRun(node.NodeId);
                    refId = TakeRef(chrome);
                    if (refId != null)
                    {
                        refs[refId] = new BrowserRefEntry
                        {
                            AxNodeId = node.NodeId,
                            BackendDomNodeId = node.BackendDomNodeId,
                            Role = role,
                            Name = name
                        };
                    }
                    else if (!reserve)
                    {
                        truncated = true;
                        truncatedReason = truncatedReason ?? ("ref_nodes>" + MaxRefNodes);
                    }
                }
                else if (wantRef && !node.BackendDomNodeId.HasValue)
                {
                    // 无可操作 backend id：仍打印，不发 ref
                    wantRef = false;
                }

                bool print = wantRef
                    || namedGroup
                    || cardPic
                    || inputLike
                    || StructureRoles.Contains(role)
                    || (!string.IsNullOrWhiteSpace(name) && !IsNoiseRole(role))
                    || depth == 0;

                if (print)
                {
                    string note = (hoverReveal && node.Ignored) ? "隐藏" : null;
                    string value = ShouldPrintValue(node) ? (node.Value ?? "") : null;
                    AppendLine(sb, depth, role, name, refId, note, value);
                    if (sb.Length >= MaxTreeChars)
                    {
                        truncated = true;
                        truncatedReason = "tree_chars>" + MaxTreeChars;
                        return;
                    }
                }

                int childDepth = print ? depth + 1 : depth;
                string childInherit = namedGroup && !string.IsNullOrWhiteSpace(name)
                    ? name
                    : inheritName;
                bool childInCard = inCardGroup || namedGroup;
                VisitChildren(node, childDepth, childInherit, childInCard);
            }

            Walk(startId, 0, null, false);

            if (windowInvalid || (window != null && !windowApplied))
            {
                result.Error = "ref 无效或已过期，请重新 F_browser_snapshot";
                result.Success = false;
                return result;
            }

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
                string name = SanitizeAccessibleName(node.Name ?? "");
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
                string value = ShouldPrintValue(node) ? (node.Value ?? "") : null;
                AppendLine(sb, 1, role, name, refId, note: null, value);
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

            if (IsDetailRevealName(node.Name))
            {
                return true;
            }

            if (string.Equals(role, "heading", StringComparison.OrdinalIgnoreCase)
                && !string.IsNullOrWhiteSpace(node.Name))
            {
                return true;
            }

            if (RowRoles.Contains(role) && !string.IsNullOrWhiteSpace(node.Name))
            {
                return true;
            }

            return false;
        }

        private static bool LooksLikeCardGroup(string name, AxNode node, Dictionary<string, AxNode> byId)
        {
            if (string.IsNullOrWhiteSpace(name) || name.Trim().Length < 8 || node == null)
            {
                return false;
            }

            int links = CountRoleDescendants(node, byId, "link", maxCount: 9);
            return links >= 1 && links <= 8;
        }

        private static bool IsImageRole(string role)
        {
            return string.Equals(role, "image", StringComparison.OrdinalIgnoreCase)
                || string.Equals(role, "img", StringComparison.OrdinalIgnoreCase)
                || string.Equals(role, "Image", StringComparison.OrdinalIgnoreCase);
        }

        private static bool HasImageDescendant(AxNode node, Dictionary<string, AxNode> byId)
        {
            return CountRoleDescendants(node, byId, "image", maxCount: 1) > 0
                || CountRoleDescendants(node, byId, "img", maxCount: 1) > 0;
        }

        private static int CountRoleDescendants(
            AxNode node,
            Dictionary<string, AxNode> byId,
            string role,
            int maxCount)
        {
            if (node == null || byId == null || string.IsNullOrEmpty(role) || maxCount < 1)
            {
                return 0;
            }

            var stack = new Stack<string>();
            var seen = new HashSet<string>(StringComparer.Ordinal);
            foreach (string id in node.ChildIds)
            {
                stack.Push(id);
            }

            int count = 0;
            int budget = 80;
            while (stack.Count > 0 && budget-- > 0 && count < maxCount)
            {
                string id = stack.Pop();
                if (string.IsNullOrEmpty(id) || !seen.Add(id) || !byId.TryGetValue(id, out AxNode child))
                {
                    continue;
                }

                if (string.Equals(child.Role ?? "", role, StringComparison.OrdinalIgnoreCase))
                {
                    count++;
                }

                foreach (string next in child.ChildIds)
                {
                    stack.Push(next);
                }
            }

            return count;
        }

        private static bool IsNoiseRole(string role)
        {
            return string.Equals(role, "InlineTextBox", StringComparison.OrdinalIgnoreCase)
                || string.Equals(role, "LineBreak", StringComparison.OrdinalIgnoreCase)
                || string.Equals(role, "generic", StringComparison.OrdinalIgnoreCase)
                || string.Equals(role, "Group", StringComparison.OrdinalIgnoreCase)
                || string.Equals(role, "none", StringComparison.OrdinalIgnoreCase)
                || string.Equals(role, "presentation", StringComparison.OrdinalIgnoreCase)
                || string.Equals(role, "LayoutTable", StringComparison.OrdinalIgnoreCase)
                || string.Equals(role, "LayoutTableCell", StringComparison.OrdinalIgnoreCase);
        }

        internal static bool IsDetailRevealName(string name)
        {
            if (string.IsNullOrWhiteSpace(name))
            {
                return false;
            }

            string n = name.Trim();
            if (n.IndexOf("查看详情", StringComparison.Ordinal) >= 0)
            {
                return true;
            }

            if (n.IndexOf("查看更多", StringComparison.Ordinal) >= 0)
            {
                return true;
            }

            return string.Equals(n, "详情", StringComparison.Ordinal);
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

        internal static bool IsValidatorAccessibleName(string name)
        {
            if (string.IsNullOrWhiteSpace(name))
            {
                return false;
            }

            string s = name.Replace("\r", " ").Replace("\n", " ").Trim();
            if (string.Equals(s, "不能为空", StringComparison.Ordinal)
                || string.Equals(s, "必填", StringComparison.Ordinal)
                || string.Equals(s, "此项必填", StringComparison.Ordinal)
                || s.IndexOf("输入格式不正确", StringComparison.Ordinal) >= 0)
            {
                return true;
            }

            string low = s.ToLowerInvariant();
            return low == "required"
                || low == "this field is required"
                || low == "cannot be empty"
                || low == "must not be empty";
        }

        private static string SanitizeAccessibleName(string name)
        {
            return IsValidatorAccessibleName(name) ? "" : (name ?? "");
        }

        private static bool ShouldPrintValue(AxNode node)
        {
            if (node == null || !IsInputLike(node))
            {
                return false;
            }

            string n = node.Name ?? "";
            if (n.IndexOf("密码", StringComparison.Ordinal) >= 0
                || n.IndexOf("password", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return false;
            }

            return true;
        }

        private static void AppendLine(
            StringBuilder sb,
            int depth,
            string role,
            string name,
            string refId,
            string note,
            string value = null)
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

            if (value != null)
            {
                sb.Append(" value=\"").Append(Escape(value.Trim())).Append('"');
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
            public string Value { get; set; }
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
                    Value = ReadNamedAxProperty(jo["properties"], "value"),
                    UrlHint = ReadNamedAxProperty(jo["properties"], "url"),
                    Ignored = jo["ignored"] != null && jo["ignored"].Type == JTokenType.Boolean && (bool)jo["ignored"],
                    Editable = ReadEditable(jo["properties"]),
                    BackendDomNodeId = backend,
                    ChildIds = childIds
                };
            }

            private static string ReadNamedAxProperty(JToken propertiesToken, string propertyName)
            {
                if (!(propertiesToken is JArray props) || string.IsNullOrEmpty(propertyName))
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
                    if (!string.Equals(pname, propertyName, StringComparison.OrdinalIgnoreCase))
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

        /// <summary>带序号范围时，本次 repeat 行的新 ref。</summary>
        public string FocusRef { get; set; }

        public bool Truncated { get; set; }
        public string TruncatedReason { get; set; }
        public Dictionary<string, BrowserRefEntry> Refs { get; set; }
            = new Dictionary<string, BrowserRefEntry>(StringComparer.OrdinalIgnoreCase);
    }
}
