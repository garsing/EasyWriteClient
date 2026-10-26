using System;
using System.Collections.Generic;
using System.Text;

namespace WordAddIn1.BrowserHost
{
    internal static partial class BrowserAxTreeBuilder
    {
        internal const int RepeatMinCount = 3;
        internal const int RepeatMinDescendants = 12;
        internal const int RepeatWindowSize = 300;
        internal const int RepeatNameChars = 40;

        internal static string ValidateRepeatRange(
            BrowserRefEntry prior,
            int? rangeStart,
            int? rangeEnd,
            bool hasRef)
        {
            bool hasStart = rangeStart.HasValue;
            bool hasEnd = rangeEnd.HasValue;
            if (hasStart != hasEnd)
            {
                return "要两个一起传";
            }

            if (!hasStart)
            {
                if (prior != null && prior.IsRepeat)
                {
                    return "重复段请同时传入 range_start 和 range_end";
                }

                return null;
            }

            if (!hasRef || prior == null || !prior.IsRepeat)
            {
                return "范围只对重复段 ref 使用";
            }

            if (rangeStart.Value < 1 || rangeStart.Value > rangeEnd.Value)
            {
                return "序号范围无效";
            }

            return null;
        }

        internal sealed class RepeatWindowRequest
        {
            public int ParentBackendId { get; set; }
            public string Signature { get; set; }
            public int Ordinal { get; set; }
            public int Start { get; set; }
            public int End { get; set; }
        }

        private sealed class RepeatRun
        {
            public int StartIndex { get; set; }
            public int Count { get; set; }
            public int Ordinal { get; set; }
            public string Signature { get; set; }
            public List<string> MemberIds { get; set; }
        }

        private sealed class RepeatAnalysis
        {
            public Dictionary<string, List<RepeatRun>> RunsByParent { get; set; }
            public HashSet<string> RunMembers { get; set; }
            public Dictionary<string, string> ParentOf { get; set; }
            public int ChromeSlots { get; set; }

            public bool IsUnderRun(string id)
            {
                var seen = new HashSet<string>(StringComparer.Ordinal);
                while (!string.IsNullOrEmpty(id) && seen.Add(id))
                {
                    if (RunMembers.Contains(id))
                    {
                        return true;
                    }

                    if (!ParentOf.TryGetValue(id, out id))
                    {
                        return false;
                    }
                }

                return false;
            }
        }

        private sealed class NodeSig
        {
            public string Signature { get; set; }
            public List<string> Parts { get; set; }
            public int Descendants { get; set; }
        }

        private static bool IsRepeatNoise(string role)
        {
            return string.Equals(role, "generic", StringComparison.OrdinalIgnoreCase)
                || string.Equals(role, "InlineTextBox", StringComparison.OrdinalIgnoreCase)
                || string.Equals(role, "LineBreak", StringComparison.OrdinalIgnoreCase);
        }

        private static bool IsChromeRole(string role)
        {
            return string.Equals(role, "button", StringComparison.OrdinalIgnoreCase)
                || string.Equals(role, "link", StringComparison.OrdinalIgnoreCase)
                || string.Equals(role, "textbox", StringComparison.OrdinalIgnoreCase)
                || string.Equals(role, "checkbox", StringComparison.OrdinalIgnoreCase)
                || string.Equals(role, "radio", StringComparison.OrdinalIgnoreCase)
                || string.Equals(role, "combobox", StringComparison.OrdinalIgnoreCase)
                || string.Equals(role, "searchbox", StringComparison.OrdinalIgnoreCase)
                || string.Equals(role, "text field", StringComparison.OrdinalIgnoreCase)
                || string.Equals(role, "search box", StringComparison.OrdinalIgnoreCase)
                || string.Equals(role, "combo box", StringComparison.OrdinalIgnoreCase);
        }

        private static RepeatAnalysis AnalyzeRepeats(Dictionary<string, AxNode> byId)
        {
            var cache = new Dictionary<string, NodeSig>(StringComparer.Ordinal);
            var stack = new HashSet<string>(StringComparer.Ordinal);
            foreach (AxNode node in byId.Values)
            {
                ComputeSig(node, byId, cache, stack);
            }

            var parentOf = new Dictionary<string, string>(StringComparer.Ordinal);
            foreach (AxNode node in byId.Values)
            {
                foreach (string child in node.ChildIds)
                {
                    if (!parentOf.ContainsKey(child))
                    {
                        parentOf[child] = node.NodeId;
                    }
                }
            }

            var runsByParent = new Dictionary<string, List<RepeatRun>>(StringComparer.Ordinal);
            var runMembers = new HashSet<string>(StringComparer.Ordinal);
            int runHeaders = 0;
            foreach (AxNode parent in byId.Values)
            {
                var runs = new List<RepeatRun>();
                int ordinal = 0;
                int i = 0;
                List<string> kids = parent.ChildIds;
                while (i < kids.Count)
                {
                    if (!cache.TryGetValue(kids[i], out NodeSig sig) || sig.Descendants < RepeatMinDescendants)
                    {
                        i++;
                        continue;
                    }

                    int j = i + 1;
                    while (j < kids.Count
                        && cache.TryGetValue(kids[j], out NodeSig next)
                        && next.Descendants >= RepeatMinDescendants
                        && string.Equals(next.Signature, sig.Signature, StringComparison.Ordinal))
                    {
                        j++;
                    }

                    int count = j - i;
                    if (count >= RepeatMinCount)
                    {
                        var members = new List<string>(count);
                        for (int k = i; k < j; k++)
                        {
                            members.Add(kids[k]);
                            runMembers.Add(kids[k]);
                        }

                        runs.Add(new RepeatRun
                        {
                            StartIndex = i,
                            Count = count,
                            Ordinal = ordinal,
                            Signature = sig.Signature,
                            MemberIds = members
                        });
                        ordinal++;
                        runHeaders++;
                        i = j;
                    }
                    else
                    {
                        i++;
                    }
                }

                if (runs.Count > 0)
                {
                    runsByParent[parent.NodeId] = runs;
                }
            }

            int chrome = runHeaders;
            foreach (AxNode node in byId.Values)
            {
                if (!node.BackendDomNodeId.HasValue)
                {
                    continue;
                }

                string role = node.Role ?? "";
                if (!IsChromeRole(role))
                {
                    continue;
                }

                if (IsUnder(node.NodeId, runMembers, parentOf))
                {
                    continue;
                }

                bool hover = HoverRevealRoles.Contains(role) || IsDetailRevealName(node.Name);
                if (node.Ignored && !hover && !IsRootLike(role))
                {
                    continue;
                }

                chrome++;
            }

            return new RepeatAnalysis
            {
                RunsByParent = runsByParent,
                RunMembers = runMembers,
                ParentOf = parentOf,
                ChromeSlots = chrome
            };
        }

        private static bool IsUnder(
            string id,
            HashSet<string> runMembers,
            Dictionary<string, string> parentOf)
        {
            var seen = new HashSet<string>(StringComparer.Ordinal);
            while (!string.IsNullOrEmpty(id) && seen.Add(id))
            {
                if (runMembers.Contains(id))
                {
                    return true;
                }

                if (!parentOf.TryGetValue(id, out id))
                {
                    return false;
                }
            }

            return false;
        }

        private static NodeSig ComputeSig(
            AxNode node,
            Dictionary<string, AxNode> byId,
            Dictionary<string, NodeSig> cache,
            HashSet<string> stack)
        {
            if (node == null)
            {
                return new NodeSig { Signature = "generic()", Parts = new List<string>(), Descendants = 1 };
            }

            if (cache.TryGetValue(node.NodeId, out NodeSig hit))
            {
                return hit;
            }

            if (!stack.Add(node.NodeId))
            {
                return new NodeSig
                {
                    Signature = RoleOf(node) + "()",
                    Parts = new List<string>(),
                    Descendants = 1
                };
            }

            int desc = 1;
            var raw = new List<string>();
            foreach (string cid in node.ChildIds)
            {
                if (!byId.TryGetValue(cid, out AxNode child))
                {
                    continue;
                }

                NodeSig cs = ComputeSig(child, byId, cache, stack);
                desc += cs.Descendants;
                if (IsRepeatNoise(child.Role))
                {
                    if (cs.Parts != null)
                    {
                        raw.AddRange(cs.Parts);
                    }
                }
                else
                {
                    raw.Add(cs.Signature);
                }
            }

            stack.Remove(node.NodeId);
            List<string> parts = CollapseParts(raw);
            var sig = new NodeSig
            {
                Signature = RoleOf(node) + "(" + string.Join(",", parts) + ")",
                Parts = parts,
                Descendants = desc
            };
            cache[node.NodeId] = sig;
            return sig;
        }

        private static string RoleOf(AxNode node)
        {
            return string.IsNullOrWhiteSpace(node?.Role) ? "generic" : node.Role;
        }

        private static List<string> CollapseParts(List<string> parts)
        {
            var rle = new List<string>();
            foreach (string part in parts)
            {
                if (rle.Count == 0)
                {
                    rle.Add(part);
                    continue;
                }

                string last = rle[rle.Count - 1];
                if (string.Equals(StripTimes(last), StripTimes(part), StringComparison.Ordinal))
                {
                    rle[rle.Count - 1] = StripTimes(part) + "×";
                }
                else
                {
                    rle.Add(part);
                }
            }

            return CollapsePeriodic(rle);
        }

        private static string StripTimes(string value)
        {
            if (!string.IsNullOrEmpty(value) && value.EndsWith("×", StringComparison.Ordinal))
            {
                return value.Substring(0, value.Length - 1);
            }

            return value ?? "";
        }

        private static List<string> CollapsePeriodic(List<string> seq)
        {
            var result = new List<string>();
            int i = 0;
            while (i < seq.Count)
            {
                int bestP = 0;
                int bestK = 0;
                int maxP = (seq.Count - i) / 2;
                for (int p = 1; p <= maxP; p++)
                {
                    int k = 1;
                    while (i + (k + 1) * p <= seq.Count && BlockEquals(seq, i, i + k * p, p))
                    {
                        k++;
                    }

                    if (k >= 2 && k * p > bestK * bestP)
                    {
                        bestP = p;
                        bestK = k;
                    }
                }

                if (bestK >= 2)
                {
                    var block = new StringBuilder();
                    for (int j = 0; j < bestP; j++)
                    {
                        if (j > 0)
                        {
                            block.Append(',');
                        }

                        block.Append(seq[i + j]);
                    }

                    result.Add("(" + block + ")×");
                    i += bestK * bestP;
                }
                else
                {
                    result.Add(seq[i]);
                    i++;
                }
            }

            return result;
        }

        private static bool BlockEquals(List<string> seq, int a, int b, int len)
        {
            for (int j = 0; j < len; j++)
            {
                if (!string.Equals(seq[a + j], seq[b + j], StringComparison.Ordinal))
                {
                    return false;
                }
            }

            return true;
        }
    }
}
