using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace WordAddIn1.BrowserHost
{
    /// <summary>
    /// 嵌套 iframe 收进同一次快照。总 ref 数到上限就停，停住的那一层仍留结构行。
    /// </summary>
    internal static class BrowserFrameInline
    {
        private const int MaxDepth = 8;

        public static async Task ExpandAsync(
            IBrowserCdp cdp,
            BrowserAxBuildResult built,
            BrowserFrameMap map)
        {
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            await ExpandAsync(cdp, built, map, seen, 0).ConfigureAwait(true);
        }

        private static async Task ExpandAsync(
            IBrowserCdp cdp,
            BrowserAxBuildResult built,
            BrowserFrameMap map,
            HashSet<string> seen,
            int depth)
        {
            if (cdp == null || built == null || !built.Success || built.Refs == null || depth >= MaxDepth)
            {
                return;
            }

            if (built.Refs.Count >= BrowserAxTreeBuilder.MaxRefNodes
                || (built.TreeText ?? "").Length >= BrowserAxTreeBuilder.MaxTreeChars)
            {
                built.Truncated = true;
                if (string.IsNullOrEmpty(built.TruncatedReason))
                {
                    built.TruncatedReason = "ref_nodes>" + BrowserAxTreeBuilder.MaxRefNodes;
                }

                return;
            }

            List<string> frames = FrameRefsInOrder(built);
            foreach (string frameRef in frames)
            {
                if (built.Refs.Count >= BrowserAxTreeBuilder.MaxRefNodes
                    || (built.TreeText ?? "").Length >= BrowserAxTreeBuilder.MaxTreeChars)
                {
                    built.Truncated = true;
                    if (string.IsNullOrEmpty(built.TruncatedReason))
                    {
                        built.TruncatedReason = "ref_nodes>" + BrowserAxTreeBuilder.MaxRefNodes;
                    }

                    return;
                }

                if (!built.Refs.TryGetValue(frameRef, out BrowserRefEntry entry) || entry == null
                    || !BrowserRiskGuard.IsFrameRole(entry))
                {
                    continue;
                }

                string childId = entry.ChildFrameId;
                if (string.IsNullOrWhiteSpace(childId)
                    && entry.BackendDomNodeId.HasValue
                    && map != null)
                {
                    childId = map.ChildFrameIdForOwner(entry.BackendDomNodeId.Value);
                }

                if (string.IsNullOrWhiteSpace(childId) && entry.BackendDomNodeId.HasValue)
                {
                    childId = await BrowserFrameResolver
                        .TryDescribeChildFrameIdAsync(cdp, entry.BackendDomNodeId.Value, entry.CdpSessionId)
                        .ConfigureAwait(true);
                }

                if (string.IsNullOrWhiteSpace(childId) || !seen.Add(childId))
                {
                    continue;
                }

                if (string.Equals(childId, entry.FrameId, StringComparison.Ordinal))
                {
                    continue;
                }

                int remaining = BrowserAxTreeBuilder.MaxRefNodes - built.Refs.Count;
                if (remaining <= 0)
                {
                    return;
                }

                BrowserFrameAxResult ax = await BrowserFrameResolver
                    .GetFrameAxTreeAsync(cdp, childId, BrowserAxTreeBuilder.DefaultDepth, entry.CdpSessionId)
                    .ConfigureAwait(true);
                if (ax == null || string.IsNullOrWhiteSpace(ax.Json))
                {
                    continue;
                }

                int? modalRoot = await BrowserHtmlModalProbe
                    .TryGetBackendNodeIdAsync(cdp, ax.SessionId)
                    .ConfigureAwait(true);
                BrowserAxBuildResult child = BrowserAxTreeBuilder.BuildOverview(ax.Json, modalRoot, remaining);
                if (!child.Success || string.IsNullOrWhiteSpace(child.TreeText))
                {
                    continue;
                }

                if (BrowserFrameResolver.ParentShellReason(child.TreeText, cdp.CurrentTitle) != null)
                {
                    continue;
                }

                BrowserFrameResolver.Stamp(child, childId, map, ax.SessionId);
                await BrowserDomClickableSupplement
                    .MergeAsync(cdp, child, childId, ax.SessionId)
                    .ConfigureAwait(true);
                await BrowserDomClickableSupplement
                    .PruneHiddenRowsAsync(cdp, child, ax.SessionId)
                    .ConfigureAwait(true);
                await BrowserAccessibleNameFill
                    .FillEmptyAsync(cdp, child, ax.SessionId)
                    .ConfigureAwait(true);
                await BrowserAccessibleNameFill
                    .FillSelectHintsAsync(cdp, child, ax.SessionId)
                    .ConfigureAwait(true);
                await BrowserAccessibleNameFill
                    .FillTemporalValuesAsync(cdp, child, ax.SessionId)
                    .ConfigureAwait(true);
                TrimTo(child, remaining);
                await ExpandAsync(cdp, child, map, seen, depth + 1).ConfigureAwait(true);
                TrimTo(child, remaining);

                if (built.TreeText.Length + child.TreeText.Length > BrowserAxTreeBuilder.MaxTreeChars)
                {
                    built.Truncated = true;
                    built.TruncatedReason = built.TruncatedReason
                        ?? ("tree_chars>" + BrowserAxTreeBuilder.MaxTreeChars);
                    return;
                }

                Splice(built, child, frameRef);
                if (child.Truncated)
                {
                    built.Truncated = true;
                    built.TruncatedReason = built.TruncatedReason
                        ?? child.TruncatedReason
                        ?? ("ref_nodes>" + BrowserAxTreeBuilder.MaxRefNodes);
                }
            }
        }

        private static List<string> FrameRefsInOrder(BrowserAxBuildResult built)
        {
            var list = new List<string>();
            if (string.IsNullOrEmpty(built.TreeText))
            {
                return list;
            }

            foreach (string line in SplitLines(built.TreeText))
            {
                string id = RefOnLine(line);
                if (id == null || !built.Refs.TryGetValue(id, out BrowserRefEntry entry) || entry == null)
                {
                    continue;
                }

                if (BrowserRiskGuard.IsFrameRole(entry))
                {
                    list.Add(id);
                }
            }

            return list;
        }

        private static void TrimTo(BrowserAxBuildResult built, int budget)
        {
            if (built?.Refs == null || budget < 0)
            {
                return;
            }

            while (built.Refs.Count > budget)
            {
                string id = built.Refs.Keys
                    .OrderByDescending(ParseNum)
                    .FirstOrDefault();
                if (string.IsNullOrEmpty(id))
                {
                    break;
                }

                built.Refs.Remove(id);
                built.TreeText = RemoveRefLine(built.TreeText, id);
                built.Truncated = true;
                built.TruncatedReason = "ref_nodes>" + BrowserAxTreeBuilder.MaxRefNodes;
            }
        }

        private static void Splice(BrowserAxBuildResult parent, BrowserAxBuildResult child, string frameRef)
        {
            var map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            int next = MaxNum(parent.Refs) + 1;
            foreach (string old in child.Refs.Keys.OrderBy(ParseNum).ToList())
            {
                map[old] = "e" + next.ToString(CultureInfo.InvariantCulture);
                next++;
            }

            string childText = child.TreeText ?? "";
            foreach (string old in map.Keys.OrderByDescending(id => id.Length).ThenByDescending(ParseNum))
            {
                childText = ReplaceRef(childText, old, map[old]);
            }

            List<string> parentLines = SplitLines(parent.TreeText);
            int at = -1;
            int indent = 0;
            for (int i = 0; i < parentLines.Count; i++)
            {
                if (!LineHasRef(parentLines[i], frameRef))
                {
                    continue;
                }

                at = i;
                indent = LeadingSpaces(parentLines[i]);
                break;
            }

            if (at < 0)
            {
                return;
            }

            var insert = new List<string>();
            bool first = true;
            foreach (string raw in SplitLines(childText))
            {
                if (string.IsNullOrWhiteSpace(raw))
                {
                    continue;
                }

                string body = raw.TrimStart();
                if (first
                    && (body.StartsWith("- RootWebArea", StringComparison.OrdinalIgnoreCase)
                        || body.StartsWith("- WebArea", StringComparison.OrdinalIgnoreCase)))
                {
                    first = false;
                    continue;
                }

                first = false;
                int rel = LeadingSpaces(raw);
                insert.Add(new string(' ', indent + rel) + body);
            }

            if (!child.Truncated && insert.Count > 0)
            {
                parentLines.RemoveAt(at);
                parentLines.InsertRange(at, insert);
                parent.Refs.Remove(frameRef);
            }
            else if (insert.Count > 0)
            {
                parentLines.InsertRange(at + 1, insert);
            }
            else
            {
                return;
            }

            parent.TreeText = string.Join(Environment.NewLine, parentLines);
            foreach (KeyValuePair<string, string> pair in map)
            {
                if (child.Refs.TryGetValue(pair.Key, out BrowserRefEntry entry) && entry != null)
                {
                    parent.Refs[pair.Value] = entry;
                }
            }
        }

        private static int MaxNum(Dictionary<string, BrowserRefEntry> refs)
        {
            int max = 0;
            if (refs == null)
            {
                return max;
            }

            foreach (string id in refs.Keys)
            {
                int n = ParseNum(id);
                if (n > max)
                {
                    max = n;
                }
            }

            return max;
        }

        private static int ParseNum(string refId)
        {
            if (string.IsNullOrEmpty(refId) || (refId[0] != 'e' && refId[0] != 'E'))
            {
                return 0;
            }

            return int.TryParse(
                refId.Substring(1),
                NumberStyles.Integer,
                CultureInfo.InvariantCulture,
                out int n)
                ? n
                : 0;
        }

        private static List<string> SplitLines(string text)
        {
            if (string.IsNullOrEmpty(text))
            {
                return new List<string>();
            }

            return new List<string>(Regex.Split(text, @"\r\n|\n|\r"));
        }

        private static int LeadingSpaces(string line)
        {
            int n = 0;
            if (line == null)
            {
                return n;
            }

            while (n < line.Length && line[n] == ' ')
            {
                n++;
            }

            return n;
        }

        private static string RefOnLine(string line)
        {
            if (string.IsNullOrEmpty(line))
            {
                return null;
            }

            Match match = Regex.Match(line, @"\bref=(e\d+)\b", RegexOptions.IgnoreCase);
            return match.Success ? match.Groups[1].Value : null;
        }

        private static bool LineHasRef(string line, string refId)
        {
            if (string.IsNullOrEmpty(line) || string.IsNullOrEmpty(refId))
            {
                return false;
            }

            return Regex.IsMatch(
                line,
                @"\bref=" + Regex.Escape(refId) + @"\b",
                RegexOptions.IgnoreCase);
        }

        private static string ReplaceRef(string text, string oldId, string newId)
        {
            if (string.IsNullOrEmpty(text))
            {
                return text ?? "";
            }

            return Regex.Replace(
                text,
                @"\bref=" + Regex.Escape(oldId) + @"\b",
                "ref=" + newId,
                RegexOptions.IgnoreCase);
        }

        private static string RemoveRefLine(string text, string refId)
        {
            List<string> lines = SplitLines(text);
            lines.RemoveAll(line => LineHasRef(line, refId));
            return string.Join(Environment.NewLine, lines);
        }
    }
}
