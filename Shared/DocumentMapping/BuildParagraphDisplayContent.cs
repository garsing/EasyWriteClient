using System.Collections.Generic;

namespace WordAddIn1.DocumentMapping
{
    /// <summary>
    /// L4：paragraph display（仅 [P_*] 标记）。撞码写 P_xxxxx_N，标记空格 start/end。
    /// </summary>
    public static class BuildParagraphDisplayContent
    {
        public static string Build(
            IReadOnlyList<string> nonSegs,
            IReadOnlyList<IReadOnlyList<string>> segsParagraphNames,
            IReadOnlyDictionary<string, string> nameToParagraphContent)
        {
            var displayParts = new List<string>();
            Dictionary<string, int> baseCounts = CountBases(segsParagraphNames);
            var nextOrdinal = new Dictionary<string, int>();

            for (int segIdx = 0; segIdx < segsParagraphNames.Count; segIdx++)
            {
                if (segIdx < nonSegs.Count)
                {
                    displayParts.Add(nonSegs[segIdx] ?? "");
                }

                var segDisplayParts = new List<string>();
                foreach (string paragraphName in segsParagraphNames[segIdx])
                {
                    nameToParagraphContent.TryGetValue(paragraphName, out string paragraphContent);
                    paragraphContent = paragraphContent ?? "";
                    string publicCode = ToPublicCode(paragraphName, baseCounts, nextOrdinal);
                    segDisplayParts.Add($"[{publicCode} start]{paragraphContent}[{publicCode} end]");
                }

                displayParts.Add(string.Join("", segDisplayParts));
            }

            if (nonSegs.Count > segsParagraphNames.Count)
            {
                displayParts.Add(nonSegs[nonSegs.Count - 1] ?? "");
            }

            return InTableBoundaryMarker.AnnotateParagraphs(string.Join("", displayParts));
        }

        private static Dictionary<string, int> CountBases(
            IReadOnlyList<IReadOnlyList<string>> segsParagraphNames)
        {
            var counts = new Dictionary<string, int>();
            if (segsParagraphNames == null)
            {
                return counts;
            }

            foreach (IReadOnlyList<string> seg in segsParagraphNames)
            {
                if (seg == null)
                {
                    continue;
                }

                foreach (string name in seg)
                {
                    if (string.IsNullOrEmpty(name))
                    {
                        continue;
                    }

                    if (!counts.ContainsKey(name))
                    {
                        counts[name] = 0;
                    }

                    counts[name]++;
                }
            }

            return counts;
        }

        private static string ToPublicCode(
            string baseCode,
            Dictionary<string, int> baseCounts,
            Dictionary<string, int> nextOrdinal)
        {
            if (string.IsNullOrEmpty(baseCode))
            {
                return baseCode;
            }

            int total;
            if (!baseCounts.TryGetValue(baseCode, out total) || total < 2)
            {
                return baseCode;
            }

            int n;
            if (!nextOrdinal.TryGetValue(baseCode, out n))
            {
                n = 0;
            }

            n++;
            nextOrdinal[baseCode] = n;
            return baseCode + "_" + n;
        }
    }
}
