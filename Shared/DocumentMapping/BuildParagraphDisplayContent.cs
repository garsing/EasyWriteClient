using System.Collections.Generic;

namespace WordAddIn1.DocumentMapping
{
    /// <summary>
    /// L4：paragraph display（仅 [P_*] 标记）。撞码写 P_xxxxx_N，标记空格 start/end。
    /// </summary>
    public static class BuildParagraphDisplayContent
    {
        /// <summary>
        /// 与 <see cref="Build"/> 同一次公开码规则：撞码写成 P_xxxxx_N。
        /// </summary>
        public static List<List<string>> BuildPublicSegs(
            IReadOnlyList<IReadOnlyList<string>> segsParagraphNames)
        {
            var result = new List<List<string>>();
            if (segsParagraphNames == null)
            {
                return result;
            }

            Dictionary<string, int> baseCounts = CountBases(segsParagraphNames);
            var nextOrdinal = new Dictionary<string, int>();
            foreach (IReadOnlyList<string> seg in segsParagraphNames)
            {
                var published = new List<string>();
                if (seg != null)
                {
                    foreach (string paragraphName in seg)
                    {
                        published.Add(ToPublicCode(paragraphName, baseCounts, nextOrdinal));
                    }
                }

                result.Add(published);
            }

            return result;
        }

        public static void JoinSnapshots(
            IEnumerable<IReadOnlyList<string>> publicSegs,
            out string paragraphSnapshot,
            out string paragraphSegSnapshot)
        {
            var flat = new List<string>();
            var segParts = new List<string>();
            if (publicSegs != null)
            {
                foreach (IReadOnlyList<string> seg in publicSegs)
                {
                    if (seg == null)
                    {
                        segParts.Add("");
                        continue;
                    }

                    flat.AddRange(seg);
                    segParts.Add(string.Join(",", seg));
                }
            }

            paragraphSnapshot = string.Join(",", flat);
            paragraphSegSnapshot = string.Join(";", segParts);
        }

        public static string Build(
            IReadOnlyList<string> nonSegs,
            IReadOnlyList<IReadOnlyList<string>> segsParagraphNames,
            IReadOnlyDictionary<string, string> nameToParagraphContent)
        {
            var displayParts = new List<string>();
            List<List<string>> publicSegs = BuildPublicSegs(segsParagraphNames);
            nonSegs = nonSegs ?? new List<string>();

            for (int segIdx = 0; segIdx < publicSegs.Count; segIdx++)
            {
                if (segIdx < nonSegs.Count)
                {
                    displayParts.Add(nonSegs[segIdx] ?? "");
                }

                IReadOnlyList<string> baseSeg = segsParagraphNames != null && segIdx < segsParagraphNames.Count
                    ? segsParagraphNames[segIdx]
                    : null;
                var segDisplayParts = new List<string>();
                for (int i = 0; i < publicSegs[segIdx].Count; i++)
                {
                    string paragraphName = baseSeg != null && i < baseSeg.Count ? baseSeg[i] : "";
                    string paragraphContent = "";
                    if (nameToParagraphContent != null && paragraphName != null)
                    {
                        nameToParagraphContent.TryGetValue(paragraphName, out paragraphContent);
                    }

                    paragraphContent = paragraphContent ?? "";
                    string publicCode = publicSegs[segIdx][i];
                    segDisplayParts.Add($"[{publicCode} start]{paragraphContent}[{publicCode} end]");
                }

                displayParts.Add(string.Join("", segDisplayParts));
            }

            int sourceSegCount = segsParagraphNames == null ? 0 : segsParagraphNames.Count;
            if (nonSegs.Count > sourceSegCount)
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
