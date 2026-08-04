using System.Collections.Generic;

namespace WordAddIn1.DocumentMapping
{
    /// <summary>
    /// L4：sentence display（仅 [S_*] 标记），镜像 WordDocumentExtractor.BuildDisplayContent。
    /// </summary>
    internal static class BuildSentenceDisplayContent
    {
        public static string Build(
            IReadOnlyList<string> nonSegs,
            IReadOnlyList<IReadOnlyList<string>> segsSentenceNames,
            IReadOnlyDictionary<string, string> nameToContentMap)
        {
            var displayParts = new List<string>();

            for (int segIdx = 0; segIdx < segsSentenceNames.Count; segIdx++)
            {
                if (segIdx < nonSegs.Count)
                {
                    displayParts.Add(nonSegs[segIdx] ?? "");
                }

                var segDisplayParts = new List<string>();
                foreach (string sentenceName in segsSentenceNames[segIdx])
                {
                    nameToContentMap.TryGetValue(sentenceName, out string sentenceContent);
                    sentenceContent = sentenceContent ?? "";
                    segDisplayParts.Add($"[{sentenceName}-start]{sentenceContent}[{sentenceName}-end]");
                }

                displayParts.Add(string.Join("", segDisplayParts));
            }

            if (nonSegs.Count > segsSentenceNames.Count)
            {
                displayParts.Add(nonSegs[nonSegs.Count - 1] ?? "");
            }

            return string.Join("", displayParts);
        }
    }
}
