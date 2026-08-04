using System.Collections.Generic;
using System.Text;

namespace WordAddIn1.DocumentMapping
{
    /// <summary>
    /// L4：paragraph display（仅 [P_*] 标记）。
    /// </summary>
    public static class BuildParagraphDisplayContent
    {
        public static string Build(
            IReadOnlyList<string> nonSegs,
            IReadOnlyList<IReadOnlyList<string>> segsParagraphNames,
            IReadOnlyDictionary<string, string> nameToParagraphContent)
        {
            var displayParts = new List<string>();

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
                    segDisplayParts.Add($"[{paragraphName}-start]{paragraphContent}[{paragraphName}-end]");
                }

                displayParts.Add(string.Join("", segDisplayParts));
            }

            if (nonSegs.Count > segsParagraphNames.Count)
            {
                displayParts.Add(nonSegs[nonSegs.Count - 1] ?? "");
            }

            return string.Join("", displayParts);
        }
    }
}
