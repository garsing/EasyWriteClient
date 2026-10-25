using System.Collections.Generic;

namespace WordAddIn1.DocumentMapping
{
    /// <summary>
    /// L2：seg 内按 \r 切段落，片内再按段中 \f / ^b 切开（与 COM Paragraphs 对齐）。
    /// </summary>
    public static class SplitTextByParagraphs
    {
        public static List<ParagraphBlock> Split(string segText)
        {
            var blocks = new List<ParagraphBlock>();
            if (string.IsNullOrEmpty(segText))
            {
                return blocks;
            }

            string[] parts = segText.Split('\r');
            int blockIndex = 0;
            for (int i = 0; i < parts.Length; i++)
            {
                bool isLast = i == parts.Length - 1;
                bool trailingCr = !isLast;
                string part = parts[i];

                if (part == "" && !trailingCr)
                {
                    continue;
                }

                FormFeedText.AppendSplitAfterMidFormFeed(
                    blocks, ref blockIndex, part, trailingCr);
            }

            return blocks;
        }
    }
}
