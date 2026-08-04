using System.Collections.Generic;

namespace WordAddIn1.DocumentMapping
{
    /// <summary>
    /// L2：seg 内按 \r 切段落（与 Python split_text_by_paragraphs 对齐）。
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

                blocks.Add(new ParagraphBlock(blockIndex, part, trailingCr));
                blockIndex++;
            }

            return blocks;
        }
    }
}
