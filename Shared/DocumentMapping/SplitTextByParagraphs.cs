using System.Collections.Generic;
using System.Text;

namespace WordAddIn1.DocumentMapping
{
    /// <summary>
    /// L2：seg 内按 \r 切段落，片内再按段中 \f / ^b 切开（与 COM Paragraphs 对齐）。
    /// 目录标记在切 seg 时已经拿掉，不会进到这里。段界仍是换行。
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

            var part = new StringBuilder();
            int blockIndex = 0;
            for (int i = 0; i < segText.Length; i++)
            {
                char ch = segText[i];
                if (ch == '\r')
                {
                    FormFeedText.AppendSplitAfterMidFormFeed(
                        blocks, ref blockIndex, part.ToString(), true);
                    part.Clear();
                    continue;
                }

                part.Append(ch);
            }

            if (part.Length > 0)
            {
                FormFeedText.AppendSplitAfterMidFormFeed(
                    blocks, ref blockIndex, part.ToString(), false);
            }

            return blocks;
        }
    }
}
