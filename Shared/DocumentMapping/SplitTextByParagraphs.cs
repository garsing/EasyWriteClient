using System.Collections.Generic;
using System.Text;

namespace WordAddIn1.DocumentMapping
{
    /// <summary>
    /// L2：seg 内按 \r 切段落，片内再按段中 \f / ^b 切开（与 COM Paragraphs 对齐）。
    /// 目录标记和切句同一条：标记是切分点，不进段文。只剩标记的那一截丢掉。
    /// 标记不结束段落，段界仍是换行，所以标记前后的字仍是同一段。
    /// 标记后面的换行留在后文里，按换行收成空段。
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
                if (ch == WordAddIn1.TocFieldStamp.Edge)
                {
                    continue;
                }

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
