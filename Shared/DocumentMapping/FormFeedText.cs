using System;

namespace WordAddIn1.DocumentMapping
{
    /// <summary>
    /// w:t 里的字符 12 写成 &lt;c12/&gt;；真分页仍是 \f。
    /// </summary>
    public static class FormFeedText
    {
        public const string FakeTag = "<c12/>";

        public static string EscapeInXmlText(string xmlEscapedRun)
        {
            if (string.IsNullOrEmpty(xmlEscapedRun))
            {
                return xmlEscapedRun;
            }

            return xmlEscapedRun.Replace("\f", FakeTag);
        }

        public static string UnescapeToChar12(string text)
        {
            if (string.IsNullOrEmpty(text)
                || text.IndexOf(FakeTag, StringComparison.Ordinal) < 0)
            {
                return text ?? "";
            }

            return text.Replace(FakeTag, "\f");
        }

        /// <summary>
        /// 片内 \f 不在末尾则在其后切开，\f 留前片。末片沿用原 trailingCr。
        /// </summary>
        public static void AppendSplitAfterMidFormFeed(
            System.Collections.Generic.List<ParagraphBlock> blocks,
            ref int blockIndex,
            string part,
            bool trailingCr)
        {
            if (part == null)
            {
                part = "";
            }

            int from = 0;
            for (int i = 0; i < part.Length; i++)
            {
                if (part[i] != '\f' || i + 1 >= part.Length)
                {
                    continue;
                }

                string left = part.Substring(from, i - from + 1);
                blocks.Add(new ParagraphBlock(blockIndex, left, false));
                blockIndex++;
                from = i + 1;
            }

            string rest = from == 0 ? part : part.Substring(from);
            if (rest.Length == 0 && !trailingCr && from > 0)
            {
                return;
            }

            blocks.Add(new ParagraphBlock(blockIndex, rest, trailingCr));
            blockIndex++;
        }
    }
}
