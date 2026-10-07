using System.Collections.Generic;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1.DocumentMapping.CodeResolve
{
    /// <summary>
    /// 在统一栈得到的段落序列内，按 P_ 落到单段 Range（字符序，不 Find）。
    /// </summary>
    public static class SpanParagraphLocator
    {
        public static Word.Range LocateParagraph(
            Word.Document doc,
            Word.Range span,
            string paragraphCode,
            string debugTag = null)
        {
            if (doc == null || string.IsNullOrEmpty(paragraphCode))
            {
                return null;
            }

            return ParagraphRangeLocator.TryResolve(doc, paragraphCode, debugTag);
        }

        public static Word.Range LocateFirstParagraph(
            Word.Document doc,
            Word.Range span,
            IReadOnlyList<string> codeList,
            string debugTag = null)
        {
            if (codeList == null || codeList.Count == 0)
            {
                return null;
            }

            if (codeList.Count == 1)
            {
                return span;
            }

            return ParagraphRangeLocator.TryResolve(doc, codeList[0], debugTag);
        }

        public static Word.Range LocateLastParagraph(
            Word.Document doc,
            Word.Range span,
            IReadOnlyList<string> codeList,
            string debugTag = null)
        {
            if (codeList == null || codeList.Count == 0)
            {
                return null;
            }

            if (codeList.Count == 1)
            {
                return span;
            }

            return ParagraphRangeLocator.TryResolve(doc, codeList[codeList.Count - 1], debugTag);
        }
    }
}
