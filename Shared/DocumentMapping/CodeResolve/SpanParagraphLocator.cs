using System.Collections.Generic;
using WordAddIn1;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1.DocumentMapping.CodeResolve
{
    /// <summary>
    /// 在统一栈得到的段落序列 Span 内，按 P_ 编码定位单段 Range（无 occurrenceIndex）。
    /// </summary>
    public static class SpanParagraphLocator
    {
        public static Word.Range LocateParagraph(
            Word.Document doc,
            Word.Range span,
            string paragraphCode,
            string debugTag = null)
        {
            if (doc == null || span == null || string.IsNullOrEmpty(paragraphCode))
            {
                return null;
            }

            string storedText = DocumentState.GetParagraphContent(paragraphCode);
            if (string.IsNullOrEmpty(storedText))
            {
                return null;
            }

            string tag = string.IsNullOrEmpty(debugTag) ? paragraphCode : debugTag;
            return StoredTextRangeLocator.LocateInSpan(doc, span, storedText, tag);
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

            return LocateParagraph(doc, span, codeList[0], debugTag);
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

            return LocateParagraph(doc, span, codeList[codeList.Count - 1], debugTag);
        }
    }
}
