using System.Collections.Generic;
using WordAddIn1;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1.DocumentMapping.CodeResolve
{
    /// <summary>
    /// 在统一栈得到的序列 Span 内，按句编码定位单句 Range（无 occurrenceIndex）。
    /// 供 format_context / inherit 首句 / provenance 末句等路径复用。
    /// </summary>
    public static class SpanSentenceLocator
    {
        public static Word.Range LocateSentence(
            Word.Document doc,
            Word.Range span,
            string sentenceCode,
            string debugTag = null)
        {
            if (doc == null || span == null || string.IsNullOrEmpty(sentenceCode))
            {
                return null;
            }

            string storedText = DocumentState.GetSentenceContent(sentenceCode);
            if (string.IsNullOrEmpty(storedText))
            {
                return null;
            }

            string tag = string.IsNullOrEmpty(debugTag) ? sentenceCode : debugTag;
            return StoredTextRangeLocator.LocateInSpan(doc, span, storedText, tag);
        }

        public static Word.Range LocateFirstSentence(
            Word.Document doc,
            Word.Range span,
            IReadOnlyList<string> codeList,
            string debugTag = null)
        {
            if (codeList == null || codeList.Count == 0)
            {
                return null;
            }

            return LocateSentence(doc, span, codeList[0], debugTag);
        }

        public static Word.Range LocateLastSentence(
            Word.Document doc,
            Word.Range span,
            IReadOnlyList<string> codeList,
            string debugTag = null)
        {
            if (codeList == null || codeList.Count == 0)
            {
                return null;
            }

            return LocateSentence(doc, span, codeList[codeList.Count - 1], debugTag);
        }
    }
}
