using System;
using System.Collections.Generic;
using System.Linq;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1.DocumentMapping.CodeResolve
{
    /// <summary>
    /// L5（0b）：P_ 编码解析 — DocumentState 索引 + Word 段落 Range。
    /// </summary>
    public static class ParagraphCodeResolver
    {
        /// <summary>Find 失败（含空段 I4-A）。Step B+ 使用。</summary>
        public const string ErrorParagraphRangeNotFound = "paragraph_range_not_found";

        /// <summary>未 ProcessDocument / 无 P_ mapping。</summary>
        public const string ErrorParagraphMappingMissing = "paragraph_mapping_missing";

        /// <summary>重复 P_ 且 codes 无法唯一定位（与 S_ 同码 ambiguous_sentence_code）。</summary>
        public const string ErrorAmbiguousParagraphCode = DisplaySequenceLocator.ErrorAmbiguousSentenceCode;

        /// <summary>S_ → 所属 P_（仅查 DocumentState.SentenceToParagraph）。</summary>
        public static string ResolveParagraphFromSentence(string sentenceCode)
        {
            if (string.IsNullOrEmpty(sentenceCode))
            {
                return null;
            }

            return DocumentState.SentenceToParagraph.TryGetValue(sentenceCode, out string paragraphCode)
                ? paragraphCode
                : null;
        }

        /// <summary>P_ → 段内 S_ 有序列表；空段或无映射时返回空列表。</summary>
        public static IReadOnlyList<string> ExpandParagraphToSentenceCodes(string paragraphCode)
        {
            if (string.IsNullOrEmpty(paragraphCode))
            {
                return Array.Empty<string>();
            }

            if (!DocumentState.ParagraphToSentences.TryGetValue(paragraphCode, out List<string> sentences)
                || sentences == null
                || sentences.Count == 0)
            {
                return Array.Empty<string>();
            }

            return sentences.ToList().AsReadOnly();
        }

        /// <summary>P_ → Word 段落 Range（整段）。映射缺失或 Find 失败返回 null（I4-A）。</summary>
        public static Word.Range ResolveParagraphRange(
            Word.Document doc,
            string paragraphCode,
            string tableId = null,
            TableScopeIndex tableScope = null,
            bool allowAutoTableScope = false,
            string debugTag = null)
        {
            return ParagraphCodeLocator.LocateRange(
                doc,
                paragraphCode,
                tableId,
                tableScope,
                allowAutoTableScope,
                debugTag);
        }

        /// <summary>带 errorCode 的解析；Find / mapping 失败时不抛异常。</summary>
        public static bool TryResolveParagraphRange(
            Word.Document doc,
            string paragraphCode,
            out Word.Range paragraphRange,
            out string errorCode,
            out string errorMessage,
            string tableId = null,
            TableScopeIndex tableScope = null,
            bool allowAutoTableScope = false,
            string debugTag = null)
        {
            paragraphRange = null;
            errorCode = null;
            errorMessage = null;

            if (string.IsNullOrEmpty(paragraphCode))
            {
                errorCode = ErrorParagraphMappingMissing;
                errorMessage = "paragraphCode 为空";
                return false;
            }

            string storedText = DocumentState.GetParagraphContent(paragraphCode);
            if (string.IsNullOrEmpty(storedText))
            {
                errorCode = ErrorParagraphMappingMissing;
                errorMessage = $"未找到段落映射或 content 为空: {paragraphCode}";
                return false;
            }

            if (ParagraphCodeAmbiguityHelper.IsEmptyParagraphStoredText(storedText))
            {
                errorCode = ErrorParagraphRangeNotFound;
                errorMessage = $"空段 {paragraphCode} 无法定位 Word 段落 Range（I4-A）";
                return false;
            }

            if (doc == null)
            {
                errorCode = ErrorParagraphMappingMissing;
                errorMessage = "Word 文档实例不可用";
                return false;
            }

            paragraphRange = ParagraphCodeLocator.LocateRange(
                doc,
                paragraphCode,
                tableId,
                tableScope,
                allowAutoTableScope,
                debugTag);

            if (paragraphRange != null)
            {
                return true;
            }

            errorCode = ErrorParagraphRangeNotFound;
            errorMessage = $"无法定位段落 Range: {paragraphCode}";
            return false;
        }
    }
}
