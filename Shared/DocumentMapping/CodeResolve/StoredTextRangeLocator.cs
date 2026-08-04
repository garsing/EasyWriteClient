using System;
using WordAddIn1;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1.DocumentMapping.CodeResolve
{
    /// <summary>
    /// 按 mapping storedText 在 Word 中 Find（全文或限表第一处命中）。S_/P_ 共用，与编码类型无关。
    /// </summary>
    public static class StoredTextRangeLocator
    {
        /// <summary>
        /// 无序号单次 Find：全文或限表第一处命中。供 ResolveSpan 拼串落 Span 等路径。
        /// </summary>
        public static Word.Range LocateFirst(
            Word.Document doc,
            string storedText,
            string tableId,
            string debugTag = null)
        {
            if (doc == null || string.IsNullOrEmpty(storedText))
            {
                return null;
            }

            string tag = debugTag ?? "";

            if (string.IsNullOrEmpty(tableId))
            {
                return WordRangeFinder.FindFirstInDocument(doc, storedText, tag);
            }

            Word.Table table = SentenceCodeLocator.ResolveTableById(doc, tableId);
            if (table == null)
            {
                System.Diagnostics.Debug.WriteLine(
                    $"[StoredTextRangeLocator] 未找到表格 {tableId}，无法 LocateFirst {tag}");
                return null;
            }

            System.Diagnostics.Debug.WriteLine(
                $"[StoredTextRangeLocator] 表内 LocateFirst tag={tag} table={tableId}");
            return WordRangeFinder.FindFirstInTable(
                table,
                storedText,
                string.IsNullOrEmpty(tag) ? $"table={tableId}" : tag);
        }

        /// <summary>
        /// 在已有 Span 内按 storedText 找第一处完全落在 Span 内的命中。
        /// </summary>
        public static Word.Range LocateInSpan(
            Word.Document doc,
            Word.Range span,
            string storedText,
            string debugTag = null)
        {
            if (doc == null || span == null || string.IsNullOrEmpty(storedText))
            {
                return null;
            }

            return WordRangeFinder.FindFirstInRange(span, storedText, debugTag ?? "");
        }
    }
}
