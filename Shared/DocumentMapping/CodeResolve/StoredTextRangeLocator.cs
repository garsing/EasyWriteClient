using System;
using WordAddIn1;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1.DocumentMapping.CodeResolve
{
    /// <summary>
    /// 按 mapping storedText 在 Word 中 Find（全文或限表 nth）。S_/P_ 共用，与编码类型无关。
    /// </summary>
    public static class StoredTextRangeLocator
    {
        public static Word.Range Locate(
            Word.Document doc,
            string storedText,
            int occurrenceIndex,
            string tableId,
            string debugTag = null)
        {
            if (doc == null || string.IsNullOrEmpty(storedText) || occurrenceIndex < 0)
            {
                return null;
            }

            string tag = debugTag ?? "";

            if (string.IsNullOrEmpty(tableId))
            {
                return WordRangeFinder.FindSentenceRangeInDocumentNth(doc, storedText, occurrenceIndex, tag);
            }

            Word.Table table = SentenceCodeLocator.ResolveTableById(doc, tableId);
            if (table == null)
            {
                System.Diagnostics.Debug.WriteLine(
                    $"[StoredTextRangeLocator] 未找到表格 {tableId}，无法定位 {tag}");
                return null;
            }

            System.Diagnostics.Debug.WriteLine(
                $"[StoredTextRangeLocator] 表内 Find tag={tag} table={tableId} nth={occurrenceIndex} (使用Cell并集遍历)");
            return WordRangeFinder.FindSentenceRangeInTableNth(
                table,
                storedText,
                occurrenceIndex,
                string.IsNullOrEmpty(tag) ? $"table={tableId}" : tag);
        }
    }
}
