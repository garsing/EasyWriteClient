using WordAddIn1;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1.DocumentMapping.CodeResolve
{
    /// <summary>
    /// P_ → Word 段落 Range 定位（0b Step B+）。
    /// Find 走 StoredTextRangeLocator.LocateFirst；命中后再 ExpandToParagraphRange。
    /// </summary>
    public static class ParagraphCodeLocator
    {
        public static Word.Range LocateRange(
            Word.Document doc,
            string paragraphCode,
            string explicitTableId,
            TableScopeIndex tableScope,
            bool allowAutoTableScope = false,
            string debugTag = null)
        {
            if (doc == null || string.IsNullOrEmpty(paragraphCode))
            {
                return null;
            }

            string paragraphContent = DocumentState.GetParagraphContent(paragraphCode);
            if (string.IsNullOrEmpty(paragraphContent))
            {
                System.Diagnostics.Debug.WriteLine(
                    $"[ParagraphLocator] 无 P_ mapping 或 content 为空: {paragraphCode}");
                return null;
            }

            if (ParagraphCodeAmbiguityHelper.IsEmptyParagraphStoredText(paragraphContent))
            {
                System.Diagnostics.Debug.WriteLine(
                    $"[ParagraphLocator] 空段 {paragraphCode} 跳过 Find（I4-A）");
                return null;
            }

            string tableId = explicitTableId;
            if (string.IsNullOrEmpty(tableId) && allowAutoTableScope && tableScope != null)
            {
                tableId = TableParagraphScopeHelper.ResolveEffectiveTableId(
                    paragraphCode,
                    explicitTableId,
                    tableScope);
            }

            string tag = string.IsNullOrEmpty(debugTag) ? paragraphCode : debugTag;

            Word.Range hit = StoredTextRangeLocator.LocateFirst(
                doc,
                paragraphContent,
                tableId,
                tag);

            if (hit == null)
            {
                System.Diagnostics.Debug.WriteLine(
                    $"[ParagraphLocator] Find 未命中 P_={paragraphCode}");
                return null;
            }

            Word.Range paragraphRange = WordRangeFinder.ExpandToParagraphRange(hit);
            if (paragraphRange != null)
            {
                System.Diagnostics.Debug.WriteLine(
                    $"[ParagraphLocator] P_={paragraphCode} 段落 Start={paragraphRange.Start} End={paragraphRange.End}");
            }

            return paragraphRange;
        }
    }
}
