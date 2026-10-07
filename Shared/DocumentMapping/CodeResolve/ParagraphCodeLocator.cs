using System;
using WordAddIn1;
using WordAddIn1.DocumentMapping;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1.DocumentMapping.CodeResolve
{
    /// <summary>
    /// P_ → Word 段落 Range。按 display 转字流后的字符序落到 Paragraphs[k]，不 Find。
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

            if (!HasParagraphMapping(paragraphCode))
            {
                System.Diagnostics.Debug.WriteLine(
                    $"[ParagraphLocator] 无 P_ mapping: {paragraphCode}");
                return null;
            }

            int paraNo = DisplayParagraphIndexMap.ResolveParaIndex(paragraphCode);
            if (paraNo <= 0)
            {
                System.Diagnostics.Debug.WriteLine(
                    $"[ParagraphLocator] 字符序未落到段号 P_={paragraphCode}");
                return null;
            }

            Word.Range paragraphRange;
            try
            {
                if (paraNo > doc.Paragraphs.Count)
                {
                    System.Diagnostics.Debug.WriteLine(
                        $"[ParagraphLocator] 段号越界 P_={paragraphCode} → {paraNo}/{doc.Paragraphs.Count}");
                    return null;
                }

                paragraphRange = doc.Paragraphs[paraNo].Range;
            }
            catch
            {
                return null;
            }

            string tag = string.IsNullOrEmpty(debugTag) ? paragraphCode : debugTag;
            System.Diagnostics.Debug.WriteLine(
                $"[ParagraphLocator] P_={tag} → 段[{paraNo}] Start={paragraphRange.Start} End={paragraphRange.End}");
            return paragraphRange;
        }

        private static bool HasParagraphMapping(string paragraphCode)
        {
            foreach (var row in DocumentState.ParagraphNameMapping)
            {
                if (row != null && row.Count >= 1
                    && string.Equals(row[0], paragraphCode, StringComparison.Ordinal))
                {
                    return true;
                }
            }

            return false;
        }
    }
}
