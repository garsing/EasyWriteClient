using System;
using WordAddIn1;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1.DocumentMapping.CodeResolve
{
    /// <summary>
    /// P_ → Word 段落 Range。转调 <see cref="ParagraphRangeLocator"/>，不 Find。
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

            if (!DocumentState.HasParagraphMapping(paragraphCode))
            {
                System.Diagnostics.Debug.WriteLine(
                    $"[ParagraphLocator] 无 P_ mapping: {paragraphCode}");
                return null;
            }

            return ParagraphRangeLocator.TryResolve(doc, paragraphCode, debugTag);
        }
    }
}
