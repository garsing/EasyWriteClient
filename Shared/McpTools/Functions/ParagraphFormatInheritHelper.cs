using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    /// <summary>
    /// 批次 2 I7：inherit 路径 Word 级段落属性复制（不改 run char）。
    /// </summary>
    public static class ParagraphFormatInheritHelper
    {
        public static void CopyParagraphFormat(Word.Range sourceParagraphRange, Word.Range targetParagraphRange)
        {
            if (sourceParagraphRange == null || targetParagraphRange == null)
            {
                return;
            }

            Word.ParagraphFormat src = sourceParagraphRange.ParagraphFormat;
            Word.ParagraphFormat dst = targetParagraphRange.ParagraphFormat;

            try
            {
                var style = sourceParagraphRange.get_Style() as Word.Style;
                if (style != null && !string.IsNullOrEmpty(style.NameLocal))
                {
                    targetParagraphRange.set_Style(style.NameLocal);
                }
            }
            catch
            {
                // 样式复制失败时仍尝试 direct 属性
            }

            dst.Alignment = src.Alignment;
            dst.LeftIndent = src.LeftIndent;
            dst.RightIndent = src.RightIndent;
            dst.FirstLineIndent = src.FirstLineIndent;
            dst.LineSpacingRule = src.LineSpacingRule;
            dst.LineSpacing = src.LineSpacing;
            dst.SpaceBefore = src.SpaceBefore;
            dst.SpaceAfter = src.SpaceAfter;
        }
    }
}
