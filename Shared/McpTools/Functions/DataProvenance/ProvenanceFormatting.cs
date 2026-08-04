using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    public static class ProvenanceFormatting
    {
        public const Word.WdColor ReferenceColor = Word.WdColor.wdColorGray50;

        /// <summary>表/图/图片下与文末参考行统一字号（小四）；避免继承标题段大字。</summary>
        public const float ReferenceFontSize = 12f;

        /// <summary>独立参考说明统一字体（与正文默认 insert 一致）。</summary>
        public const string ReferenceFontName = "宋体";

        public static void ApplyReferenceStyle(Word.Range range)
        {
            if (range == null)
            {
                return;
            }

            range.Font.Italic = -1;
            range.Font.Color = ReferenceColor;
        }

        /// <summary>
        /// 独立成段的参考说明（表/Chart/图片下 * 行、文末参考区行）：灰斜 + 固定字体/字号/不加粗。
        /// 句末角标仍用 <see cref="ApplyMarkerStyle"/>（字体与字号继承宿主句）。
        /// </summary>
        public static void ApplyStandaloneReferenceStyle(Word.Range range)
        {
            if (range == null)
            {
                return;
            }

            ApplyReferenceStyle(range);
            range.Font.Name = ReferenceFontName;
            range.Font.Bold = 0;
            range.Font.Size = ReferenceFontSize;
            range.Font.Superscript = 0;
        }

        public static void ApplyMarkerStyle(Word.Range range)
        {
            ApplyReferenceStyle(range);
            if (range != null)
            {
                range.Font.Superscript = -1;
            }
        }
    }
}
