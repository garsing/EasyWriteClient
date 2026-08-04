using System;
using System.Collections.Generic;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    /// <summary>
    /// 批次 2：从 Word 段落 Range 读取 §2.1 para_format（effective COM）。
    /// </summary>
    public static class ParaFormatReader
    {
        public const string DbgPrefix = "[段落格式读][DBG]";

        public static void DbgLog(string message)
        {
            if (!FormatInheritHelper.EnableDbg)
            {
                return;
            }

            System.Diagnostics.Debug.WriteLine($"{DbgPrefix} {message}");
        }

        /// <summary>paragraphRange 须为整段（ParagraphCodeResolver 上卷后）。</summary>
        public static Dictionary<string, object> Extract(Word.Range paragraphRange)
        {
            var format = new Dictionary<string, object>();
            if (paragraphRange == null)
            {
                DbgLog("Extract range=null");
                return format;
            }

            try
            {
                Word.ParagraphFormat pf = paragraphRange.ParagraphFormat;
                if (pf == null)
                {
                    DbgLog("Extract ParagraphFormat=null");
                    ApplyDefaultLineSpacing(format);
                    return format;
                }

                string alignment = MapAlignment(pf.Alignment);
                if (!string.IsNullOrEmpty(alignment))
                {
                    format["alignment"] = alignment;
                }

                AppendIndentFields(format, pf);
                AppendSpacingFields(format, pf);
                ApplyDefaultLineSpacing(format);
            }
            catch (Exception ex)
            {
                DbgLog($"Extract 异常: {ex.Message}");
                ApplyDefaultLineSpacing(format);
            }

            return format;
        }

        private static void AppendIndentFields(Dictionary<string, object> format, Word.ParagraphFormat pf)
        {
            float left = pf.LeftIndent;
            if (IsNonZero(left))
            {
                format["left_indent"] = RoundPt(left);
            }

            float right = pf.RightIndent;
            if (IsNonZero(right))
            {
                format["right_indent"] = RoundPt(right);
            }

            float firstLine = pf.FirstLineIndent;
            if (firstLine < -0.05f)
            {
                format["hanging_indent"] = RoundPt(-firstLine);
            }
            else if (IsNonZero(firstLine))
            {
                format["first_line_indent"] = RoundPt(firstLine);
            }
        }

        private static void AppendSpacingFields(Dictionary<string, object> format, Word.ParagraphFormat pf)
        {
            float spaceBefore = pf.SpaceBefore;
            if (IsNonZero(spaceBefore))
            {
                format["space_before"] = RoundPt(spaceBefore);
            }

            float spaceAfter = pf.SpaceAfter;
            if (IsNonZero(spaceAfter))
            {
                format["space_after"] = RoundPt(spaceAfter);
            }

            Dictionary<string, object> linePair = ReadLineSpacing(pf);
            format["line_spacing_rule"] = linePair["line_spacing_rule"];
            format["line_spacing"] = linePair["line_spacing"];
        }

        private static void ApplyDefaultLineSpacing(Dictionary<string, object> format)
        {
            if (!format.ContainsKey("line_spacing_rule") || !format.ContainsKey("line_spacing"))
            {
                format["line_spacing_rule"] = "single";
                format["line_spacing"] = 1.0;
            }
        }

        internal static Dictionary<string, object> ReadLineSpacing(Word.ParagraphFormat pf)
        {
            Word.WdLineSpacing rule = pf.LineSpacingRule;
            float raw = pf.LineSpacing;

            switch (rule)
            {
                case Word.WdLineSpacing.wdLineSpace1pt5:
                    return NormalizeLineSpacingPair("1.5", 1.5);
                case Word.WdLineSpacing.wdLineSpaceDouble:
                    return NormalizeLineSpacingPair("double", 2.0);
                case Word.WdLineSpacing.wdLineSpaceAtLeast:
                    return NormalizeLineSpacingPair("at_least", raw);
                case Word.WdLineSpacing.wdLineSpaceExactly:
                    return NormalizeLineSpacingPair("exact", raw);
                case Word.WdLineSpacing.wdLineSpaceMultiple:
                    return NormalizeMultipleLineSpacing(raw);
                case Word.WdLineSpacing.wdLineSpaceSingle:
                default:
                    return NormalizeLineSpacingPair("single", 1.0);
            }
        }

        private static Dictionary<string, object> NormalizeMultipleLineSpacing(float raw)
        {
            double mult = raw;
            if (raw > 5.0f)
            {
                mult = raw / 12.0;
            }

            return NormalizeLineSpacingPair("multiple", mult);
        }

        internal static Dictionary<string, object> NormalizeLineSpacingPair(string rule, double lineSpacing)
        {
            string normalizedRule = string.IsNullOrEmpty(rule) ? "single" : rule;
            switch (normalizedRule)
            {
                case "single":
                    return Pair("single", RoundLine(1.0));
                case "1.5":
                    return Pair("1.5", RoundLine(1.5));
                case "double":
                    return Pair("double", RoundLine(2.0));
                case "multiple":
                {
                    double ls = RoundLine(lineSpacing <= 0 ? 1.0 : lineSpacing);
                    if (Math.Abs(ls - 1.0) < 0.001)
                    {
                        return Pair("single", 1.0);
                    }

                    if (Math.Abs(ls - 1.5) < 0.001)
                    {
                        return Pair("1.5", 1.5);
                    }

                    if (Math.Abs(ls - 2.0) < 0.001)
                    {
                        return Pair("double", 2.0);
                    }

                    return Pair("multiple", ls);
                }

                case "exact":
                case "at_least":
                    return Pair(normalizedRule, Math.Round(lineSpacing, 1));
                default:
                    return Pair("single", 1.0);
            }
        }

        private static Dictionary<string, object> Pair(string rule, double spacing)
        {
            return new Dictionary<string, object>
            {
                ["line_spacing_rule"] = rule,
                ["line_spacing"] = spacing,
            };
        }

        private static string MapAlignment(Word.WdParagraphAlignment alignment)
        {
            switch (alignment)
            {
                case Word.WdParagraphAlignment.wdAlignParagraphCenter:
                    return "center";
                case Word.WdParagraphAlignment.wdAlignParagraphRight:
                    return "right";
                case Word.WdParagraphAlignment.wdAlignParagraphJustify:
                    return "justify";
                case Word.WdParagraphAlignment.wdAlignParagraphDistribute:
                    return "distribute";
                default:
                    return "left";
            }
        }

        private static bool IsNonZero(float value)
        {
            return Math.Abs(value) > 0.05f;
        }

        private static double RoundPt(float value)
        {
            return Math.Round(value, 1);
        }

        private static double RoundLine(double value)
        {
            return Math.Round(value, 2);
        }
    }
}
