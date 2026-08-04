using System;
using System.Collections.Generic;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    /// <summary>
    /// 批次 2：explicit 段落 format 写入（§2.1 direct 字段；不 set_Style）。
    /// </summary>
    public static class ParaFormatWriter
    {
        public const string WarningStyleDirectOverride = "para_format_style_direct_override";

        /// <summary>旧 cluster / 显式 payload 中的 paragraph_style 忽略，不参与 apply。</summary>
        public static Dictionary<string, object> SanitizeParaFormatForApply(
            IReadOnlyDictionary<string, object> paraFormat)
        {
            var sanitized = new Dictionary<string, object>(StringComparer.Ordinal);
            if (paraFormat == null)
            {
                return sanitized;
            }

            foreach (KeyValuePair<string, object> kv in paraFormat)
            {
                if (string.Equals(kv.Key, "paragraph_style", StringComparison.Ordinal))
                {
                    continue;
                }

                sanitized[kv.Key] = kv.Value;
            }

            return sanitized;
        }

        public static List<string> Apply(
            Word.Range paragraphRange,
            IReadOnlyDictionary<string, object> paraFormat,
            out List<string> warnings)
        {
            warnings = new List<string>();
            var applied = new List<string>();
            Dictionary<string, object> sanitized = SanitizeParaFormatForApply(paraFormat);
            if (paragraphRange == null || sanitized.Count == 0)
            {
                return applied;
            }

            foreach (KeyValuePair<string, object> kv in sanitized)
            {
                string field = kv.Key;
                if (string.Equals(field, "line_spacing_rule", StringComparison.Ordinal)
                    || string.Equals(field, "line_spacing", StringComparison.Ordinal))
                {
                    continue;
                }

                if (ShouldSkipDirectField(field, kv.Value))
                {
                    continue;
                }

                ApplyDirectField(paragraphRange, field, kv.Value);
                applied.Add(field);
            }

            if (sanitized.ContainsKey("line_spacing_rule") || sanitized.ContainsKey("line_spacing"))
            {
                string rule = sanitized.TryGetValue("line_spacing_rule", out object r)
                    ? r?.ToString() ?? "single"
                    : "single";
                double spacing = sanitized.TryGetValue("line_spacing", out object s)
                    ? Convert.ToDouble(s)
                    : 1.0;
                Dictionary<string, object> normalized = ParaFormatReader.NormalizeLineSpacingPair(rule, spacing);
                string normRule = normalized["line_spacing_rule"]?.ToString();
                double normSpacing = Convert.ToDouble(normalized["line_spacing"]);

                ApplyLineSpacingPair(paragraphRange.ParagraphFormat, normRule, normSpacing);
                applied.Add("line_spacing_rule");
                applied.Add("line_spacing");
            }

            return applied;
        }

        private static bool ShouldSkipDirectField(string key, object value)
        {
            if (string.Equals(key, "alignment", StringComparison.Ordinal)
                && string.Equals(value?.ToString(), "inherit", StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            return false;
        }

        private static void ApplyDirectField(Word.Range paragraphRange, string key, object value)
        {
            Word.ParagraphFormat pf = paragraphRange.ParagraphFormat;
            switch (key)
            {
                case "alignment":
                    pf.Alignment = ParseAlignment(value?.ToString());
                    break;
                case "left_indent":
                    pf.LeftIndent = ToFloat(value);
                    break;
                case "right_indent":
                    pf.RightIndent = ToFloat(value);
                    break;
                case "first_line_indent":
                    pf.FirstLineIndent = ToFloat(value);
                    break;
                case "hanging_indent":
                    pf.FirstLineIndent = -ToFloat(value);
                    break;
                case "space_before":
                    pf.SpaceBefore = ToFloat(value);
                    break;
                case "space_after":
                    pf.SpaceAfter = ToFloat(value);
                    break;
            }
        }

        internal static void ApplyLineSpacingPair(Word.ParagraphFormat pf, string rule, double spacing)
        {
            switch ((rule ?? "single").ToLowerInvariant())
            {
                case "1.5":
                    pf.LineSpacingRule = Word.WdLineSpacing.wdLineSpace1pt5;
                    break;
                case "double":
                    pf.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceDouble;
                    break;
                case "at_least":
                    pf.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceAtLeast;
                    pf.LineSpacing = (float)spacing;
                    break;
                case "exact":
                    pf.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceExactly;
                    pf.LineSpacing = (float)spacing;
                    break;
                case "multiple":
                    pf.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceMultiple;
                    pf.LineSpacing = spacing <= 3.0 ? (float)spacing * 12f : (float)spacing;
                    break;
                default:
                    pf.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceSingle;
                    break;
            }
        }

        private static Word.WdParagraphAlignment ParseAlignment(string s)
        {
            switch ((s ?? "").ToLowerInvariant())
            {
                case "center":
                    return Word.WdParagraphAlignment.wdAlignParagraphCenter;
                case "right":
                    return Word.WdParagraphAlignment.wdAlignParagraphRight;
                case "justify":
                    return Word.WdParagraphAlignment.wdAlignParagraphJustify;
                case "distribute":
                    return Word.WdParagraphAlignment.wdAlignParagraphDistribute;
                default:
                    return Word.WdParagraphAlignment.wdAlignParagraphLeft;
            }
        }

        private static float ToFloat(object value)
        {
            return Convert.ToSingle(value);
        }
    }
}
