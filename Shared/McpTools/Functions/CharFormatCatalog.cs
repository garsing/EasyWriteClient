using System;
using System.Collections.Generic;
using System.Globalization;

namespace WordAddIn1
{
    /// <summary>
    /// 字符格式键目录。无 COM。真源给快照过滤、句内合并和技能白名单对齐。
    /// </summary>
    public static class CharFormatCatalog
    {
        public const int WdUndefined = 9999999;
        public const int WdToggle = 9999998;

        public const string KindCountKey = "format_kind_count";
        public const string KindNoteKey = "format_note";

        /// <summary>套用顺序：字符样式最先，旧键先于更具体的键。</summary>
        public static readonly string[] ApplyOrder =
        {
            "character_style",
            "font_name",
            "name", "name_ascii", "name_far_east", "name_bi", "name_other",
            "font_size", "size_bi",
            "font_color", "color_index", "color_index_bi", "text_color",
            "underline_color", "diacritic_color",
            "is_bold", "bold_bi", "is_italic", "italic_bi",
            "has_underline", "underline",
            "has_strikethrough", "double_strikethrough",
            "is_subscript", "is_superscript",
            "shadow", "outline", "emboss", "engrave", "small_caps", "all_caps", "hidden", "emphasis_mark",
            "scaling", "spacing", "position", "kerning", "disable_character_space_grid",
            "ligatures", "number_form", "number_spacing", "stylistic_set", "contextual_alternates",
            "animation",
            "background_color", "shading", "highlight_color_index",
            "borders", "fill", "glow", "line", "reflection", "text_shadow", "three_d",
            "language_id", "language_id_far_east", "no_proofing", "fit_text_width", "character_width"
        };

        private static readonly HashSet<string> KeySet = new HashSet<string>(ApplyOrder, StringComparer.Ordinal);

        /// <summary>读结果里始终保留的键，默认值也留下。</summary>
        public static readonly string[] AlwaysReadKeys =
        {
            "font_name", "font_size", "font_color", "background_color",
            "is_bold", "is_italic", "has_underline", "has_strikethrough",
            "is_superscript", "is_subscript"
        };

        private static readonly HashSet<string> AlwaysReadSet = new HashSet<string>(AlwaysReadKeys, StringComparer.Ordinal);

        public static bool IsKnownKey(string key)
        {
            return !string.IsNullOrEmpty(key) && KeySet.Contains(key);
        }

        public static bool IsSentinel(object value)
        {
            if (value == null)
            {
                return false;
            }

            if (value is bool)
            {
                return false;
            }

            if (TryGetFiniteNumber(value, out double number))
            {
                int rounded = (int)Math.Round(number);
                return rounded == WdUndefined || rounded == WdToggle;
            }

            string text = value as string;
            if (text == null)
            {
                return false;
            }

            text = text.Trim();
            return text == "9999999" || text == "9999998"
                || string.Equals(text, "wdUndefined", StringComparison.OrdinalIgnoreCase)
                || string.Equals(text, "wdToggle", StringComparison.OrdinalIgnoreCase);
        }

        /// <summary>读路径：10 个重要键保留，其余键只有偏离默认才留下。</summary>
        public static void OmitReadDefaults(Dictionary<string, object> format)
        {
            if (format == null || format.Count == 0)
            {
                return;
            }

            var remove = new List<string>();
            foreach (string key in format.Keys)
            {
                if (AlwaysReadSet.Contains(key))
                {
                    continue;
                }

                if (IsReadDefault(key, format[key], format))
                {
                    remove.Add(key);
                }
            }

            foreach (string key in remove)
            {
                format.Remove(key);
            }
        }

        public static bool IsReadDefault(string key, object value, Dictionary<string, object> format)
        {
            if (value == null || IsSentinel(value))
            {
                return true;
            }

            switch (key)
            {
                case "name":
                case "name_ascii":
                case "name_far_east":
                case "name_bi":
                case "name_other":
                    return SameText(value, format, "font_name") || string.IsNullOrWhiteSpace(value.ToString());
                case "size_bi":
                    return Near(value, 0) || SameNumber(value, format, "font_size");
                case "bold_bi":
                case "italic_bi":
                case "double_strikethrough":
                case "shadow":
                case "outline":
                case "emboss":
                case "engrave":
                case "small_caps":
                case "all_caps":
                case "hidden":
                case "contextual_alternates":
                case "disable_character_space_grid":
                case "no_proofing":
                    return IsFalse(value);
                case "color_index":
                case "color_index_bi":
                    return EnumIs(value, "wdAuto");
                case "text_color":
                    return TextColorIsDefault(value, format);
                case "underline_color":
                case "diacritic_color":
                    return IsAutomaticColor(value);
                case "underline":
                    return EnumIs(value, "wdUnderlineNone") || EnumIs(value, "wdUnderlineSingle");
                case "emphasis_mark":
                    return EnumIs(value, "wdEmphasisMarkNone");
                case "animation":
                    return EnumIs(value, "wdAnimationNone");
                case "scaling":
                    return Near(value, 100);
                case "spacing":
                case "position":
                case "kerning":
                case "fit_text_width":
                    return Near(value, 0);
                case "ligatures":
                    return EnumIs(value, "wdLigaturesNone");
                case "number_form":
                    return EnumIs(value, "wdNumberFormDefault");
                case "number_spacing":
                    return EnumIs(value, "wdNumberSpacingDefault");
                case "stylistic_set":
                    return EnumIs(value, "wdStylisticSetDefault");
                case "shading":
                    return ShadingIsDefault(value);
                case "borders":
                    return BordersAreDefault(value);
                case "fill":
                case "line":
                case "text_shadow":
                case "three_d":
                    return EffectOff(value, "visible");
                case "glow":
                    return GlowIsDefault(value);
                case "reflection":
                    return ReflectionIsDefault(value);
                case "highlight_color_index":
                    return EnumIs(value, "wdNoHighlight");
                case "character_style":
                    return IsDefaultCharacterStyle(value);
                case "language_id":
                case "language_id_far_east":
                    return EnumIs(value, "wdLanguageNone");
                case "character_width":
                    return EnumIs(value, "wdWidthDefault");
                default:
                    return false;
            }
        }

        private static bool TextColorIsDefault(object value, Dictionary<string, object> format)
        {
            Dictionary<string, object> map = AsStringDict(value);
            if (map == null)
            {
                return true;
            }

            if (map.TryGetValue("object_theme_color", out object theme) && theme != null && !IsSentinel(theme))
            {
                return false;
            }

            if (!map.TryGetValue("rgb", out object rgb) || rgb == null)
            {
                return true;
            }

            return IsAutomaticColor(rgb) || SameText(rgb, format, "font_color");
        }

        private static bool ShadingIsDefault(object value)
        {
            Dictionary<string, object> map = AsStringDict(value);
            if (map == null)
            {
                return true;
            }

            bool textureDefault = !map.TryGetValue("texture", out object texture)
                || texture == null
                || EnumIs(texture, "wdTextureNone");
            bool backDefault = !map.TryGetValue("background_pattern_color", out object back)
                || back == null
                || IsAutomaticColor(back);
            bool foreDefault = !map.TryGetValue("foreground_pattern_color", out object fore)
                || fore == null
                || IsAutomaticColor(fore);
            return textureDefault && backDefault && foreDefault;
        }

        private static bool BordersAreDefault(object value)
        {
            Dictionary<string, object> map = AsStringDict(value);
            if (map == null)
            {
                return true;
            }

            return !map.TryGetValue("outside_line_style", out object style)
                || style == null
                || EnumIs(style, "wdLineStyleNone");
        }

        private static bool EffectOff(object value, string visibleKey)
        {
            Dictionary<string, object> map = AsStringDict(value);
            if (map == null)
            {
                return true;
            }

            if (!map.TryGetValue(visibleKey, out object visible) || visible == null)
            {
                return true;
            }

            return IsFalse(visible);
        }

        private static bool GlowIsDefault(object value)
        {
            Dictionary<string, object> map = AsStringDict(value);
            if (map == null)
            {
                return true;
            }

            return !map.TryGetValue("radius", out object radius) || radius == null || Near(radius, 0);
        }

        private static bool ReflectionIsDefault(object value)
        {
            Dictionary<string, object> map = AsStringDict(value);
            if (map == null)
            {
                return true;
            }

            if (!map.TryGetValue("type", out object type) || type == null)
            {
                return true;
            }

            return EnumIs(type, "msoReflectionTypeNone");
        }

        private static bool IsDefaultCharacterStyle(object value)
        {
            string text = value?.ToString()?.Trim() ?? "";
            return text.Length == 0
                || text == "默认段落字体"
                || string.Equals(text, "Default Paragraph Font", StringComparison.OrdinalIgnoreCase);
        }

        private static bool IsAutomaticColor(object value)
        {
            string text = value?.ToString()?.Trim() ?? "";
            return text.Length == 0 || string.Equals(text, "automatic", StringComparison.OrdinalIgnoreCase);
        }

        private static bool IsFalse(object value)
        {
            if (value is bool flag)
            {
                return !flag;
            }

            string text = value?.ToString();
            return text == "false" || text == "False" || text == "0";
        }

        private static bool EnumIs(object value, string name)
        {
            return string.Equals(value?.ToString()?.Trim(), name, StringComparison.OrdinalIgnoreCase);
        }

        private static bool SameText(object value, Dictionary<string, object> format, string otherKey)
        {
            if (format == null || !format.TryGetValue(otherKey, out object other) || other == null)
            {
                return false;
            }

            return string.Equals(value.ToString().Trim(), other.ToString().Trim(), StringComparison.OrdinalIgnoreCase);
        }

        private static bool SameNumber(object value, Dictionary<string, object> format, string otherKey)
        {
            if (format == null || !format.TryGetValue(otherKey, out object other))
            {
                return false;
            }

            return Near(value, other);
        }

        private static bool Near(object value, object other)
        {
            return TryGetFiniteNumber(value, out double left)
                && TryGetFiniteNumber(other, out double right)
                && Math.Abs(left - right) < 0.001;
        }

        private static bool Near(object value, double expected)
        {
            return TryGetFiniteNumber(value, out double number) && Math.Abs(number - expected) < 0.001;
        }

        public static string KindNote(int kindCount)
        {
            if (kindCount <= 1)
            {
                return null;
            }

            string[] names = { "", "", "两", "三", "四", "五", "六", "七", "八", "九", "十" };
            string num = kindCount <= 10 ? names[kindCount] : kindCount.ToString(CultureInfo.InvariantCulture);
            return "这一句有" + num + "种格式";
        }

        public static void MergeRunSnapshots(
            IList<Dictionary<string, object>> runs,
            out Dictionary<string, object> merged,
            out int kindCount,
            out string note)
        {
            merged = new Dictionary<string, object>();
            var kinds = new List<Dictionary<string, object>>();
            if (runs != null)
            {
                foreach (Dictionary<string, object> run in runs)
                {
                    if (run == null)
                    {
                        continue;
                    }

                    bool seen = false;
                    foreach (Dictionary<string, object> kind in kinds)
                    {
                        if (SameSnapshot(kind, run))
                        {
                            seen = true;
                            break;
                        }
                    }

                    if (!seen)
                    {
                        kinds.Add(run);
                    }
                }
            }

            kindCount = kinds.Count;
            note = KindNote(kindCount);
            if (kinds.Count == 0)
            {
                return;
            }

            if (kinds.Count == 1)
            {
                foreach (KeyValuePair<string, object> kv in kinds[0])
                {
                    merged[kv.Key] = kv.Value;
                }

                return;
            }

            foreach (string key in kinds[0].Keys)
            {
                object value = kinds[0][key];
                bool same = true;
                for (int i = 1; i < kinds.Count; i++)
                {
                    if (!kinds[i].TryGetValue(key, out object other) || !SameValue(value, other))
                    {
                        same = false;
                        break;
                    }
                }

                if (same)
                {
                    merged[key] = value;
                }
            }
        }

        public static bool SameSnapshot(Dictionary<string, object> left, Dictionary<string, object> right)
        {
            if (left == null || right == null)
            {
                return left == right;
            }

            if (left.Count != right.Count)
            {
                return false;
            }

            foreach (KeyValuePair<string, object> kv in left)
            {
                if (!right.TryGetValue(kv.Key, out object other) || !SameValue(kv.Value, other))
                {
                    return false;
                }
            }

            return true;
        }

        public static bool SameValue(object left, object right)
        {
            if (left == null || right == null)
            {
                return left == null && right == null;
            }

            Dictionary<string, object> leftMap = AsStringDict(left);
            Dictionary<string, object> rightMap = AsStringDict(right);
            if (leftMap != null || rightMap != null)
            {
                return leftMap != null && rightMap != null && SameSnapshot(leftMap, rightMap);
            }

            if (TryGetFiniteNumber(left, out double ln) && TryGetFiniteNumber(right, out double rn))
            {
                return Math.Abs(ln - rn) < 0.001;
            }

            if (left is bool lb && right is bool rb)
            {
                return lb == rb;
            }

            return string.Equals(Convert.ToString(left, CultureInfo.InvariantCulture),
                Convert.ToString(right, CultureInfo.InvariantCulture),
                StringComparison.Ordinal);
        }

        public static Dictionary<string, object> AsStringDict(object value)
        {
            return value as Dictionary<string, object>;
        }

        public static bool TryGetFiniteNumber(object value, out double number)
        {
            switch (value)
            {
                case byte n:
                    number = n;
                    return true;
                case short n:
                    number = n;
                    return true;
                case int n:
                    number = n;
                    return true;
                case long n:
                    number = n;
                    return true;
                case float n:
                    number = n;
                    return !float.IsNaN(n) && !float.IsInfinity(n);
                case double n:
                    number = n;
                    return !double.IsNaN(n) && !double.IsInfinity(n);
                case decimal n:
                    number = (double)n;
                    return true;
                default:
                    return double.TryParse(
                        Convert.ToString(value, CultureInfo.InvariantCulture),
                        NumberStyles.Float,
                        CultureInfo.InvariantCulture,
                        out number);
            }
        }
    }
}
