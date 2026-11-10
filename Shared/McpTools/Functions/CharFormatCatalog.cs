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
