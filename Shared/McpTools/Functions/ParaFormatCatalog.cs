using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;

namespace WordAddIn1
{
    /// <summary>
    /// 段落格式键目录。无 COM。读省略和写入前清洗都在这里。
    /// </summary>
    public static class ParaFormatCatalog
    {
        public const int WdUndefined = 9999999;
        public const int WdToggle = 9999998;

        public static readonly string[] FingerprintKeys =
        {
            "alignment",
            "first_line_indent",
            "left_indent",
            "right_indent",
            "hanging_indent",
            "line_spacing_rule",
            "line_spacing",
            "space_before",
            "space_after",
        };

        private static readonly HashSet<string> KnownKeys = new HashSet<string>(StringComparer.Ordinal)
        {
            "alignment",
            "left_indent",
            "right_indent",
            "first_line_indent",
            "hanging_indent",
            "space_before",
            "space_after",
            "line_spacing_rule",
            "line_spacing",
            "paragraph_style",
            "outline_level",
            "baseline_alignment",
            "reading_order",
            "mirror_indents",
            "auto_adjust_right_indent",
            "character_unit_left_indent",
            "character_unit_right_indent",
            "character_unit_first_line_indent",
            "line_unit_before",
            "line_unit_after",
            "space_before_auto",
            "space_after_auto",
            "disable_line_height_grid",
            "widow_control",
            "keep_with_next",
            "keep_together",
            "page_break_before",
            "no_line_number",
            "hyphenation",
            "far_east_line_break_control",
            "word_wrap",
            "hanging_punctuation",
            "half_width_punctuation_on_top_of_line",
            "add_space_between_far_east_and_alpha",
            "add_space_between_far_east_and_digit",
            "textbox_tight_wrap",
            "collapsed_by_default",
            "shading",
            "borders",
            "tab_stops",
            "list",
            "drop_cap",
        };

        private static readonly HashSet<string> ZeroNumberKeys = new HashSet<string>(StringComparer.Ordinal)
        {
            "left_indent",
            "right_indent",
            "first_line_indent",
            "hanging_indent",
            "space_before",
            "space_after",
            "character_unit_left_indent",
            "character_unit_right_indent",
            "character_unit_first_line_indent",
            "line_unit_before",
            "line_unit_after",
        };

        private static readonly HashSet<string> FalseDefaultKeys = new HashSet<string>(StringComparer.Ordinal)
        {
            "keep_with_next",
            "keep_together",
            "page_break_before",
            "no_line_number",
            "mirror_indents",
            "space_before_auto",
            "space_after_auto",
            "disable_line_height_grid",
            "half_width_punctuation_on_top_of_line",
            "word_wrap",
            "collapsed_by_default",
        };

        private static readonly HashSet<string> TrueDefaultKeys = new HashSet<string>(StringComparer.Ordinal)
        {
            "widow_control",
            "hyphenation",
            "auto_adjust_right_indent",
            "far_east_line_break_control",
            "hanging_punctuation",
            "add_space_between_far_east_and_alpha",
            "add_space_between_far_east_and_digit",
        };

        private static readonly HashSet<string> AlignmentValues = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "left",
            "center",
            "right",
            "justify",
            "distribute",
            "justify_med",
            "justify_hi",
            "justify_low",
            "thai_justify",
            "inherit",
        };

        public static bool IsKnownKey(string key)
        {
            return !string.IsNullOrEmpty(key) && KnownKeys.Contains(key);
        }

        public static bool IsDefaultToken(object value)
        {
            return CharFormatCatalog.IsDefaultToken(value);
        }

        public static bool IsBoolKey(string key)
        {
            return FalseDefaultKeys.Contains(key) || TrueDefaultKeys.Contains(key);
        }

        /// <summary>文档默认层有该键就用它。没有则用与读省略同一含义的回退值。paragraph_style 不在这里。</summary>
        public static bool TryResolveWriteDefault(
            string key,
            IReadOnlyDictionary<string, object> documentLayer,
            out object value)
        {
            value = null;
            if (string.IsNullOrEmpty(key) || key == "paragraph_style" || !IsKnownKey(key))
            {
                return false;
            }

            if (documentLayer != null
                && documentLayer.TryGetValue(key, out object fromDocument)
                && fromDocument != null
                && !string.IsNullOrWhiteSpace(fromDocument.ToString()))
            {
                value = fromDocument;
                return true;
            }

            if (ZeroNumberKeys.Contains(key))
            {
                value = 0d;
                return true;
            }

            if (FalseDefaultKeys.Contains(key))
            {
                value = false;
                return true;
            }

            if (TrueDefaultKeys.Contains(key))
            {
                value = true;
                return true;
            }

            switch (key)
            {
                case "alignment":
                    value = "left";
                    return true;
                case "line_spacing_rule":
                    value = "single";
                    return true;
                case "line_spacing":
                    value = 1d;
                    return true;
                case "outline_level":
                    value = "wdOutlineLevelBodyText";
                    return true;
                case "baseline_alignment":
                    value = "wdBaselineAlignAuto";
                    return true;
                case "reading_order":
                    value = "wdReadingOrderLtr";
                    return true;
                case "textbox_tight_wrap":
                    value = "wdTightNone";
                    return true;
                case "tab_stops":
                    value = new List<object>();
                    return true;
                case "shading":
                    value = new Dictionary<string, object>
                    {
                        ["texture"] = "wdTextureNone",
                        ["background_pattern_color"] = "automatic",
                        ["foreground_pattern_color"] = "automatic",
                    };
                    return true;
                case "borders":
                    value = new Dictionary<string, object>
                    {
                        ["top"] = NoneEdge(),
                        ["bottom"] = NoneEdge(),
                        ["left"] = NoneEdge(),
                        ["right"] = NoneEdge(),
                        ["horizontal"] = NoneEdge(),
                    };
                    return true;
                case "list":
                    value = new Dictionary<string, object> { ["kind"] = "none" };
                    return true;
                case "drop_cap":
                    value = new Dictionary<string, object> { ["position"] = "none" };
                    return true;
                default:
                    return false;
            }
        }

        private static Dictionary<string, object> NoneEdge()
        {
            return new Dictionary<string, object> { ["line_style"] = "wdLineStyleNone" };
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

            if (value is string text)
            {
                text = text.Trim();
                return text == "9999999" || text == "9999998"
                    || string.Equals(text, "wdUndefined", StringComparison.OrdinalIgnoreCase)
                    || string.Equals(text, "wdToggle", StringComparison.OrdinalIgnoreCase);
            }

            try
            {
                double number = Convert.ToDouble(value, CultureInfo.InvariantCulture);
                int rounded = (int)Math.Round(number);
                return rounded == WdUndefined || rounded == WdToggle;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>写入前清洗。不改调用方字典。</summary>
        public static Dictionary<string, object> PrepareWriteSnapshot(
            IReadOnlyDictionary<string, object> paraFormat,
            List<Dictionary<string, object>> warnings)
        {
            var snap = new Dictionary<string, object>(StringComparer.Ordinal);
            if (paraFormat == null)
            {
                return snap;
            }

            if (warnings == null)
            {
                warnings = new List<Dictionary<string, object>>();
            }

            IReadOnlyDictionary<string, object> documentLayer = FormatDefaultSource.Current?.Paragraph;
            foreach (KeyValuePair<string, object> kv in paraFormat)
            {
                if (!IsKnownKey(kv.Key))
                {
                    Warn(warnings, kv.Key, "未知段落格式键");
                    continue;
                }

                object value = kv.Value;
                if (IsDefaultToken(value))
                {
                    if (kv.Key == "paragraph_style")
                    {
                        Warn(warnings, kv.Key, "不接受 default，套样式走样式工具");
                        continue;
                    }

                    if (!TryResolveWriteDefault(kv.Key, documentLayer, out value))
                    {
                        Warn(warnings, kv.Key, "解析不出文档默认");
                        continue;
                    }
                }

                if (IsBoolKey(kv.Key))
                {
                    if (!CharFormatCatalog.TryCoerceBool(value, out bool bit))
                    {
                        Warn(warnings, kv.Key, "布尔值无法识别");
                        continue;
                    }

                    value = bit;
                }

                if (IsSentinel(value))
                {
                    Warn(warnings, kv.Key, "哨兵值不写入");
                    continue;
                }

                snap[kv.Key] = value;
            }

            if (snap.TryGetValue("alignment", out object alignObj))
            {
                string align = alignObj?.ToString() ?? "";
                if (string.Equals(align, "inherit", StringComparison.OrdinalIgnoreCase))
                {
                    snap.Remove("alignment");
                }
                else if (!AlignmentValues.Contains(align))
                {
                    snap.Remove("alignment");
                    Warn(warnings, "alignment", "无法识别的对齐 " + align);
                }
            }

            if (snap.ContainsKey("first_line_indent") && snap.ContainsKey("hanging_indent"))
            {
                snap.Remove("first_line_indent");
                Warn(warnings, "first_line_indent", "与 hanging_indent 同时出现，只写悬挂缩进");
            }

            if (snap.TryGetValue("list", out object listRaw))
            {
                Dictionary<string, object> list = CopyDict(AsDict(listRaw));
                if (list == null)
                {
                    snap.Remove("list");
                    Warn(warnings, "list", "list 须为对象");
                }
                else
                {
                    if (list.ContainsKey("list_string"))
                    {
                        list.Remove("list_string");
                        Warn(warnings, "list", "list_string 只读");
                    }

                    if (list.TryGetValue("kind", out object kindObj))
                    {
                        string kind = (kindObj?.ToString() ?? "").Trim().ToLowerInvariant();
                        if (kind == "mixed" || kind == "list_num_only")
                        {
                            list.Remove("kind");
                            Warn(warnings, "list", "kind " + kind + " 只用于读");
                        }
                    }

                    if (list.Count == 0)
                    {
                        snap.Remove("list");
                    }
                    else
                    {
                        snap["list"] = list;
                    }
                }
            }

            return snap;
        }

        public static void OmitReadDefaults(Dictionary<string, object> format, string normalStyleName)
        {
            if (format == null || format.Count == 0)
            {
                return;
            }

            if (format.TryGetValue("line_spacing_rule", out object rule)
                && format.TryGetValue("line_spacing", out object spacing)
                && LineSpacingIsDefault(rule, spacing))
            {
                format.Remove("line_spacing_rule");
                format.Remove("line_spacing");
            }

            var remove = new List<string>();
            foreach (string key in format.Keys)
            {
                if (IsReadDefault(key, format[key], normalStyleName))
                {
                    remove.Add(key);
                }
            }

            foreach (string key in remove)
            {
                format.Remove(key);
            }
        }

        public static bool IsReadDefault(string key, object value, string normalStyleName)
        {
            if (value == null || IsSentinel(value))
            {
                return true;
            }

            if (ZeroNumberKeys.Contains(key))
            {
                return Near(value, 0);
            }

            if (FalseDefaultKeys.Contains(key))
            {
                return IsFalse(value);
            }

            if (TrueDefaultKeys.Contains(key))
            {
                return IsTrue(value);
            }

            switch (key)
            {
                case "alignment":
                    return TextIs(value, "left");
                case "paragraph_style":
                {
                    string style = value.ToString().Trim();
                    if (style.Length == 0)
                    {
                        return true;
                    }

                    return !string.IsNullOrEmpty(normalStyleName)
                        && string.Equals(style, normalStyleName.Trim(), StringComparison.Ordinal);
                }
                case "outline_level":
                    return TextIs(value, "wdOutlineLevelBodyText");
                case "baseline_alignment":
                    return TextIs(value, "wdBaselineAlignAuto");
                case "reading_order":
                    return TextIs(value, "wdReadingOrderLtr");
                case "textbox_tight_wrap":
                    return TextIs(value, "wdTightNone");
                case "shading":
                    return ShadingIsDefault(value);
                case "borders":
                    return BordersAreDefault(value);
                case "tab_stops":
                    return TabStopsAreDefault(value);
                case "list":
                    return ListIsDefault(value);
                case "drop_cap":
                    return DropCapIsDefault(value);
                default:
                    return false;
            }
        }

        public static bool LineSpacingIsDefault(object rule, object spacing)
        {
            return TextIs(rule, "single") && Near(spacing, 1.0);
        }

        private static bool ShadingIsDefault(object value)
        {
            Dictionary<string, object> shading = AsDict(value);
            if (shading == null || shading.Count == 0)
            {
                return true;
            }

            bool textureDefault = !shading.ContainsKey("texture") || TextIs(shading["texture"], "wdTextureNone");
            bool backDefault = !shading.ContainsKey("background_pattern_color")
                || IsAutomaticColor(shading["background_pattern_color"]);
            bool foreDefault = !shading.ContainsKey("foreground_pattern_color")
                || IsAutomaticColor(shading["foreground_pattern_color"]);
            return textureDefault && backDefault && foreDefault;
        }

        private static bool BordersAreDefault(object value)
        {
            Dictionary<string, object> borders = AsDict(value);
            if (borders == null || borders.Count == 0)
            {
                return true;
            }

            foreach (KeyValuePair<string, object> side in borders)
            {
                Dictionary<string, object> edge = AsDict(side.Value);
                if (edge == null)
                {
                    continue;
                }

                if (edge.ContainsKey("line_style") && !TextIs(edge["line_style"], "wdLineStyleNone"))
                {
                    return false;
                }
            }

            return true;
        }

        private static bool TabStopsAreDefault(object value)
        {
            if (value is string || value is IDictionary)
            {
                return false;
            }

            if (!(value is IEnumerable seq))
            {
                return true;
            }

            foreach (object item in seq)
            {
                if (item != null)
                {
                    return false;
                }
            }

            return true;
        }

        private static bool ListIsDefault(object value)
        {
            Dictionary<string, object> list = AsDict(value);
            if (list == null || list.Count == 0)
            {
                return true;
            }

            if (!list.TryGetValue("kind", out object kind) || kind == null)
            {
                return false;
            }

            return TextIs(kind, "none");
        }

        private static bool DropCapIsDefault(object value)
        {
            Dictionary<string, object> drop = AsDict(value);
            if (drop == null || drop.Count == 0)
            {
                return true;
            }

            return !drop.ContainsKey("position") || TextIs(drop["position"], "none");
        }

        private static bool IsAutomaticColor(object value)
        {
            if (value == null)
            {
                return true;
            }

            string text = value.ToString().Trim();
            return text.Length == 0 || string.Equals(text, "automatic", StringComparison.OrdinalIgnoreCase);
        }

        private static bool TextIs(object value, string expected)
        {
            return value != null
                && string.Equals(value.ToString().Trim(), expected, StringComparison.OrdinalIgnoreCase);
        }

        private static bool Near(object value, double expected)
        {
            try
            {
                double number = Convert.ToDouble(value, CultureInfo.InvariantCulture);
                return Math.Abs(number - expected) < 0.001;
            }
            catch
            {
                return false;
            }
        }

        private static bool IsFalse(object value)
        {
            if (value is bool flag)
            {
                return !flag;
            }

            string text = value?.ToString()?.Trim() ?? "";
            return text == "0" || string.Equals(text, "false", StringComparison.OrdinalIgnoreCase);
        }

        private static bool IsTrue(object value)
        {
            if (value is bool flag)
            {
                return flag;
            }

            string text = value?.ToString()?.Trim() ?? "";
            return text == "1" || text == "-1" || string.Equals(text, "true", StringComparison.OrdinalIgnoreCase);
        }

        private static void Warn(List<Dictionary<string, object>> warnings, string key, string message)
        {
            warnings.Add(new Dictionary<string, object>
            {
                ["key"] = key ?? "",
                ["message"] = message ?? "",
            });
        }

        private static Dictionary<string, object> CopyDict(Dictionary<string, object> source)
        {
            if (source == null)
            {
                return null;
            }

            return new Dictionary<string, object>(source, StringComparer.Ordinal);
        }

        private static Dictionary<string, object> AsDict(object raw)
        {
            if (raw is Dictionary<string, object> dict)
            {
                return dict;
            }

            if (raw is IDictionary generic)
            {
                var copy = new Dictionary<string, object>(StringComparer.Ordinal);
                foreach (DictionaryEntry entry in generic)
                {
                    copy[Convert.ToString(entry.Key)] = entry.Value;
                }

                return copy;
            }

            return null;
        }
    }
}
