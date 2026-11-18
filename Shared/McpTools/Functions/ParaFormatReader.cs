using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    /// <summary>
    /// 从段落 Range 读取 para_format。默认值由 ParaFormatCatalog.OmitReadDefaults 拿掉。
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

            string normalStyle = ReadNormalStyleName(paragraphRange);
            TryRead(format, "alignment", () => MapAlignment(paragraphRange.ParagraphFormat.Alignment));
            TryReadParagraphFormat(format, paragraphRange);
            TryReadStyle(format, paragraphRange);
            TryReadShading(format, paragraphRange);
            TryReadBorders(format, paragraphRange);
            TryReadTabStops(format, paragraphRange);
            TryReadList(format, paragraphRange);
            TryReadDropCap(format, paragraphRange);
            ParaFormatCatalog.OmitReadDefaults(format, normalStyle);
            return format;
        }

        private static void TryReadParagraphFormat(Dictionary<string, object> format, Word.Range range)
        {
            Word.ParagraphFormat pf = range.ParagraphFormat;
            if (pf == null)
            {
                return;
            }

            TryReadIndentAndSpacing(format, pf);
            TryReadFloat(format, "character_unit_left_indent", () => pf.CharacterUnitLeftIndent);
            TryReadFloat(format, "character_unit_right_indent", () => pf.CharacterUnitRightIndent);
            TryReadFloat(format, "character_unit_first_line_indent", () => pf.CharacterUnitFirstLineIndent);
            TryReadFloat(format, "line_unit_before", () => pf.LineUnitBefore);
            TryReadFloat(format, "line_unit_after", () => pf.LineUnitAfter);
            TryReadBool(format, "mirror_indents", () => pf.MirrorIndents);
            TryReadBool(format, "auto_adjust_right_indent", () => pf.AutoAdjustRightIndent);
            TryReadBool(format, "space_before_auto", () => pf.SpaceBeforeAuto);
            TryReadBool(format, "space_after_auto", () => pf.SpaceAfterAuto);
            TryReadBool(format, "disable_line_height_grid", () => pf.DisableLineHeightGrid);
            TryReadBool(format, "widow_control", () => pf.WidowControl);
            TryReadBool(format, "keep_with_next", () => pf.KeepWithNext);
            TryReadBool(format, "keep_together", () => pf.KeepTogether);
            TryReadBool(format, "page_break_before", () => pf.PageBreakBefore);
            TryReadBool(format, "no_line_number", () => pf.NoLineNumber);
            TryReadBool(format, "hyphenation", () => pf.Hyphenation);
            TryReadBool(format, "far_east_line_break_control", () => pf.FarEastLineBreakControl);
            TryReadBool(format, "word_wrap", () => pf.WordWrap);
            TryReadBool(format, "hanging_punctuation", () => pf.HangingPunctuation);
            TryReadBool(format, "half_width_punctuation_on_top_of_line", () => pf.HalfWidthPunctuationOnTopOfLine);
            TryReadFarEastSpacing(format, range, pf);
            TryReadBool(format, "collapsed_by_default", () => pf.CollapsedByDefault);
            TryReadEnum(format, "outline_level", () => pf.OutlineLevel);
            TryReadEnum(format, "baseline_alignment", () => pf.BaseLineAlignment);
            TryReadEnum(format, "reading_order", () => pf.ReadingOrder);
            TryReadEnum(format, "textbox_tight_wrap", () => pf.TextboxTightWrap);
        }

        private static void TryReadIndentAndSpacing(Dictionary<string, object> format, Word.ParagraphFormat pf)
        {
            TryReadFloat(format, "left_indent", () => pf.LeftIndent);
            TryReadFloat(format, "right_indent", () => pf.RightIndent);
            try
            {
                float firstLine = pf.FirstLineIndent;
                if (Math.Abs(firstLine - ParaFormatCatalog.WdUndefined) < 1f)
                {
                    return;
                }

                if (firstLine < -0.05f)
                {
                    format["hanging_indent"] = RoundPt(Math.Abs(firstLine));
                }
                else
                {
                    format["first_line_indent"] = RoundPt(firstLine);
                }
            }
            catch (Exception ex)
            {
                DbgLog("first_line_indent 读取失败: " + ex.Message);
            }

            TryReadFloat(format, "space_before", () => pf.SpaceBefore);
            TryReadFloat(format, "space_after", () => pf.SpaceAfter);
            try
            {
                Dictionary<string, object> linePair = ReadLineSpacing(pf);
                format["line_spacing_rule"] = linePair["line_spacing_rule"];
                format["line_spacing"] = linePair["line_spacing"];
            }
            catch (Exception ex)
            {
                DbgLog("line_spacing 读取失败: " + ex.Message);
            }
        }

        private static void TryReadStyle(Dictionary<string, object> format, Word.Range range)
        {
            try
            {
                object styleObj = range.ParagraphFormat?.get_Style();
                Word.Style style = styleObj as Word.Style;
                string name = style?.NameLocal;
                if (!string.IsNullOrWhiteSpace(name))
                {
                    format["paragraph_style"] = name;
                }
            }
            catch (Exception ex)
            {
                DbgLog("paragraph_style 读取失败: " + ex.Message);
            }
        }

        private static string ReadNormalStyleName(Word.Range range)
        {
            try
            {
                object index = Word.WdBuiltinStyle.wdStyleNormal;
                Word.Style style = range.Document.Styles.get_Item(ref index);
                return style?.NameLocal ?? "";
            }
            catch (Exception ex)
            {
                DbgLog("Normal 样式名读取失败: " + ex.Message);
                return "";
            }
        }

        private static void TryReadShading(Dictionary<string, object> format, Word.Range range)
        {
            try
            {
                Word.Shading shading = range.ParagraphFormat.Shading;
                var map = new Dictionary<string, object>();
                TryPutEnum(map, "texture", () => shading.Texture);
                TryPutColor(map, "background_pattern_color", () => shading.BackgroundPatternColor);
                TryPutColor(map, "foreground_pattern_color", () => shading.ForegroundPatternColor);
                if (map.Count > 0)
                {
                    format["shading"] = map;
                }
            }
            catch (Exception ex)
            {
                DbgLog("shading 读取失败: " + ex.Message);
            }
        }

        private static void TryReadBorders(Dictionary<string, object> format, Word.Range range)
        {
            try
            {
                Word.Borders borders = range.ParagraphFormat.Borders;
                var map = new Dictionary<string, object>();
                ReadBorder(map, borders, "top", Word.WdBorderType.wdBorderTop);
                ReadBorder(map, borders, "bottom", Word.WdBorderType.wdBorderBottom);
                ReadBorder(map, borders, "left", Word.WdBorderType.wdBorderLeft);
                ReadBorder(map, borders, "right", Word.WdBorderType.wdBorderRight);
                ReadBorder(map, borders, "horizontal", Word.WdBorderType.wdBorderHorizontal);
                if (map.Count > 0)
                {
                    format["borders"] = map;
                }
            }
            catch (Exception ex)
            {
                DbgLog("borders 读取失败: " + ex.Message);
            }
        }

        private static void ReadBorder(
            Dictionary<string, object> map,
            Word.Borders borders,
            string key,
            Word.WdBorderType which)
        {
            try
            {
                Word.Border border = borders[which];
                if (border == null || border.LineStyle == Word.WdLineStyle.wdLineStyleNone)
                {
                    return;
                }

                var edge = new Dictionary<string, object>();
                string lineStyle = Enum.GetName(typeof(Word.WdLineStyle), border.LineStyle);
                if (!string.IsNullOrEmpty(lineStyle))
                {
                    edge["line_style"] = lineStyle;
                }

                string lineWidth = Enum.GetName(typeof(Word.WdLineWidth), border.LineWidth);
                if (!string.IsNullOrEmpty(lineWidth))
                {
                    edge["line_width"] = lineWidth;
                }

                edge["color"] = ColorToRgb(border.Color);
                if (edge.Count > 0)
                {
                    map[key] = edge;
                }
            }
            catch (Exception ex)
            {
                DbgLog("border " + key + " 读取失败: " + ex.Message);
            }
        }

        private static void TryReadTabStops(Dictionary<string, object> format, Word.Range range)
        {
            try
            {
                var tabs = new List<Dictionary<string, object>>();
                foreach (Word.TabStop stop in range.ParagraphFormat.TabStops)
                {
                    try
                    {
                        if (!stop.CustomTab)
                        {
                            continue;
                        }

                        var item = new Dictionary<string, object>
                        {
                            ["position"] = RoundPt(stop.Position),
                        };
                        string alignment = Enum.GetName(typeof(Word.WdTabAlignment), stop.Alignment);
                        if (!string.IsNullOrEmpty(alignment))
                        {
                            item["alignment"] = alignment;
                        }

                        string leader = Enum.GetName(typeof(Word.WdTabLeader), stop.Leader);
                        if (!string.IsNullOrEmpty(leader))
                        {
                            item["leader"] = leader;
                        }

                        tabs.Add(item);
                    }
                    catch (Exception ex)
                    {
                        DbgLog("tab 读取失败: " + ex.Message);
                    }
                }

                if (tabs.Count > 0)
                {
                    format["tab_stops"] = tabs;
                }
            }
            catch (Exception ex)
            {
                DbgLog("tab_stops 读取失败: " + ex.Message);
            }
        }

        private static void TryReadList(Dictionary<string, object> format, Word.Range range)
        {
            try
            {
                Word.ListFormat listFormat = range.ListFormat;
                if (listFormat == null)
                {
                    return;
                }

                string kind = MapListKind(listFormat.ListType);
                if (string.IsNullOrEmpty(kind))
                {
                    return;
                }

                var list = new Dictionary<string, object>
                {
                    ["kind"] = kind,
                };
                try
                {
                    list["level"] = listFormat.ListLevelNumber;
                }
                catch (Exception ex)
                {
                    DbgLog("list.level 读取失败: " + ex.Message);
                }

                try
                {
                    string shown = listFormat.ListString;
                    if (!string.IsNullOrEmpty(shown))
                    {
                        list["list_string"] = shown;
                    }
                }
                catch (Exception ex)
                {
                    DbgLog("list_string 读取失败: " + ex.Message);
                }

                TryReadListLevel(list, listFormat, kind);
                format["list"] = list;
            }
            catch (Exception ex)
            {
                DbgLog("list 读取失败: " + ex.Message);
            }
        }

        private static void TryReadListLevel(Dictionary<string, object> list, Word.ListFormat listFormat, string kind)
        {
            try
            {
                Word.ListTemplate template = listFormat.ListTemplate;
                if (template == null)
                {
                    return;
                }

                int level = 1;
                if (list.TryGetValue("level", out object levelObj))
                {
                    level = Convert.ToInt32(levelObj);
                }

                if (level < 1)
                {
                    level = 1;
                }

                Word.ListLevel listLevel = template.ListLevels[level];
                string numberStyle = Enum.GetName(typeof(Word.WdListNumberStyle), listLevel.NumberStyle);
                if (!string.IsNullOrEmpty(numberStyle))
                {
                    list["number_style"] = numberStyle;
                }

                if (kind == "bullet" || kind == "picture_bullet")
                {
                    if (!string.IsNullOrEmpty(listLevel.NumberFormat))
                    {
                        list["bullet_char"] = listLevel.NumberFormat;
                    }
                }
                else if (!string.IsNullOrEmpty(listLevel.NumberFormat))
                {
                    list["number_format"] = listLevel.NumberFormat;
                }

                list["start_at"] = listLevel.StartAt;
            }
            catch (Exception ex)
            {
                DbgLog("list level 读取失败: " + ex.Message);
            }
        }

        private static void TryReadDropCap(Dictionary<string, object> format, Word.Range range)
        {
            try
            {
                if (range.Paragraphs == null || range.Paragraphs.Count < 1)
                {
                    return;
                }

                Word.DropCap drop = range.Paragraphs[1].DropCap;
                if (drop == null)
                {
                    return;
                }

                var map = new Dictionary<string, object>
                {
                    ["position"] = MapDropPosition(drop.Position),
                };
                map["lines"] = drop.LinesToDrop;
                if (!string.IsNullOrEmpty(drop.FontName))
                {
                    map["font_name"] = drop.FontName;
                }

                map["distance"] = RoundPt(drop.DistanceFromText);
                format["drop_cap"] = map;
            }
            catch (Exception ex)
            {
                DbgLog("drop_cap 读取失败: " + ex.Message);
            }
        }

        private static string MapListKind(Word.WdListType listType)
        {
            switch (listType)
            {
                case Word.WdListType.wdListBullet:
                    return "bullet";
                case Word.WdListType.wdListPictureBullet:
                    return "picture_bullet";
                case Word.WdListType.wdListSimpleNumbering:
                    return "number";
                case Word.WdListType.wdListOutlineNumbering:
                    return "outline";
                case Word.WdListType.wdListMixedNumbering:
                    return "mixed";
                case Word.WdListType.wdListListNumOnly:
                    return "list_num_only";
                default:
                    return null;
            }
        }

        private static string MapDropPosition(Word.WdDropPosition position)
        {
            switch (position)
            {
                case Word.WdDropPosition.wdDropNormal:
                    return "drop";
                case Word.WdDropPosition.wdDropMargin:
                    return "margin";
                default:
                    return "none";
            }
        }

        private static string MapAlignment(Word.WdParagraphAlignment alignment)
        {
            if ((int)alignment == ParaFormatCatalog.WdUndefined)
            {
                return null;
            }

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
                case Word.WdParagraphAlignment.wdAlignParagraphJustifyMed:
                    return "justify_med";
                case Word.WdParagraphAlignment.wdAlignParagraphJustifyHi:
                    return "justify_hi";
                case Word.WdParagraphAlignment.wdAlignParagraphJustifyLow:
                    return "justify_low";
                case Word.WdParagraphAlignment.wdAlignParagraphThaiJustify:
                    return "thai_justify";
                case Word.WdParagraphAlignment.wdAlignParagraphLeft:
                    return "left";
                default:
                    return null;
            }
        }

        private static void TryRead(Dictionary<string, object> format, string key, Func<string> read)
        {
            try
            {
                string value = read();
                if (!string.IsNullOrEmpty(value))
                {
                    format[key] = value;
                }
            }
            catch (Exception ex)
            {
                DbgLog(key + " 读取失败: " + ex.Message);
            }
        }

        private static void TryReadFloat(Dictionary<string, object> format, string key, Func<float> read)
        {
            try
            {
                float value = read();
                if (Math.Abs(value - ParaFormatCatalog.WdUndefined) < 1f)
                {
                    return;
                }

                format[key] = RoundPt(value);
            }
            catch (Exception ex)
            {
                DbgLog(key + " 读取失败: " + ex.Message);
            }
        }

        private static void TryReadFarEastSpacing(Dictionary<string, object> format, Word.Range range, Word.ParagraphFormat pf)
        {
            string xml = "";
            try
            {
                if (range.Paragraphs != null && range.Paragraphs.Count > 0)
                {
                    xml = range.Paragraphs[1].Range.WordOpenXML ?? "";
                }
            }
            catch (Exception ex)
            {
                DbgLog("中西文间距 XML 读取失败: " + ex.Message);
            }

            ReadFarEastOne(format, "add_space_between_far_east_and_alpha", () => pf.AddSpaceBetweenFarEastAndAlpha, xml, "autoSpaceDE");
            ReadFarEastOne(format, "add_space_between_far_east_and_digit", () => pf.AddSpaceBetweenFarEastAndDigit, xml, "autoSpaceDN");
        }

        private static void ReadFarEastOne(
            Dictionary<string, object> format,
            string key,
            Func<int> readCom,
            string xml,
            string tag)
        {
            bool? fromXml = ReadPPrBool(xml, tag);
            if (fromXml.HasValue)
            {
                format[key] = fromXml.Value;
                return;
            }

            TryReadBool(format, key, readCom);
        }

        private static bool? ReadPPrBool(string xml, string tag)
        {
            if (string.IsNullOrEmpty(xml))
            {
                return null;
            }

            Match match = Regex.Match(xml, "<w:" + tag + "\\b[^>]*/>");
            if (!match.Success)
            {
                return null;
            }

            Match val = Regex.Match(match.Value, "w:val=\"(\\d+)\"");
            if (!val.Success)
            {
                return true;
            }

            return val.Groups[1].Value != "0";
        }

        private static void TryReadBool(Dictionary<string, object> format, string key, Func<int> read)
        {
            try
            {
                int value = read();
                if (value == ParaFormatCatalog.WdUndefined || value == ParaFormatCatalog.WdToggle)
                {
                    return;
                }

                format[key] = value != 0;
            }
            catch (Exception ex)
            {
                DbgLog(key + " 读取失败: " + ex.Message);
            }
        }

        private static void TryReadEnum<T>(Dictionary<string, object> format, string key, Func<T> read)
            where T : struct
        {
            try
            {
                T value = read();
                int number = Convert.ToInt32(value);
                if (number == ParaFormatCatalog.WdUndefined || number == ParaFormatCatalog.WdToggle)
                {
                    return;
                }

                string name = Enum.GetName(typeof(T), value);
                if (!string.IsNullOrEmpty(name))
                {
                    format[key] = name;
                }
            }
            catch (Exception ex)
            {
                DbgLog(key + " 读取失败: " + ex.Message);
            }
        }

        private static void TryPutEnum<T>(Dictionary<string, object> map, string key, Func<T> read)
            where T : struct
        {
            try
            {
                T value = read();
                string name = Enum.GetName(typeof(T), value);
                if (!string.IsNullOrEmpty(name))
                {
                    map[key] = name;
                }
            }
            catch (Exception ex)
            {
                DbgLog(key + " 读取失败: " + ex.Message);
            }
        }

        private static void TryPutColor(Dictionary<string, object> map, string key, Func<Word.WdColor> read)
        {
            try
            {
                map[key] = ColorToRgb(read());
            }
            catch (Exception ex)
            {
                DbgLog(key + " 读取失败: " + ex.Message);
            }
        }

        private static string ColorToRgb(Word.WdColor wdColor)
        {
            if (wdColor == Word.WdColor.wdColorAutomatic || (int)wdColor == -16777216)
            {
                return "automatic";
            }

            int colorValue = (int)wdColor;
            int b = (colorValue & 0xFF0000) >> 16;
            int g = (colorValue & 0x00FF00) >> 8;
            int r = colorValue & 0x0000FF;
            return $"#{r:X2}{g:X2}{b:X2}";
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
            string normalizedRule = MapLineSpacingRule(rule);
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

        private static string MapLineSpacingRule(string rule)
        {
            switch ((rule ?? "").Trim().ToLowerInvariant())
            {
                case "":
                case "single":
                case "wdlinespacesingle":
                    return "single";
                case "1.5":
                case "wdlinespace1pt5":
                    return "1.5";
                case "double":
                case "wdlinespacedouble":
                    return "double";
                case "multiple":
                case "wdlinespacemultiple":
                    return "multiple";
                case "exact":
                case "wdlinespaceexactly":
                    return "exact";
                case "at_least":
                case "wdlinespaceatleast":
                    return "at_least";
                default:
                    return rule ?? "single";
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
