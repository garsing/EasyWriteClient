using System;
using System.Collections.Generic;
using System.Globalization;
using Office = Microsoft.Office.Core;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    public static partial class FormatInheritHelper
    {
        [ThreadStatic]
        private static List<Dictionary<string, object>> t_formatWarnings;

        public static void ResetFormatWarnings()
        {
            t_formatWarnings = new List<Dictionary<string, object>>();
        }

        public static void NoteFormatWarning(string key, string message)
        {
            if (t_formatWarnings == null)
            {
                t_formatWarnings = new List<Dictionary<string, object>>();
            }

            t_formatWarnings.Add(new Dictionary<string, object>
            {
                ["key"] = key ?? "",
                ["message"] = message ?? ""
            });
        }

        public static List<Dictionary<string, object>> TakeFormatWarnings()
        {
            List<Dictionary<string, object>> found = t_formatWarnings;
            t_formatWarnings = null;
            if (found == null || found.Count == 0)
            {
                return null;
            }

            return found;
        }

        private static Dictionary<string, object> ReadMergedCharFormat(
            Word.Range range,
            out int kindCount,
            out string note)
        {
            var runs = new List<Dictionary<string, object>>();
            Word.Range trimmed = TrimTrailingMarks(range) ?? range;
            if (trimmed != null && trimmed.End > trimmed.Start && trimmed.Document != null)
            {
                Word.Document doc = trimmed.Document;
                int pos = trimmed.Start;
                int end = trimmed.End;
                while (pos < end)
                {
                    Word.Range ch = doc.Range(pos, pos + 1);
                    if (IsSkippedChar(SafeCharText(ch)))
                    {
                        pos++;
                        continue;
                    }

                    Dictionary<string, object> snap = ReadOneCharFormat(ch);
                    int runEnd = pos + 1;
                    while (runEnd < end)
                    {
                        Word.Range next = doc.Range(runEnd, runEnd + 1);
                        if (IsSkippedChar(SafeCharText(next)))
                        {
                            break;
                        }

                        Dictionary<string, object> nextSnap = ReadOneCharFormat(next);
                        if (!CharFormatCatalog.SameSnapshot(snap, nextSnap))
                        {
                            break;
                        }

                        runEnd++;
                    }

                    runs.Add(snap);
                    pos = runEnd;
                }
            }

            CharFormatCatalog.MergeRunSnapshots(runs, out Dictionary<string, object> merged, out kindCount, out note);
            return merged;
        }

        private static bool IsSkippedChar(string text)
        {
            if (string.IsNullOrEmpty(text))
            {
                return true;
            }

            char c = text[0];
            return c == '\r' || c == '\n' || c == '\a' || c == '\f' || c == '\v' || c == '\x01' || c == '\x07';
        }

        private static string SafeCharText(Word.Range range)
        {
            try
            {
                return range?.Text ?? "";
            }
            catch
            {
                return "";
            }
        }

        private static Dictionary<string, object> ReadOneCharFormat(Word.Range range)
        {
            var format = new Dictionary<string, object>();
            Word.Font font = null;
            try
            {
                font = range.Font;
            }
            catch
            {
                font = null;
            }

            if (font != null)
            {
                Put(format, "font_name", () =>
                {
                    string name = GetEffectiveFontName(font);
                    return string.IsNullOrWhiteSpace(name) ? null : name.Trim();
                });
                PutString(format, "name", () => font.Name);
                PutString(format, "name_ascii", () => font.NameAscii);
                PutString(format, "name_far_east", () => font.NameFarEast);
                PutString(format, "name_bi", () => font.NameBi);
                PutString(format, "name_other", () => font.NameOther);
                PutFloat(format, "font_size", () => font.Size);
                PutFloat(format, "size_bi", () => font.SizeBi);
                PutRgb(format, "font_color", () => font.Color);
                PutRgb(format, "underline_color", () => font.UnderlineColor);
                PutRgb(format, "diacritic_color", () => font.DiacriticColor);
                PutEnum(format, "color_index", () => font.ColorIndex);
                PutEnum(format, "color_index_bi", () => font.ColorIndexBi);
                Put(format, "text_color", () => ReadColorFormat(font.TextColor));
                PutTri(format, "is_bold", () => font.Bold);
                PutTri(format, "bold_bi", () => font.BoldBi);
                PutTri(format, "is_italic", () => font.Italic);
                PutTri(format, "italic_bi", () => font.ItalicBi);
                Put(format, "has_underline", () =>
                {
                    int raw = (int)font.Underline;
                    if (IsUndefinedInt(raw))
                    {
                        return null;
                    }

                    return (object)(font.Underline != Word.WdUnderline.wdUnderlineNone);
                });
                PutEnum(format, "underline", () => font.Underline);
                PutTri(format, "has_strikethrough", () => font.StrikeThrough);
                PutTri(format, "double_strikethrough", () => font.DoubleStrikeThrough);
                PutTri(format, "is_superscript", () => font.Superscript);
                PutTri(format, "is_subscript", () => font.Subscript);
                PutTri(format, "shadow", () => font.Shadow);
                PutTri(format, "outline", () => font.Outline);
                PutTri(format, "emboss", () => font.Emboss);
                PutTri(format, "engrave", () => font.Engrave);
                PutTri(format, "small_caps", () => font.SmallCaps);
                PutTri(format, "all_caps", () => font.AllCaps);
                PutEnum(format, "emphasis_mark", () => font.EmphasisMark);
                PutFloat(format, "scaling", () => (float)font.Scaling);
                PutFloat(format, "spacing", () => font.Spacing);
                PutFloat(format, "position", () => (float)font.Position);
                PutFloat(format, "kerning", () => font.Kerning);
                Put(format, "disable_character_space_grid", () => (object)font.DisableCharacterSpaceGrid);
                PutEnum(format, "ligatures", () => font.Ligatures);
                PutEnum(format, "number_form", () => font.NumberForm);
                PutEnum(format, "number_spacing", () => font.NumberSpacing);
                PutEnum(format, "stylistic_set", () => font.StylisticSet);
                PutTri(format, "contextual_alternates", () => font.ContextualAlternates);
                Put(format, "shading", () => ReadShading(font));
                Put(format, "borders", () => ReadBorders(font));
                Put(format, "fill", () => ReadFill(font.Fill));
                Put(format, "glow", () => ReadGlow(font.Glow));
                Put(format, "line", () => ReadLine(font.Line));
                Put(format, "reflection", () => ReadReflection(font.Reflection));
                Put(format, "text_shadow", () => ReadShadow(font.TextShadow));
            }

            PutRgb(format, "background_color", () => ReadCharBackground(range));
            PutEnum(format, "highlight_color_index", () => range.HighlightColorIndex);
            PutEnum(format, "language_id", () => range.LanguageID);
            PutEnum(format, "language_id_far_east", () => range.LanguageIDFarEast);
            PutTri(format, "no_proofing", () => range.NoProofing);
            PutFloat(format, "fit_text_width", () => range.FitTextWidth);
            CharFormatCatalog.OmitReadDefaults(format);
            return format;
        }

        private static void ApplyCharFormatKey(Word.Range range, string key, object value)
        {
            Word.Font font = range.Font;
            switch (key)
            {
                case "font_name":
                    ApplyFontName(range, value?.ToString() ?? "");
                    return;
                case "name":
                    font.Name = value?.ToString() ?? "";
                    return;
                case "name_ascii":
                    font.NameAscii = value?.ToString() ?? "";
                    return;
                case "name_far_east":
                    font.NameFarEast = value?.ToString() ?? "";
                    return;
                case "name_bi":
                    font.NameBi = value?.ToString() ?? "";
                    return;
                case "name_other":
                    font.NameOther = value?.ToString() ?? "";
                    return;
                case "font_size":
                    AssignPositive(number => font.Size = number, value);
                    return;
                case "size_bi":
                    AssignPositive(number => font.SizeBi = number, value);
                    return;
                case "font_color":
                    font.Color = ConvertRgbToWdColor(value?.ToString() ?? "");
                    return;
                case "underline_color":
                    font.UnderlineColor = ConvertRgbToWdColor(value?.ToString() ?? "");
                    return;
                case "diacritic_color":
                    font.DiacriticColor = ConvertRgbToWdColor(value?.ToString() ?? "");
                    return;
                case "color_index":
                    font.ColorIndex = ParseEnum<Word.WdColorIndex>(value);
                    return;
                case "color_index_bi":
                    font.ColorIndexBi = ParseEnum<Word.WdColorIndex>(value);
                    return;
                case "text_color":
                    WriteColorFormat(font.TextColor, value);
                    return;
                case "is_bold":
                    font.Bold = ReadBool(value) ? 1 : 0;
                    return;
                case "bold_bi":
                    font.BoldBi = TriInt(value);
                    return;
                case "is_italic":
                    font.Italic = TriInt(value);
                    return;
                case "italic_bi":
                    font.ItalicBi = TriInt(value);
                    return;
                case "has_underline":
                    font.Underline = ReadBool(value)
                        ? Word.WdUnderline.wdUnderlineSingle
                        : Word.WdUnderline.wdUnderlineNone;
                    return;
                case "underline":
                    font.Underline = ParseEnum<Word.WdUnderline>(value);
                    return;
                case "has_strikethrough":
                    font.StrikeThrough = ReadBool(value) ? 1 : 0;
                    return;
                case "double_strikethrough":
                    font.DoubleStrikeThrough = TriInt(value);
                    return;
                case "is_subscript":
                    font.Subscript = TriInt(value);
                    return;
                case "is_superscript":
                    font.Superscript = TriInt(value);
                    return;
                case "shadow":
                    font.Shadow = TriInt(value);
                    return;
                case "outline":
                    font.Outline = TriInt(value);
                    return;
                case "emboss":
                    font.Emboss = TriInt(value);
                    return;
                case "engrave":
                    font.Engrave = TriInt(value);
                    return;
                case "small_caps":
                    font.SmallCaps = TriInt(value);
                    return;
                case "all_caps":
                    font.AllCaps = TriInt(value);
                    return;
                case "emphasis_mark":
                    font.EmphasisMark = ParseEnum<Word.WdEmphasisMark>(value);
                    return;
                case "scaling":
                    font.Scaling = (int)Math.Round(ParseFloat(value));
                    return;
                case "spacing":
                    font.Spacing = ParseFloat(value);
                    return;
                case "position":
                    font.Position = (int)Math.Round(ParseFloat(value));
                    return;
                case "kerning":
                    font.Kerning = ParseFloat(value);
                    return;
                case "disable_character_space_grid":
                    font.DisableCharacterSpaceGrid = ReadBool(value);
                    return;
                case "ligatures":
                    font.Ligatures = ParseEnum<Word.WdLigatures>(value);
                    return;
                case "number_form":
                    font.NumberForm = ParseEnum<Word.WdNumberForm>(value);
                    return;
                case "number_spacing":
                    font.NumberSpacing = ParseEnum<Word.WdNumberSpacing>(value);
                    return;
                case "stylistic_set":
                    font.StylisticSet = ParseEnum<Word.WdStylisticSet>(value);
                    return;
                case "contextual_alternates":
                    font.ContextualAlternates = TriInt(value);
                    return;
                case "background_color":
                    WriteCharBackground(range, ConvertRgbToWdColor(value?.ToString() ?? ""));
                    return;
                case "shading":
                    WriteShading(font, value);
                    return;
                case "highlight_color_index":
                    range.HighlightColorIndex = ParseEnum<Word.WdColorIndex>(value);
                    return;
                case "borders":
                    WriteBorders(font, value);
                    return;
                case "fill":
                    WriteFill(font.Fill, value);
                    return;
                case "glow":
                    WriteGlow(font.Glow, value);
                    return;
                case "line":
                    WriteLine(font.Line, value);
                    return;
                case "reflection":
                    WriteReflection(font.Reflection, value);
                    return;
                case "text_shadow":
                    WriteShadow(font.TextShadow, value);
                    return;
                case "language_id":
                    range.LanguageID = ParseEnum<Word.WdLanguageID>(value);
                    return;
                case "language_id_far_east":
                    range.LanguageIDFarEast = ParseEnum<Word.WdLanguageID>(value);
                    return;
                case "no_proofing":
                    range.NoProofing = TriInt(value);
                    return;
                case "fit_text_width":
                    range.FitTextWidth = ParseFloat(value);
                    return;
            }
        }

        private static void AssignPositive(Action<float> assign, object value)
        {
            float number = ParseFloat(value);
            if (number > 0)
            {
                assign(number);
            }
        }

        private static void Put(Dictionary<string, object> format, string key, Func<object> read)
        {
            try
            {
                object value = read();
                if (value == null || CharFormatCatalog.IsSentinel(value))
                {
                    return;
                }

                format[key] = value;
            }
            catch
            {
            }
        }

        private static void PutString(Dictionary<string, object> format, string key, Func<string> read)
        {
            Put(format, key, () =>
            {
                string text = read();
                return string.IsNullOrWhiteSpace(text) ? null : text.Trim();
            });
        }

        private static void PutFloat(Dictionary<string, object> format, string key, Func<float> read)
        {
            Put(format, key, () =>
            {
                float number = read();
                if (IsUndefinedInt((int)Math.Round(number)))
                {
                    return null;
                }

                return (object)number;
            });
        }

        private static void PutRgb(Dictionary<string, object> format, string key, Func<Word.WdColor> read)
        {
            Put(format, key, () =>
            {
                Word.WdColor color = read();
                if (IsUndefinedInt((int)color))
                {
                    return null;
                }

                return ConvertWdColorToRgb(color);
            });
        }

        private static void PutTri(Dictionary<string, object> format, string key, Func<int> read)
        {
            Put(format, key, () =>
            {
                int raw = read();
                if (IsUndefinedInt(raw))
                {
                    return null;
                }

                return (object)(raw != 0);
            });
        }

        private static void PutEnum(Dictionary<string, object> format, string key, Func<object> read)
        {
            Put(format, key, () =>
            {
                object value = read();
                if (value == null)
                {
                    return null;
                }

                int raw = Convert.ToInt32(value, CultureInfo.InvariantCulture);
                if (IsUndefinedInt(raw))
                {
                    return null;
                }

                return value.ToString();
            });
        }

        private static bool IsUndefinedInt(int value)
        {
            return value == CharFormatCatalog.WdUndefined || value == CharFormatCatalog.WdToggle;
        }

        private static int TriInt(object value)
        {
            return ReadBool(value) ? -1 : 0;
        }

        private static bool ReadBool(object value)
        {
            if (value is bool flag)
            {
                return flag;
            }

            string text = value?.ToString();
            return text == "true" || text == "True" || text == "1" || text == "-1";
        }

        private static float ParseFloat(object value)
        {
            if (value is float f)
            {
                return f;
            }

            if (value is double d)
            {
                return (float)d;
            }

            if (value is int i)
            {
                return i;
            }

            return float.Parse(value?.ToString() ?? "0", CultureInfo.InvariantCulture);
        }

        private static T ParseEnum<T>(object value) where T : struct
        {
            return (T)Enum.Parse(typeof(T), value?.ToString() ?? "", true);
        }

        private static Dictionary<string, object> MapOf(object value)
        {
            return CharFormatCatalog.AsStringDict(value);
        }

        private static object ReadColorFormat(Word.ColorFormat color)
        {
            if (color == null)
            {
                return null;
            }

            var map = new Dictionary<string, object>();
            try
            {
                int raw = color.RGB;
                if (!IsUndefinedInt(raw))
                {
                    map["rgb"] = ConvertWdColorToRgb((Word.WdColor)raw);
                }
            }
            catch
            {
            }

            try
            {
                map["type"] = color.Type.ToString();
            }
            catch
            {
            }

            try
            {
                Word.WdThemeColorIndex theme = color.ObjectThemeColor;
                int raw = (int)theme;
                if (!IsUndefinedInt(raw) && raw > 0)
                {
                    map["object_theme_color"] = theme.ToString();
                }
            }
            catch
            {
            }

            try
            {
                float brightness = color.Brightness;
                if (!IsUndefinedInt((int)Math.Round(brightness)))
                {
                    map["brightness"] = brightness;
                }
            }
            catch
            {
            }

            return map.Count == 0 ? null : map;
        }

        private static void WriteColorFormat(Word.ColorFormat color, object value)
        {
            Dictionary<string, object> map = MapOf(value);
            if (color == null || map == null)
            {
                return;
            }

            if (map.TryGetValue("object_theme_color", out object theme) && theme != null
                && !CharFormatCatalog.IsSentinel(theme))
            {
                color.ObjectThemeColor = ParseEnum<Word.WdThemeColorIndex>(theme);
                if (map.TryGetValue("brightness", out object brightness) && brightness != null)
                {
                    color.Brightness = ParseFloat(brightness);
                }

                return;
            }

            if (map.TryGetValue("rgb", out object rgb) && rgb != null)
            {
                color.RGB = (int)ConvertRgbToWdColor(rgb.ToString());
            }
        }

        private static object ReadShading(Word.Font font)
        {
            var map = new Dictionary<string, object>();
            try
            {
                Word.WdTextureIndex texture = font.Shading.Texture;
                if (!IsUndefinedInt((int)texture))
                {
                    map["texture"] = texture.ToString();
                }
            }
            catch
            {
            }

            try
            {
                map["background_pattern_color"] = ConvertWdColorToRgb(font.Shading.BackgroundPatternColor);
            }
            catch
            {
            }

            try
            {
                map["foreground_pattern_color"] = ConvertWdColorToRgb(font.Shading.ForegroundPatternColor);
            }
            catch
            {
            }

            return map.Count == 0 ? null : map;
        }

        private static void WriteShading(Word.Font font, object value)
        {
            Dictionary<string, object> map = MapOf(value);
            if (map == null)
            {
                return;
            }

            if (map.TryGetValue("texture", out object texture) && texture != null)
            {
                font.Shading.Texture = ParseEnum<Word.WdTextureIndex>(texture);
            }

            if (map.TryGetValue("background_pattern_color", out object back) && back != null)
            {
                font.Shading.BackgroundPatternColor = ConvertRgbToWdColor(back.ToString());
            }

            if (map.TryGetValue("foreground_pattern_color", out object fore) && fore != null)
            {
                font.Shading.ForegroundPatternColor = ConvertRgbToWdColor(fore.ToString());
            }
        }

        private static object ReadBorders(Word.Font font)
        {
            var map = new Dictionary<string, object>();
            try
            {
                map["outside_line_style"] = font.Borders.OutsideLineStyle.ToString();
            }
            catch
            {
            }

            try
            {
                map["outside_line_width"] = font.Borders.OutsideLineWidth.ToString();
            }
            catch
            {
            }

            try
            {
                map["outside_color"] = ConvertWdColorToRgb(font.Borders.OutsideColor);
            }
            catch
            {
            }

            return map.Count == 0 ? null : map;
        }

        private static void WriteBorders(Word.Font font, object value)
        {
            Dictionary<string, object> map = MapOf(value);
            if (map == null)
            {
                return;
            }

            if (map.TryGetValue("outside_line_style", out object style) && style != null)
            {
                font.Borders.OutsideLineStyle = ParseEnum<Word.WdLineStyle>(style);
            }

            if (map.TryGetValue("outside_line_width", out object width) && width != null)
            {
                font.Borders.OutsideLineWidth = ParseEnum<Word.WdLineWidth>(width);
            }

            if (map.TryGetValue("outside_color", out object color) && color != null)
            {
                font.Borders.OutsideColor = ConvertRgbToWdColor(color.ToString());
            }
        }

        private static object ReadFill(Word.FillFormat fill)
        {
            if (fill == null)
            {
                return null;
            }

            var map = new Dictionary<string, object>();
            PutTriState(map, "visible", () => fill.Visible);
            PutPlainFloat(map, "transparency", () => fill.Transparency);
            try
            {
                object color = ReadColorFormat(fill.ForeColor);
                if (color != null)
                {
                    map["fore_color"] = color;
                }
            }
            catch
            {
            }

            return map.Count == 0 ? null : map;
        }

        private static void WriteFill(Word.FillFormat fill, object value)
        {
            Dictionary<string, object> map = MapOf(value);
            if (fill == null || map == null)
            {
                return;
            }

            if (map.ContainsKey("visible"))
            {
                fill.Visible = ReadBool(map["visible"]) ? Office.MsoTriState.msoTrue : Office.MsoTriState.msoFalse;
            }

            if (map.TryGetValue("transparency", out object transparency) && transparency != null)
            {
                fill.Transparency = ParseFloat(transparency);
            }

            if (map.TryGetValue("fore_color", out object color))
            {
                WriteColorFormat(fill.ForeColor, color);
            }
        }

        private static object ReadGlow(Word.GlowFormat glow)
        {
            if (glow == null)
            {
                return null;
            }

            var map = new Dictionary<string, object>();
            PutPlainFloat(map, "radius", () => glow.Radius);
            PutPlainFloat(map, "transparency", () => glow.Transparency);
            try
            {
                object color = ReadColorFormat(glow.Color);
                if (color != null)
                {
                    map["color"] = color;
                }
            }
            catch
            {
            }

            return map.Count == 0 ? null : map;
        }

        private static void WriteGlow(Word.GlowFormat glow, object value)
        {
            Dictionary<string, object> map = MapOf(value);
            if (glow == null || map == null)
            {
                return;
            }

            if (map.TryGetValue("radius", out object radius) && radius != null)
            {
                glow.Radius = ParseFloat(radius);
            }

            if (map.TryGetValue("transparency", out object transparency) && transparency != null)
            {
                glow.Transparency = ParseFloat(transparency);
            }

            if (map.TryGetValue("color", out object color))
            {
                WriteColorFormat(glow.Color, color);
            }
        }

        private static object ReadLine(Word.LineFormat line)
        {
            if (line == null)
            {
                return null;
            }

            var map = new Dictionary<string, object>();
            PutTriState(map, "visible", () => line.Visible);
            PutPlainFloat(map, "weight", () => line.Weight);
            PutPlainFloat(map, "transparency", () => line.Transparency);
            try
            {
                map["dash_style"] = line.DashStyle.ToString();
            }
            catch
            {
            }

            try
            {
                object color = ReadColorFormat(line.ForeColor);
                if (color != null)
                {
                    map["fore_color"] = color;
                }
            }
            catch
            {
            }

            return map.Count == 0 ? null : map;
        }

        private static void WriteLine(Word.LineFormat line, object value)
        {
            Dictionary<string, object> map = MapOf(value);
            if (line == null || map == null)
            {
                return;
            }

            if (map.ContainsKey("visible"))
            {
                line.Visible = ReadBool(map["visible"]) ? Office.MsoTriState.msoTrue : Office.MsoTriState.msoFalse;
            }

            if (map.TryGetValue("weight", out object weight) && weight != null)
            {
                line.Weight = ParseFloat(weight);
            }

            if (map.TryGetValue("transparency", out object transparency) && transparency != null)
            {
                line.Transparency = ParseFloat(transparency);
            }

            if (map.TryGetValue("dash_style", out object dash) && dash != null)
            {
                line.DashStyle = ParseEnum<Office.MsoLineDashStyle>(dash);
            }

            if (map.TryGetValue("fore_color", out object color))
            {
                WriteColorFormat(line.ForeColor, color);
            }
        }

        private static object ReadReflection(Word.ReflectionFormat reflection)
        {
            if (reflection == null)
            {
                return null;
            }

            var map = new Dictionary<string, object>();
            try
            {
                map["type"] = reflection.Type.ToString();
            }
            catch
            {
            }

            PutPlainFloat(map, "transparency", () => reflection.Transparency);
            PutPlainFloat(map, "size", () => reflection.Size);
            return map.Count == 0 ? null : map;
        }

        private static void WriteReflection(Word.ReflectionFormat reflection, object value)
        {
            Dictionary<string, object> map = MapOf(value);
            if (reflection == null || map == null)
            {
                return;
            }

            if (map.TryGetValue("type", out object type) && type != null)
            {
                reflection.Type = ParseEnum<Office.MsoReflectionType>(type);
            }

            if (map.TryGetValue("transparency", out object transparency) && transparency != null)
            {
                reflection.Transparency = ParseFloat(transparency);
            }

            if (map.TryGetValue("size", out object size) && size != null)
            {
                reflection.Size = ParseFloat(size);
            }
        }

        private static object ReadShadow(Word.ShadowFormat shadow)
        {
            if (shadow == null)
            {
                return null;
            }

            var map = new Dictionary<string, object>();
            PutTriState(map, "visible", () => shadow.Visible);
            try
            {
                map["type"] = shadow.Type.ToString();
            }
            catch
            {
            }

            PutPlainFloat(map, "transparency", () => shadow.Transparency);
            PutPlainFloat(map, "offset_x", () => shadow.OffsetX);
            PutPlainFloat(map, "offset_y", () => shadow.OffsetY);
            try
            {
                object color = ReadColorFormat(shadow.ForeColor);
                if (color != null)
                {
                    map["fore_color"] = color;
                }
            }
            catch
            {
            }

            return map.Count == 0 ? null : map;
        }

        private static void WriteShadow(Word.ShadowFormat shadow, object value)
        {
            Dictionary<string, object> map = MapOf(value);
            if (shadow == null || map == null)
            {
                return;
            }

            if (map.ContainsKey("visible"))
            {
                shadow.Visible = ReadBool(map["visible"]) ? Office.MsoTriState.msoTrue : Office.MsoTriState.msoFalse;
            }

            if (map.TryGetValue("type", out object type) && type != null)
            {
                shadow.Type = ParseEnum<Office.MsoShadowType>(type);
            }

            if (map.TryGetValue("transparency", out object transparency) && transparency != null)
            {
                shadow.Transparency = ParseFloat(transparency);
            }

            if (map.TryGetValue("offset_x", out object x) && x != null)
            {
                shadow.OffsetX = ParseFloat(x);
            }

            if (map.TryGetValue("offset_y", out object y) && y != null)
            {
                shadow.OffsetY = ParseFloat(y);
            }

            if (map.TryGetValue("fore_color", out object color))
            {
                WriteColorFormat(shadow.ForeColor, color);
            }
        }

        private static void PutTriState(Dictionary<string, object> map, string key, Func<Office.MsoTriState> read)
        {
            try
            {
                int raw = (int)read();
                if (IsUndefinedInt(raw) || raw == (int)Office.MsoTriState.msoTriStateMixed)
                {
                    return;
                }

                map[key] = raw != 0;
            }
            catch
            {
            }
        }

        private static void PutPlainFloat(Dictionary<string, object> map, string key, Func<float> read)
        {
            try
            {
                float number = read();
                if (IsUndefinedInt((int)Math.Round(number)))
                {
                    return;
                }

                map[key] = number;
            }
            catch
            {
            }
        }
    }
}
