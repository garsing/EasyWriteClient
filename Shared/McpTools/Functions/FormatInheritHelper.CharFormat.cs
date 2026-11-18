using System;
using System.Collections.Generic;
using System.Globalization;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;

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
            List<Dictionary<string, object>> list = t_formatWarnings;
            t_formatWarnings = null;
            if (list == null || list.Count == 0)
            {
                return null;
            }
            return list;
        }

        private static void ApplyCharFormatKey(Range range, string key, object value)
        {
            Font font = range.Font;
            switch (key)
            {
            case "font_name":
                ApplyFontName(range, value?.ToString() ?? "");
                break;
            case "name":
                font.Name = value?.ToString() ?? "";
                break;
            case "name_ascii":
                font.NameAscii = value?.ToString() ?? "";
                break;
            case "name_far_east":
                font.NameFarEast = value?.ToString() ?? "";
                break;
            case "name_bi":
                font.NameBi = value?.ToString() ?? "";
                break;
            case "name_other":
                font.NameOther = value?.ToString() ?? "";
                break;
            case "font_size":
                AssignPositive((float number) =>
                {
                    font.Size = number;
                }, value);
                break;
            case "size_bi":
                AssignPositive((float number) =>
                {
                    font.SizeBi = number;
                }, value);
                break;
            case "font_color":
                font.Color = ConvertRgbToWdColor(value?.ToString() ?? "");
                break;
            case "underline_color":
                font.UnderlineColor = ConvertRgbToWdColor(value?.ToString() ?? "");
                break;
            case "diacritic_color":
                WdColor diacritic = ConvertRgbToWdColor(value?.ToString() ?? "");
                try
                {
                    font.DiacriticColor = diacritic;
                    int actual = (int)font.DiacriticColor & 0xFFFFFF;
                    if (actual != ((int)diacritic & 0xFFFFFF))
                    {
                        NoteFormatWarning("diacritic_color", "注音颜色没有留在这段文字上。");
                    }
                }
                catch (Exception ex)
                {
                    NoteFormatWarning("diacritic_color", "注音颜色没有留在这段文字上。" + ex.Message);
                }

                break;
            case "color_index":
                font.ColorIndex = ParseEnum<WdColorIndex>(value);
                break;
            case "color_index_bi":
                WdColorIndex colorIndexBi = ParseEnum<WdColorIndex>(value);
                try
                {
                    font.ColorIndexBi = colorIndexBi;
                    WdColorIndex readBack = font.ColorIndexBi;
                    if (readBack != colorIndexBi)
                    {
                        NoteFormatWarning("color_index_bi", "这段文字没有复杂文种，Word 不单独保留 color_index_bi。");
                    }
                }
                catch (Exception ex)
                {
                    NoteFormatWarning("color_index_bi", "这段文字没有复杂文种，Word 不单独保留 color_index_bi。" + ex.Message);
                }

                break;
            case "text_color":
                WriteColorFormat(font.TextColor, value);
                break;
            case "is_bold":
                font.Bold = (ReadBool(value) ? 1 : 0);
                break;
            case "bold_bi":
                font.BoldBi = TriInt(value);
                break;
            case "is_italic":
                font.Italic = TriInt(value);
                break;
            case "italic_bi":
                font.ItalicBi = TriInt(value);
                break;
            case "has_underline":
                font.Underline = (ReadBool(value) ? WdUnderline.wdUnderlineSingle : WdUnderline.wdUnderlineNone);
                break;
            case "underline":
                font.Underline = ParseEnum<WdUnderline>(value);
                break;
            case "has_strikethrough":
                font.StrikeThrough = (ReadBool(value) ? 1 : 0);
                break;
            case "double_strikethrough":
                font.DoubleStrikeThrough = TriInt(value);
                break;
            case "is_subscript":
                font.Subscript = TriInt(value);
                break;
            case "is_superscript":
                font.Superscript = TriInt(value);
                break;
            case "shadow":
                font.Shadow = TriInt(value);
                break;
            case "outline":
                font.Outline = TriInt(value);
                break;
            case "emboss":
                font.Emboss = TriInt(value);
                break;
            case "engrave":
                font.Engrave = TriInt(value);
                break;
            case "small_caps":
                font.SmallCaps = TriInt(value);
                break;
            case "all_caps":
                font.AllCaps = TriInt(value);
                break;
            case "emphasis_mark":
                font.EmphasisMark = ParseEnum<WdEmphasisMark>(value);
                break;
            case "scaling":
                font.Scaling = (int)Math.Round(ParseFloat(value));
                break;
            case "spacing":
                font.Spacing = ParseFloat(value);
                break;
            case "position":
                font.Position = (int)Math.Round(ParseFloat(value));
                break;
            case "kerning":
                font.Kerning = ParseFloat(value);
                break;
            case "disable_character_space_grid":
                font.DisableCharacterSpaceGrid = ReadBool(value);
                break;
            case "ligatures":
                font.Ligatures = ParseEnum<WdLigatures>(value);
                break;
            case "number_form":
                font.NumberForm = ParseEnum<WdNumberForm>(value);
                break;
            case "number_spacing":
                font.NumberSpacing = ParseEnum<WdNumberSpacing>(value);
                break;
            case "stylistic_set":
                font.StylisticSet = ParseEnum<WdStylisticSet>(value);
                break;
            case "contextual_alternates":
                font.ContextualAlternates = TriInt(value);
                break;
            case "background_color":
                WriteCharBackground(range, ConvertRgbToWdColor(value?.ToString() ?? ""));
                break;
            case "shading":
                WriteShading(font, value);
                break;
            case "highlight_color_index":
                range.HighlightColorIndex = ParseEnum<WdColorIndex>(value);
                break;
            case "borders":
                WriteBorders(font, value);
                break;
            case "fill":
                WriteFill(font.Fill, value);
                break;
            case "glow":
                WriteGlow(font.Glow, value);
                break;
            case "line":
                WriteLine(font.Line, value);
                break;
            case "reflection":
                WriteReflection(font.Reflection, value);
                break;
            case "text_shadow":
                WriteShadow(font.TextShadow, value);
                break;
            case "language_id":
                range.LanguageID = ParseEnum<WdLanguageID>(value);
                break;
            case "language_id_far_east":
                range.LanguageIDFarEast = ParseEnum<WdLanguageID>(value);
                break;
            case "no_proofing":
                range.NoProofing = TriInt(value);
                break;
            case "fit_text_width":
                range.FitTextWidth = ParseFloat(value);
                break;
            }
        }

        private static void AssignPositive(Action<float> assign, object value)
        {
            float num = ParseFloat(value);
            if (num > 0f)
            {
                assign(num);
            }
        }

        private static int TriInt(object value)
        {
            return ReadBool(value) ? (-1) : 0;
        }

        private static bool ReadBool(object value)
        {
            if (value is bool result)
            {
                return result;
            }
            string text = value?.ToString();
            int result2;
            switch (text)
            {
            default:
                result2 = ((text == "-1") ? 1 : 0);
                break;
            case "true":
            case "True":
            case "1":
                result2 = 1;
                break;
            }
            return (byte)result2 != 0;
        }

        private static float ParseFloat(object value)
        {
            if (value is float result)
            {
                return result;
            }
            if (value is double num)
            {
                return (float)num;
            }
            if (value is int num2)
            {
                return num2;
            }
            return float.Parse(value?.ToString() ?? "0", CultureInfo.InvariantCulture);
        }

        private static T ParseEnum<T>(object value) where T : struct
        {
            return (T)Enum.Parse(typeof(T), value?.ToString() ?? "", ignoreCase: true);
        }

        private static Dictionary<string, object> MapOf(object value)
        {
            return CharFormatCatalog.AsStringDict(value);
        }

        private static void WriteColorFormat(Microsoft.Office.Interop.Word.ColorFormat color, object value)
        {
            Dictionary<string, object> dictionary = MapOf(value);
            if (color == null || dictionary == null)
            {
                return;
            }
            object value4;
            if (dictionary.TryGetValue("object_theme_color", out var value2) && value2 != null && !CharFormatCatalog.IsSentinel(value2))
            {
                color.ObjectThemeColor = ParseEnum<WdThemeColorIndex>(value2);
                if (dictionary.TryGetValue("brightness", out var value3) && value3 != null)
                {
                    color.Brightness = ParseFloat(value3);
                }
            }
            else if (dictionary.TryGetValue("rgb", out value4) && value4 != null)
            {
                color.RGB = (int)ConvertRgbToWdColor(value4.ToString());
            }
        }

        private static void WriteShading(Font font, object value)
        {
            Dictionary<string, object> dictionary = MapOf(value);
            if (dictionary != null)
            {
                if (dictionary.TryGetValue("texture", out var value2) && value2 != null)
                {
                    font.Shading.Texture = ParseEnum<WdTextureIndex>(value2);
                }
                if (dictionary.TryGetValue("background_pattern_color", out var value3) && value3 != null)
                {
                    font.Shading.BackgroundPatternColor = ConvertRgbToWdColor(value3.ToString());
                }
                if (dictionary.TryGetValue("foreground_pattern_color", out var value4) && value4 != null)
                {
                    font.Shading.ForegroundPatternColor = ConvertRgbToWdColor(value4.ToString());
                }
            }
        }

        private static void WriteBorders(Font font, object value)
        {
            Dictionary<string, object> dictionary = MapOf(value);
            if (dictionary != null)
            {
                if (dictionary.TryGetValue("outside_line_style", out var value2) && value2 != null)
                {
                    font.Borders.OutsideLineStyle = ParseEnum<WdLineStyle>(value2);
                }
                if (dictionary.TryGetValue("outside_line_width", out var value3) && value3 != null)
                {
                    font.Borders.OutsideLineWidth = ParseEnum<WdLineWidth>(value3);
                }
                if (dictionary.TryGetValue("outside_color", out var value4) && value4 != null)
                {
                    font.Borders.OutsideColor = ConvertRgbToWdColor(value4.ToString());
                }
            }
        }

        private static void WriteFill(Microsoft.Office.Interop.Word.FillFormat fill, object value)
        {
            Dictionary<string, object> dictionary = MapOf(value);
            if (fill != null && dictionary != null)
            {
                if (dictionary.ContainsKey("visible"))
                {
                    fill.Visible = (ReadBool(dictionary["visible"]) ? MsoTriState.msoTrue : MsoTriState.msoFalse);
                }
                if (dictionary.TryGetValue("transparency", out var value2) && value2 != null)
                {
                    fill.Transparency = ParseFloat(value2);
                }
                if (dictionary.TryGetValue("fore_color", out var value3))
                {
                    WriteColorFormat(fill.ForeColor, value3);
                }
            }
        }

        private static void WriteGlow(Microsoft.Office.Interop.Word.GlowFormat glow, object value)
        {
            Dictionary<string, object> dictionary = MapOf(value);
            if (glow != null && dictionary != null)
            {
                if (dictionary.TryGetValue("radius", out var value2) && value2 != null)
                {
                    glow.Radius = ParseFloat(value2);
                }
                if (dictionary.TryGetValue("transparency", out var value3) && value3 != null)
                {
                    glow.Transparency = ParseFloat(value3);
                }
                if (dictionary.TryGetValue("color", out var value4))
                {
                    WriteColorFormat(glow.Color, value4);
                }
            }
        }

        private static void WriteLine(Microsoft.Office.Interop.Word.LineFormat line, object value)
        {
            Dictionary<string, object> dictionary = MapOf(value);
            if (line != null && dictionary != null)
            {
                if (dictionary.ContainsKey("visible"))
                {
                    line.Visible = (ReadBool(dictionary["visible"]) ? MsoTriState.msoTrue : MsoTriState.msoFalse);
                }
                if (dictionary.TryGetValue("weight", out var value2) && value2 != null)
                {
                    line.Weight = ParseFloat(value2);
                }
                if (dictionary.TryGetValue("transparency", out var value3) && value3 != null)
                {
                    line.Transparency = ParseFloat(value3);
                }
                if (dictionary.TryGetValue("dash_style", out var value4) && value4 != null)
                {
                    line.DashStyle = ParseEnum<MsoLineDashStyle>(value4);
                }
                if (dictionary.TryGetValue("fore_color", out var value5))
                {
                    WriteColorFormat(line.ForeColor, value5);
                }
            }
        }

        private static void WriteReflection(Microsoft.Office.Interop.Word.ReflectionFormat reflection, object value)
        {
            Dictionary<string, object> dictionary = MapOf(value);
            TocFormatTrace.Write(
                "reflection enter valueType=" + (value == null ? "null" : value.GetType().FullName)
                + " map=" + (dictionary == null ? "null" : string.Join(",", dictionary.Keys))
                + " reflectionNull=" + (reflection == null));
            if (reflection == null || dictionary == null)
            {
                return;
            }

            MsoReflectionType requestedType = MsoReflectionType.msoReflectionTypeNone;
            bool requestedPreset = false;
            if (dictionary.TryGetValue("type", out var value2) && value2 != null)
            {
                requestedType = ParseEnum<MsoReflectionType>(value2);
                requestedPreset = requestedType != MsoReflectionType.msoReflectionTypeNone
                    && requestedType != MsoReflectionType.msoReflectionTypeMixed;
                string before = SafeReflectionType(reflection);
                try
                {
                    reflection.Type = requestedType;
                    TocFormatTrace.Write("reflection set type " + requestedType + " before=" + before + " after=" + SafeReflectionType(reflection));
                }
                catch (Exception ex)
                {
                    TocFormatTrace.Write("reflection set type failed before=" + before + " " + ex);
                    throw;
                }
            }

            if (dictionary.TryGetValue("transparency", out var value3) && value3 != null)
            {
                try
                {
                    reflection.Transparency = ParseFloat(value3);
                    TocFormatTrace.Write("reflection transparency=" + reflection.Transparency);
                }
                catch (Exception ex)
                {
                    TocFormatTrace.Write("reflection transparency failed " + ex);
                    throw;
                }
            }

            if (dictionary.TryGetValue("size", out var value4) && value4 != null)
            {
                try
                {
                    float requested = ParseFloat(value4);
                    string typeBeforeSize = SafeReflectionType(reflection);
                    float sizeBefore = SafeReflectionSize(reflection);
                    reflection.Size = requested;
                    string typeAfterSize = SafeReflectionType(reflection);
                    TocFormatTrace.Write(
                        "reflection size requested=" + requested
                        + " sizeBefore=" + sizeBefore
                        + " sizeAfter=" + SafeReflectionSize(reflection)
                        + " typeBeforeSize=" + typeBeforeSize
                        + " typeAfterSize=" + typeAfterSize);
                    if (requestedPreset && !string.Equals(typeAfterSize, requestedType.ToString(), StringComparison.Ordinal))
                    {
                        NoteFormatWarning(
                            "reflection",
                            "预设倒影自带固定大小，同时指定大小后预设类型会被清掉。倒影类型没有留下，大小已写入。");
                    }
                }
                catch (Exception ex)
                {
                    TocFormatTrace.Write("reflection size failed " + ex);
                    throw;
                }
            }
        }

        private static float SafeReflectionSize(Microsoft.Office.Interop.Word.ReflectionFormat reflection)
        {
            try
            {
                return reflection.Size;
            }
            catch (Exception)
            {
                return float.NaN;
            }
        }

        private static string SafeReflectionType(Microsoft.Office.Interop.Word.ReflectionFormat reflection)
        {
            try
            {
                return reflection.Type.ToString();
            }
            catch (Exception ex)
            {
                return "read-failed " + ex.GetType().Name + " " + ex.Message;
            }
        }

        private static void WriteShadow(Microsoft.Office.Interop.Word.ShadowFormat shadow, object value)
        {
            Dictionary<string, object> dictionary = MapOf(value);
            if (shadow != null && dictionary != null)
            {
                if (dictionary.ContainsKey("visible"))
                {
                    shadow.Visible = (ReadBool(dictionary["visible"]) ? MsoTriState.msoTrue : MsoTriState.msoFalse);
                }
                if (dictionary.TryGetValue("type", out var value2) && value2 != null)
                {
                    shadow.Type = ParseEnum<MsoShadowType>(value2);
                }
                if (dictionary.TryGetValue("transparency", out var value3) && value3 != null)
                {
                    shadow.Transparency = ParseFloat(value3);
                }
                if (dictionary.TryGetValue("offset_x", out var value4) && value4 != null)
                {
                    shadow.OffsetX = ParseFloat(value4);
                }
                if (dictionary.TryGetValue("offset_y", out var value5) && value5 != null)
                {
                    shadow.OffsetY = ParseFloat(value5);
                }
                if (dictionary.TryGetValue("fore_color", out var value6))
                {
                    WriteColorFormat(shadow.ForeColor, value6);
                }
            }
        }
    }
}
