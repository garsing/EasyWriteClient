using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using WordAddIn1.OpenFiles;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    /// <summary>
    /// 句子字符格式：Trim 后的 Range.WordOpenXML 按格式区间读。不逐字问 COM。
    /// </summary>
    public static class CharFormatRunReader
    {
        private static readonly XNamespace W = OpenXmlPackage.W;
        private static readonly XNamespace W14 = "http://schemas.microsoft.com/office/word/2010/wordml";
        private static readonly XNamespace A = "http://schemas.openxmlformats.org/drawingml/2006/main";
        private static readonly XNamespace Mc = "http://schemas.openxmlformats.org/markup-compatibility/2006";
        private static bool _trackRunOffsets;

        /// <summary>
        /// 读到默认黄突出显示时挂在区间上，不放进 char_format，避免被抄回去写成格式键。
        /// </summary>
        public const string InsertHighlightNote =
            "highlight_color_index 为 wdYellow 时，可能是插入或替换自动加上的突出显示，用来标出写入位置，不要当成异常格式。";

        public static bool TryParse(
            string wordOpenXml,
            out List<Dictionary<string, object>> runs,
            out string joinedText,
            out string error)
        {
            runs = null;
            joinedText = "";
            error = null;
            OpenXmlPackage pkg;
            try
            {
                pkg = OpenXmlPackage.Parse(wordOpenXml);
            }
            catch (Exception ex)
            {
                error = ex.Message;
                return false;
            }

            var styles = LoadStyles(pkg);
            ThemeFonts theme = LoadTheme(pkg);
            var spans = new List<Dictionary<string, object>>();
            var joined = new StringBuilder();
            var fields = new Stack<bool>();
            Walk(pkg.Body, styles, theme, fields, spans, joined);
            runs = spans;
            joinedText = joined.ToString();
            return true;
        }

        public static bool TryCollectBody(
            string wordOpenXml,
            out List<Dictionary<string, object>> runs,
            out string error)
        {
            _trackRunOffsets = true;
            try
            {
                return TryParse(wordOpenXml, out runs, out _, out error);
            }
            finally
            {
                _trackRunOffsets = false;
            }
        }

        public static bool TryRead(Word.Range range, out List<Dictionary<string, object>> runs, out string error)
        {
            runs = null;
            error = null;
            Word.Range trimmed = FormatInheritHelper.TrimTrailingMarks(range);
            if (trimmed == null)
            {
                error = "范围为空";
                return false;
            }

            string xml = ReadWordOpenXml(trimmed, out error);
            if (xml == null)
            {
                return false;
            }

            if (!TryParse(xml, out runs, out string joined, out error))
            {
                runs = null;
                return false;
            }

            string story;
            try
            {
                story = trimmed.Text ?? "";
            }
            catch (Exception ex)
            {
                error = "读取 Range.Text 失败: " + ex.Message;
                runs = null;
                return false;
            }

            if (!string.Equals(joined, story, StringComparison.Ordinal))
            {
                error = "字符格式与正文对不上：区间 " + joined.Length + " 字，Range.Text " + story.Length + " 字";
                runs = null;
                return false;
            }

            if (xml.IndexOf("http://www.wps.cn/", StringComparison.Ordinal) >= 0)
            {
                FillWpsExportGaps(trimmed, runs);
            }
            else
            {
                FillWordClassicBools(trimmed, runs);
            }

            foreach (Dictionary<string, object> format in RunFormats(runs))
            {
                CharFormatCatalog.OmitReadDefaults(format);
            }

            RefreshLayerFinal(runs);
            return true;
        }

        private static void FillWordClassicBools(Word.Range range, List<Dictionary<string, object>> runs)
        {
            bool hasEffect = false;
            foreach (Dictionary<string, object> format in RunFormats(runs))
            {
                if (format.ContainsKey("text_shadow") || format.ContainsKey("line"))
                {
                    hasEffect = true;
                    break;
                }
            }

            if (!hasEffect)
            {
                return;
            }

            bool shadow = false;
            bool emboss = false;
            bool engrave = false;
            try
            {
                Word.Font font = range.Font;
                shadow = ComOn(font.Shadow);
                emboss = ComOn(font.Emboss);
                engrave = ComOn(font.Engrave);
            }
            catch
            {
                return;
            }

            if (!shadow && !emboss && !engrave)
            {
                return;
            }

            foreach (Dictionary<string, object> format in RunFormats(runs))
            {
                if (shadow && !format.ContainsKey("shadow"))
                {
                    format["shadow"] = true;
                }

                if (emboss && !format.ContainsKey("emboss"))
                {
                    format["emboss"] = true;
                }

                if (engrave && !format.ContainsKey("engrave"))
                {
                    format["engrave"] = true;
                }

            }
        }

        private static bool ComOn(object raw)
        {
            if (raw == null)
            {
                return false;
            }

            if (raw is bool flag)
            {
                return flag;
            }

            if (raw is int number)
            {
                return number == -1 || number == 1;
            }

            string text = raw.ToString();
            return text == "-1" || text == "1" || string.Equals(text, "true", StringComparison.OrdinalIgnoreCase);
        }

        private static void FillWpsExportGaps(Word.Range range, List<Dictionary<string, object>> runs)
        {
            if (runs == null || runs.Count == 0)
            {
                return;
            }

            bool needSize = false;
            bool needLigatures = false;
            bool needStylistic = false;
            foreach (Dictionary<string, object> format in RunFormats(runs))
            {
                if (IsZeroSize(format))
                {
                    needSize = true;
                }

                if (format.TryGetValue("ligatures", out object ligatures)
                    && string.Equals(ligatures?.ToString(), "wdLigaturesAll", StringComparison.Ordinal))
                {
                    needLigatures = true;
                }

                if (!format.ContainsKey("stylistic_set"))
                {
                    needStylistic = true;
                }
            }

            if (!needSize && !needLigatures && !needStylistic)
            {
                return;
            }

            float size = 0;
            string ligatureName = "";
            string stylisticName = "";
            Word.Font font = null;
            try
            {
                font = range.Font;
            }
            catch
            {
                return;
            }

            if (needSize)
            {
                try
                {
                    float raw = font.Size;
                    if (raw > 0 && raw < 10000)
                    {
                        size = raw;
                    }
                }
                catch
                {
                }
            }

            if (needLigatures)
            {
                try
                {
                    ligatureName = DefinedEnum(font.Ligatures);
                }
                catch
                {
                }
            }

            if (needStylistic)
            {
                try
                {
                    stylisticName = DefinedEnum(font.StylisticSet);
                }
                catch
                {
                }
            }

            foreach (Dictionary<string, object> format in RunFormats(runs))
            {
                if (size > 0 && IsZeroSize(format))
                {
                    format["font_size"] = size;
                }

                if (!string.IsNullOrEmpty(ligatureName)
                    && format.TryGetValue("ligatures", out object ligatures)
                    && string.Equals(ligatures?.ToString(), "wdLigaturesAll", StringComparison.Ordinal))
                {
                    format["ligatures"] = ligatureName;
                }

                if (!string.IsNullOrEmpty(stylisticName) && !format.ContainsKey("stylistic_set"))
                {
                    format["stylistic_set"] = stylisticName;
                }

            }
        }

        private static IEnumerable<Dictionary<string, object>> RunFormats(List<Dictionary<string, object>> runs)
        {
            foreach (Dictionary<string, object> run in runs)
            {
                if (run != null
                    && run.TryGetValue("char_format", out object raw)
                    && raw is Dictionary<string, object> format)
                {
                    yield return format;
                }
            }
        }

        private static bool IsZeroSize(Dictionary<string, object> format)
        {
            if (!format.TryGetValue("font_size", out object value) || value == null)
            {
                return true;
            }

            return double.TryParse(value.ToString(), NumberStyles.Float, CultureInfo.InvariantCulture, out double number)
                && Math.Abs(number) < 0.001;
        }

        private static string DefinedEnum(object raw)
        {
            if (raw == null)
            {
                return "";
            }

            if (raw is int number && (number == 9999999 || number == 9999998))
            {
                return "";
            }

            string text = raw.ToString();
            if (text == "9999999"
                || text == "9999998"
                || text.IndexOf("Undefined", StringComparison.OrdinalIgnoreCase) >= 0
                || text.IndexOf("Toggle", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return "";
            }

            return text;
        }

        private static string ReadWordOpenXml(Word.Range range, out string error)
        {
            error = null;
            try
            {
                string xml = range.WordOpenXML;
                if (!string.IsNullOrEmpty(xml))
                {
                    return xml;
                }
            }
            catch (Exception ex)
            {
                error = ex.Message;
            }

            try
            {
                string late = Convert.ToString(WpsCom.GetProperty(range, "WordOpenXML") ?? "");
                if (!string.IsNullOrEmpty(late))
                {
                    error = null;
                    return late;
                }
            }
            catch (Exception ex)
            {
                error = ex.Message;
            }

            error = string.IsNullOrEmpty(error) ? "WordOpenXML 为空" : "导出 WordOpenXML 失败: " + error;
            return null;
        }

        private sealed class StyleRec
        {
            public string BasedOn;
            public XElement Rpr;
        }

        private sealed class ThemeFonts
        {
            public Dictionary<string, string> Scheme = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            public string MajorLatin = "";
            public string MinorLatin = "";
            public string MajorEastAsia = "";
            public string MinorEastAsia = "";
            public string MajorBidi = "";
            public string MinorBidi = "";
            public Dictionary<string, string> MajorScript = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            public Dictionary<string, string> MinorScript = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        }

        private static Dictionary<string, StyleRec> LoadStyles(OpenXmlPackage pkg)
        {
            var map = new Dictionary<string, StyleRec>(StringComparer.Ordinal);
            XElement styles = pkg.FindPartRoot("styles.xml", W + "styles");
            if (styles == null)
            {
                return map;
            }

            XElement defaults = styles.Element(W + "docDefaults")?.Element(W + "rPrDefault")?.Element(W + "rPr");
            if (defaults != null)
            {
                map["__docDefaults"] = new StyleRec { Rpr = defaults };
            }

            foreach (XElement style in styles.Elements(W + "style"))
            {
                string id = (string)style.Attribute(W + "styleId");
                if (string.IsNullOrEmpty(id))
                {
                    continue;
                }

                map[id] = new StyleRec
                {
                    BasedOn = (string)style.Element(W + "basedOn")?.Attribute(W + "val") ?? "",
                    Rpr = style.Element(W + "rPr")
                };
            }

            return map;
        }

        private static ThemeFonts LoadTheme(OpenXmlPackage pkg)
        {
            var theme = new ThemeFonts();
            XElement root = pkg.FindPartRoot("theme1.xml", A + "theme");
            if (root == null)
            {
                return theme;
            }

            XElement scheme = root.Element(A + "themeElements")?.Element(A + "clrScheme");
            if (scheme != null)
            {
                foreach (XElement slot in scheme.Elements())
                {
                    string rgb = ColorVal(slot.Element(A + "srgbClr"))
                        ?? (string)slot.Element(A + "sysClr")?.Attribute("lastClr");
                    if (!string.IsNullOrEmpty(rgb))
                    {
                        theme.Scheme[slot.Name.LocalName] = NormalizeRgb(rgb);
                    }
                }
            }

            XElement fonts = root.Element(A + "themeElements")?.Element(A + "fontScheme");
            if (fonts != null)
            {
                theme.MajorLatin = Typeface(fonts, "majorFont", "latin");
                theme.MinorLatin = Typeface(fonts, "minorFont", "latin");
                theme.MajorEastAsia = Typeface(fonts, "majorFont", "ea");
                theme.MinorEastAsia = Typeface(fonts, "minorFont", "ea");
                theme.MajorBidi = Typeface(fonts, "majorFont", "cs");
                theme.MinorBidi = Typeface(fonts, "minorFont", "cs");
                LoadScripts(fonts.Element(A + "majorFont"), theme.MajorScript);
                LoadScripts(fonts.Element(A + "minorFont"), theme.MinorScript);
            }

            return theme;
        }

        private static string Typeface(XElement fonts, string group, string slot)
        {
            return (string)fonts.Element(A + group)?.Element(A + slot)?.Attribute("typeface") ?? "";
        }

        private static void LoadScripts(XElement group, Dictionary<string, string> into)
        {
            if (group == null)
            {
                return;
            }

            foreach (XElement font in group.Elements(A + "font"))
            {
                string script = (string)font.Attribute("script");
                string face = (string)font.Attribute("typeface");
                if (!string.IsNullOrEmpty(script) && !string.IsNullOrEmpty(face))
                {
                    into[script] = face;
                }
            }
        }

        private static XElement ChildCi(XElement parent, XNamespace ns, string local)
        {
            if (parent == null)
            {
                return null;
            }

            XElement exact = parent.Element(ns + local);
            if (exact != null)
            {
                return exact;
            }

            foreach (XElement child in parent.Elements())
            {
                if (child.Name.Namespace == ns
                    && string.Equals(child.Name.LocalName, local, StringComparison.OrdinalIgnoreCase))
                {
                    return child;
                }
            }

            return null;
        }

        private static void Walk(
            XElement element,
            Dictionary<string, StyleRec> styles,
            ThemeFonts theme,
            Stack<bool> fields,
            List<Dictionary<string, object>> spans,
            StringBuilder joined)
        {
            if (element == null)
            {
                return;
            }

            foreach (XElement child in element.Elements())
            {
                if (child.Name == Mc + "AlternateContent")
                {
                    Walk(child.Element(Mc + "Choice") ?? child.Element(Mc + "Fallback"), styles, theme, fields, spans, joined);
                }
                else if (child.Name == W + "p")
                {
                    WalkInlines(child, styles, theme, fields, spans, joined);
                }
                else if (child.Name == W + "r" || child.Name == W + "hyperlink" || child.Name == W + "ins"
                    || child.Name == W + "fldSimple" || child.Name == W + "smartTag" || child.Name == W + "customXml")
                {
                    WalkInlines(element, styles, theme, fields, spans, joined);
                    return;
                }
                else if (child.Name == W + "sdt")
                {
                    Walk(child.Element(W + "sdtContent"), styles, theme, fields, spans, joined);
                }
                else if (child.Name == W + "del" || child.Name == W + "moveFrom")
                {
                }
                else
                {
                    Walk(child, styles, theme, fields, spans, joined);
                }
            }
        }

        private static void WalkInlines(
            XElement container,
            Dictionary<string, StyleRec> styles,
            ThemeFonts theme,
            Stack<bool> fields,
            List<Dictionary<string, object>> spans,
            StringBuilder joined)
        {
            foreach (XElement child in container.Elements())
            {
                if (child.Name == W + "r")
                {
                    HandleRun(child, styles, theme, fields, spans, joined);
                }
                else if (child.Name == W + "hyperlink" || child.Name == W + "ins"
                    || child.Name == W + "fldSimple" || child.Name == W + "smartTag" || child.Name == W + "customXml")
                {
                    WalkInlines(child, styles, theme, fields, spans, joined);
                }
                else if (child.Name == W + "sdt")
                {
                    WalkInlines(child.Element(W + "sdtContent") ?? child, styles, theme, fields, spans, joined);
                }
                else if (child.Name == Mc + "AlternateContent")
                {
                    XElement branch = child.Element(Mc + "Choice") ?? child.Element(Mc + "Fallback");
                    WalkInlines(branch ?? child, styles, theme, fields, spans, joined);
                }
            }
        }

        private static void HandleRun(
            XElement run,
            Dictionary<string, StyleRec> styles,
            ThemeFonts theme,
            Stack<bool> fields,
            List<Dictionary<string, object>> spans,
            StringBuilder joined)
        {
            if (run.Ancestors(W + "del").Any() || run.Ancestors(W + "moveFrom").Any())
            {
                return;
            }

            foreach (XElement fld in run.Elements(W + "fldChar"))
            {
                string kind = (string)fld.Attribute(W + "fldCharType") ?? "";
                if (kind == "begin")
                {
                    fields.Push(true);
                }
                else if (kind == "separate" && fields.Count > 0)
                {
                    fields.Pop();
                    fields.Push(false);
                }
                else if (kind == "end" && fields.Count > 0)
                {
                    fields.Pop();
                }
            }

            if (fields.Count > 0 && fields.Peek())
            {
                return;
            }

            string text = RunText(run);
            if (text.Length == 0)
            {
                return;
            }

            Dictionary<string, object> snap = BuildSnapshot(run, styles, theme);
            Dictionary<string, object> direct = BuildDirect(run, theme);
            Dictionary<string, object> style = BuildStyle(run, styles, theme);
            CharFormatCatalog.OmitReadDefaults(snap);
            int runStart = joined.Length;
            joined.Append(text);
            if (spans.Count > 0)
            {
                Dictionary<string, object> previous = spans[spans.Count - 1];
                if (CharFormatCatalog.SameSnapshot((Dictionary<string, object>)previous["char_format"], snap))
                {
                    previous["text"] = (string)previous["text"] + text;
                    if (_trackRunOffsets)
                    {
                        previous["end"] = joined.Length;
                    }

                    return;
                }
            }

            var span = new Dictionary<string, object>
            {
                ["text"] = text,
                ["char_format"] = snap,
                ["format_layers"] = new Dictionary<string, object>
                {
                    ["direct"] = direct,
                    ["style"] = style,
                    ["final"] = new Dictionary<string, object>(snap)
                }
            };
            if (_trackRunOffsets)
            {
                span["start"] = runStart;
                span["end"] = joined.Length;
            }
            if (snap.TryGetValue("highlight_color_index", out object highlight)
                && string.Equals(highlight?.ToString(), "wdYellow", StringComparison.OrdinalIgnoreCase))
            {
                span["highlight_note"] = InsertHighlightNote;
            }

            spans.Add(span);
        }

        private static string RunText(XElement run)
        {
            var sb = new StringBuilder();
            foreach (XElement child in run.Elements())
            {
                if (child.Name == W + "t" || child.Name == W + "delText")
                {
                    sb.Append((string)child);
                }
                else if (child.Name == W + "tab" || child.Name == W + "ptab")
                {
                    sb.Append('\t');
                }
                else if (child.Name == W + "br")
                {
                    string type = (string)child.Attribute(W + "type") ?? "";
                    sb.Append(type == "page" || type == "column" ? '\f' : '\v');
                }
                else if (child.Name == W + "cr")
                {
                    sb.Append('\r');
                }
                else if (child.Name == W + "noBreakHyphen")
                {
                    sb.Append('\u2011');
                }
                else if (child.Name == W + "softHyphen")
                {
                    sb.Append('\u00AD');
                }
                else if (child.Name == W + "sym")
                {
                    sb.Append(SymbolChar(child));
                }
                else if (child.Name == W + "drawing" || child.Name == W + "object" || child.Name == W + "pict")
                {
                    sb.Append('\u0001');
                }
                else if (child.Name == Mc + "AlternateContent")
                {
                    XElement branch = child.Element(Mc + "Choice") ?? child.Element(Mc + "Fallback");
                    if (branch != null && branch.Descendants().Any(d =>
                        d.Name == W + "drawing" || d.Name == W + "object" || d.Name == W + "pict"))
                    {
                        sb.Append('\u0001');
                    }
                }
            }

            return sb.ToString();
        }

        private static char SymbolChar(XElement sym)
        {
            string raw = (string)sym.Attribute(W + "char");
            if (!string.IsNullOrEmpty(raw) && int.TryParse(raw, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out int code)
                && code > 0 && code <= 0xFFFF)
            {
                return (char)code;
            }

            return '\u0001';
        }

        private static Dictionary<string, object> BuildSnapshot(
            XElement run,
            Dictionary<string, StyleRec> styles,
            ThemeFonts theme)
        {
            var format = new Dictionary<string, object>
            {
                ["font_name"] = "",
                ["font_size"] = 0f,
                ["font_color"] = "automatic",
                ["background_color"] = "automatic",
                ["is_bold"] = false,
                ["is_italic"] = false,
                ["has_underline"] = false,
                ["has_strikethrough"] = false,
                ["is_superscript"] = false,
                ["is_subscript"] = false
            };

            if (styles.TryGetValue("__docDefaults", out StyleRec defaults))
            {
                ApplyRpr(format, defaults.Rpr, theme);
            }

            XElement paragraph = run.Ancestors(W + "p").FirstOrDefault();
            string pStyle = (string)paragraph?.Element(W + "pPr")?.Element(W + "pStyle")?.Attribute(W + "val");
            foreach (XElement layer in Chain(styles, pStyle))
            {
                ApplyRpr(format, layer, theme);
            }

            XElement direct = run.Element(W + "rPr");
            string rStyle = (string)direct?.Element(W + "rStyle")?.Attribute(W + "val");
            foreach (XElement layer in Chain(styles, rStyle))
            {
                ApplyRpr(format, layer, theme);
            }

            ApplyRpr(format, direct, theme);
            FinishFontName(format, theme);
            return format;
        }

        private static void RefreshLayerFinal(List<Dictionary<string, object>> runs)
        {
            if (runs == null)
            {
                return;
            }

            foreach (Dictionary<string, object> run in runs)
            {
                if (run == null
                    || !(run.TryGetValue("char_format", out object format) && format is Dictionary<string, object> final)
                    || !(run.TryGetValue("format_layers", out object layers) && layers is Dictionary<string, object> layerMap))
                {
                    continue;
                }

                layerMap["final"] = new Dictionary<string, object>(final);
            }
        }

        private static Dictionary<string, object> BuildDirect(XElement run, ThemeFonts theme)
        {
            var format = new Dictionary<string, object>();
            ApplyRpr(format, run?.Element(W + "rPr"), theme);
            FinishFontName(format, theme);
            return format;
        }

        private static Dictionary<string, object> BuildStyle(
            XElement run,
            Dictionary<string, StyleRec> styles,
            ThemeFonts theme)
        {
            var format = new Dictionary<string, object>();
            XElement paragraph = run.Ancestors(W + "p").FirstOrDefault();
            string pStyle = (string)paragraph?.Element(W + "pPr")?.Element(W + "pStyle")?.Attribute(W + "val");
            foreach (XElement layer in Chain(styles, pStyle))
            {
                ApplyRpr(format, layer, theme);
            }

            XElement direct = run.Element(W + "rPr");
            string rStyle = (string)direct?.Element(W + "rStyle")?.Attribute(W + "val");
            foreach (XElement layer in Chain(styles, rStyle))
            {
                ApplyRpr(format, layer, theme);
            }

            FinishFontName(format, theme);
            return format;
        }

        private static List<XElement> Chain(Dictionary<string, StyleRec> styles, string styleId)
        {
            var leafToRoot = new List<XElement>();
            var guard = new HashSet<string>(StringComparer.Ordinal);
            string id = styleId;
            while (!string.IsNullOrEmpty(id) && guard.Add(id) && styles.TryGetValue(id, out StyleRec rec))
            {
                if (rec.Rpr != null)
                {
                    leafToRoot.Add(rec.Rpr);
                }

                id = rec.BasedOn;
            }

            leafToRoot.Reverse();
            return leafToRoot;
        }

        private static void FinishFontName(Dictionary<string, object> format, ThemeFonts theme)
        {
            if (string.IsNullOrWhiteSpace(Text(format, "name_far_east")))
            {
                string face = ScriptFace(theme, Text(format, "__east_token"), EastScript(format));
                if (!string.IsNullOrEmpty(face))
                {
                    format["name_far_east"] = face;
                }
            }

            format.Remove("__east_token");
            format.Remove("__hint");

            string east = Text(format, "name_far_east");
            string ascii = Text(format, "name_ascii");
            string other = Text(format, "name_other");
            string bi = Text(format, "name_bi");
            string name = FirstName(east, ascii, other, bi);
            if (!string.IsNullOrEmpty(name))
            {
                format["font_name"] = name;
            }

            if (!format.ContainsKey("name") || string.IsNullOrWhiteSpace(Text(format, "name")))
            {
                string slot = FirstName(ascii, other, east, bi);
                if (!string.IsNullOrEmpty(slot))
                {
                    format["name"] = slot;
                }
            }
        }

        private static string FirstName(params string[] slots)
        {
            foreach (string slot in slots)
            {
                if (!string.IsNullOrWhiteSpace(slot))
                {
                    return slot;
                }
            }

            return "";
        }

        private static string EastScript(Dictionary<string, object> format)
        {
            string lang = Text(format, "language_id_far_east");
            if (lang.IndexOf("Japanese", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return "Jpan";
            }

            if (lang.IndexOf("Korean", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return "Hang";
            }

            if (lang.IndexOf("Traditional", StringComparison.OrdinalIgnoreCase) >= 0
                || lang.IndexOf("HongKong", StringComparison.OrdinalIgnoreCase) >= 0
                || lang.IndexOf("Macao", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return "Hant";
            }

            if (lang.IndexOf("Chinese", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return "Hans";
            }

            if (string.Equals(Text(format, "__hint"), "eastAsia", StringComparison.OrdinalIgnoreCase))
            {
                return "Hans";
            }

            return "";
        }

        private static string ScriptFace(ThemeFonts theme, string token, string script)
        {
            if (theme == null || string.IsNullOrEmpty(token) || string.IsNullOrEmpty(script))
            {
                return "";
            }

            bool major = token.StartsWith("major", StringComparison.OrdinalIgnoreCase);
            Dictionary<string, string> map = major ? theme.MajorScript : theme.MinorScript;
            return map != null && map.TryGetValue(script, out string face) ? face : "";
        }

        private static string Text(Dictionary<string, object> format, string key)
        {
            return format.TryGetValue(key, out object value) ? value?.ToString() ?? "" : "";
        }

        private static void ApplyRpr(Dictionary<string, object> format, XElement rPr, ThemeFonts theme)
        {
            if (rPr == null)
            {
                return;
            }

            XElement fonts = rPr.Element(W + "rFonts");
            if (fonts != null)
            {
                string hint = (string)fonts.Attribute(W + "hint");
                if (!string.IsNullOrEmpty(hint))
                {
                    format["__hint"] = hint;
                }

                PutSlot(format, "name_ascii", fonts, "ascii", "asciiTheme", theme, latin: true);
                PutSlot(format, "name_far_east", fonts, "eastAsia", "eastAsiaTheme", theme, latin: false, east: true);
                PutSlot(format, "name_bi", fonts, "cs", "cstheme", theme, latin: false, bidi: true);
                PutSlot(format, "name_other", fonts, "hAnsi", "hAnsiTheme", theme, latin: true);
            }

            PutHalfPoints(format, "font_size", rPr.Element(W + "sz"));
            PutHalfPoints(format, "size_bi", rPr.Element(W + "szCs"));
            ApplyColor(format, rPr.Element(W + "color"), theme);
            ApplyUnderline(format, rPr.Element(W + "u"), theme);
            PutBool(format, "is_bold", rPr.Element(W + "b"));
            PutBool(format, "bold_bi", rPr.Element(W + "bCs"));
            PutBool(format, "is_italic", rPr.Element(W + "i"));
            PutBool(format, "italic_bi", rPr.Element(W + "iCs"));
            PutBool(format, "has_strikethrough", rPr.Element(W + "strike"));
            PutBool(format, "double_strikethrough", rPr.Element(W + "dstrike"));
            ApplyVertAlign(format, rPr.Element(W + "vertAlign"));
            PutBool(format, "shadow", rPr.Element(W + "shadow"));
            PutBool(format, "outline", rPr.Element(W + "outline"));
            PutBool(format, "emboss", rPr.Element(W + "emboss"));
            PutBool(format, "engrave", rPr.Element(W + "imprint"));
            PutBool(format, "small_caps", rPr.Element(W + "smallCaps"));
            PutBool(format, "all_caps", rPr.Element(W + "caps"));
            PutEnum(format, "emphasis_mark", rPr.Element(W + "em"), EmphasisMap);
            PutPercent(format, "scaling", rPr.Element(W + "w"));
            PutTwips(format, "spacing", rPr.Element(W + "spacing"));
            PutHalfPoints(format, "position", rPr.Element(W + "position"));
            PutHalfPoints(format, "kerning", rPr.Element(W + "kern"));
            XElement snap = rPr.Element(W + "snapToGrid");
            if (snap != null)
            {
                format["disable_character_space_grid"] = !On(snap);
            }

            ApplyShading(format, rPr.Element(W + "shd"), theme);
            PutHighlight(format, rPr.Element(W + "highlight"));
            ApplyBorders(format, rPr.Element(W + "bdr"), theme);
            ApplyLang(format, rPr.Element(W + "lang"));
            if (rPr.Element(W + "noProof") != null)
            {
                format["no_proofing"] = On(rPr.Element(W + "noProof"));
            }

            PutTwips(format, "fit_text_width", rPr.Element(W + "fitText"));
            PutEnum(format, "ligatures", ChildCi(rPr, W14, "ligatures"), LigatureMap);
            PutEnum(format, "number_form", ChildCi(rPr, W14, "numForm"), NumberFormMap);
            PutEnum(format, "number_spacing", ChildCi(rPr, W14, "numSpacing"), NumberSpacingMap);
            ApplyStylistic(format, ChildCi(rPr, W14, "stylisticSets"));
            PutBool(format, "contextual_alternates", ChildCi(rPr, W14, "cntxtAlts"));
            ApplyTextFill(format, ChildCi(rPr, W14, "textFill"), theme);
            ApplyGlow(format, ChildCi(rPr, W14, "glow"), theme);
            ApplyLine(format, ChildCi(rPr, W14, "textOutline"), theme);
            ApplyReflection(format, ChildCi(rPr, W14, "reflection"));
            ApplyTextShadow(format, ChildCi(rPr, W14, "shadow"), theme);
        }

        private static void PutSlot(
            Dictionary<string, object> format,
            string key,
            XElement fonts,
            string concrete,
            string themeAttr,
            ThemeFonts theme,
            bool latin,
            bool east = false,
            bool bidi = false)
        {
            string token = (string)fonts.Attribute(W + themeAttr);
            string name = (string)fonts.Attribute(W + concrete);
            if (string.IsNullOrWhiteSpace(name))
            {
                name = ThemeFace(theme, token, latin, east, bidi);
            }

            if (!string.IsNullOrWhiteSpace(name))
            {
                format[key] = name;
                if (east)
                {
                    format.Remove("__east_token");
                }

                return;
            }

            if (east && !string.IsNullOrEmpty(token))
            {
                format["__east_token"] = token;
            }
        }

        private static string ThemeFace(ThemeFonts theme, string token, bool latin, bool east, bool bidi)
        {
            if (string.IsNullOrEmpty(token) || theme == null)
            {
                return "";
            }

            bool major = token.StartsWith("major", StringComparison.OrdinalIgnoreCase);
            if (east)
            {
                return major ? theme.MajorEastAsia : theme.MinorEastAsia;
            }

            if (bidi)
            {
                return major ? theme.MajorBidi : theme.MinorBidi;
            }

            if (latin)
            {
                return major ? theme.MajorLatin : theme.MinorLatin;
            }

            return "";
        }

        private static void PutHalfPoints(Dictionary<string, object> format, string key, XElement element)
        {
            if (element == null)
            {
                return;
            }

            if (double.TryParse((string)element.Attribute(W + "val"), NumberStyles.Float, CultureInfo.InvariantCulture, out double half))
            {
                format[key] = (float)(half / 2.0);
            }
        }

        private static void PutTwips(Dictionary<string, object> format, string key, XElement element)
        {
            if (element == null)
            {
                return;
            }

            if (double.TryParse((string)element.Attribute(W + "val"), NumberStyles.Float, CultureInfo.InvariantCulture, out double twips))
            {
                format[key] = (float)(twips / 20.0);
            }
        }

        private static void PutPercent(Dictionary<string, object> format, string key, XElement element)
        {
            if (element == null)
            {
                return;
            }

            if (double.TryParse((string)element.Attribute(W + "val"), NumberStyles.Float, CultureInfo.InvariantCulture, out double number))
            {
                format[key] = (float)number;
            }
        }

        private static void PutBool(Dictionary<string, object> format, string key, XElement element)
        {
            if (element != null)
            {
                format[key] = On(element);
            }
        }

        private static bool On(XElement element)
        {
            if (element == null)
            {
                return false;
            }

            string val = (string)element.Attribute(W + "val") ?? (string)element.Attribute(W14 + "val");
            if (string.IsNullOrEmpty(val))
            {
                return true;
            }

            return !string.Equals(val, "0", StringComparison.OrdinalIgnoreCase)
                && !string.Equals(val, "false", StringComparison.OrdinalIgnoreCase)
                && !string.Equals(val, "off", StringComparison.OrdinalIgnoreCase)
                && !string.Equals(val, "none", StringComparison.OrdinalIgnoreCase);
        }

        private static void ApplyVertAlign(Dictionary<string, object> format, XElement element)
        {
            if (element == null)
            {
                return;
            }

            string val = (string)element.Attribute(W + "val") ?? "";
            format["is_superscript"] = val == "superscript";
            format["is_subscript"] = val == "subscript";
        }

        private static void ApplyColor(Dictionary<string, object> format, XElement color, ThemeFonts theme)
        {
            if (color == null)
            {
                return;
            }

            string rgb = ResolveRgb(color, theme, W);
            format["font_color"] = string.IsNullOrEmpty(rgb) ? "automatic" : rgb;
            Dictionary<string, object> obj = ColorObject(color, theme, W);
            if (obj != null)
            {
                format["text_color"] = obj;
            }
        }

        private static void ApplyUnderline(Dictionary<string, object> format, XElement element, ThemeFonts theme)
        {
            if (element == null)
            {
                return;
            }

            string val = (string)element.Attribute(W + "val") ?? "single";
            format["has_underline"] = !string.Equals(val, "none", StringComparison.OrdinalIgnoreCase);
            if (UnderlineMap.TryGetValue(val, out string name))
            {
                format["underline"] = name;
            }

            string rgb = ResolveRgb(element, theme, W, "color");
            if (!string.IsNullOrEmpty(rgb))
            {
                format["underline_color"] = rgb;
            }
        }

        private static void ApplyShading(Dictionary<string, object> format, XElement shd, ThemeFonts theme)
        {
            if (shd == null)
            {
                return;
            }

            string fill = ResolveRgb(shd, theme, W, "fill");
            if (!string.IsNullOrEmpty(fill))
            {
                format["background_color"] = fill;
            }

            var map = new Dictionary<string, object>();
            string texture = (string)shd.Attribute(W + "val") ?? "clear";
            if (TextureMap.TryGetValue(texture, out string textureName))
            {
                map["texture"] = textureName;
            }

            string back = ResolveRgb(shd, theme, W, "fill");
            string fore = ResolveRgb(shd, theme, W, "color");
            map["background_pattern_color"] = string.IsNullOrEmpty(back) ? "automatic" : back;
            map["foreground_pattern_color"] = string.IsNullOrEmpty(fore) ? "automatic" : fore;
            format["shading"] = map;
        }

        private static void PutHighlight(Dictionary<string, object> format, XElement element)
        {
            if (element == null)
            {
                return;
            }

            string val = (string)element.Attribute(W + "val") ?? "none";
            format["highlight_color_index"] = HighlightMap.TryGetValue(val, out string name) ? name : "wdNoHighlight";
        }

        private static void ApplyBorders(Dictionary<string, object> format, XElement borders, ThemeFonts theme)
        {
            if (borders == null)
            {
                return;
            }

            XElement side = borders.Element(W + "top")
                ?? borders.Element(W + "left")
                ?? borders.Element(W + "bottom")
                ?? borders.Element(W + "right");
            if (side == null)
            {
                side = borders.Attribute(W + "val") != null ? borders : null;
            }

            if (side == null)
            {
                return;
            }

            var map = new Dictionary<string, object>();
            string val = (string)side.Attribute(W + "val") ?? "none";
            if (LineStyleMap.TryGetValue(val, out string style))
            {
                map["outside_line_style"] = style;
            }

            if (int.TryParse((string)side.Attribute(W + "sz"), out int eighths) && WidthMap.TryGetValue(eighths, out string width))
            {
                map["outside_line_width"] = width;
            }

            string rgb = ResolveRgb(side, theme, W, "color");
            map["outside_color"] = string.IsNullOrEmpty(rgb) ? "automatic" : rgb;
            format["borders"] = map;
        }

        private static void ApplyLang(Dictionary<string, object> format, XElement lang)
        {
            if (lang == null)
            {
                return;
            }

            PutLang(format, "language_id", (string)lang.Attribute(W + "val"));
            PutLang(format, "language_id_far_east", (string)lang.Attribute(W + "eastAsia"));
        }

        private static void PutLang(Dictionary<string, object> format, string key, string bcp)
        {
            if (string.IsNullOrEmpty(bcp))
            {
                return;
            }

            if (LangMap.TryGetValue(bcp, out string name))
            {
                format[key] = name;
            }
        }

        private static void PutEnum(Dictionary<string, object> format, string key, XElement element, Dictionary<string, string> map)
        {
            if (element == null)
            {
                return;
            }

            string val = (string)element.Attribute(W + "val") ?? (string)element.Attribute(W14 + "val") ?? "";
            if (map.TryGetValue(val, out string name))
            {
                format[key] = name;
            }
        }

        private static void ApplyStylistic(Dictionary<string, object> format, XElement sets)
        {
            if (sets == null)
            {
                return;
            }

            string id = (string)ChildCi(sets, W14, "styleSet")?.Attribute(W14 + "id");
            if (string.IsNullOrEmpty(id))
            {
                return;
            }

            if (id.Length == 1)
            {
                id = "0" + id;
            }

            format["stylistic_set"] = "wdStylisticSet" + id;
        }

        private static void ApplyTextFill(Dictionary<string, object> format, XElement fill, ThemeFonts theme)
        {
            if (fill == null)
            {
                return;
            }

            XElement solid = fill.Element(W14 + "solidFill") ?? fill.Descendants(W14 + "srgbClr").FirstOrDefault()?.Parent;
            XElement color = solid?.Element(W14 + "srgbClr") ?? solid?.Element(W14 + "schemeClr");
            if (color == null)
            {
                return;
            }

            var map = new Dictionary<string, object> { ["visible"] = true };
            double transparency = AlphaTransparency(color);
            if (transparency > 0)
            {
                map["transparency"] = transparency;
            }

            Dictionary<string, object> fore = DrawingColor(color, theme);
            if (fore != null)
            {
                map["fore_color"] = fore;
                if (fore.TryGetValue("rgb", out object rgb) && rgb != null)
                {
                    format["font_color"] = rgb;
                    format["text_color"] = fore;
                }
            }

            format["fill"] = map;
        }

        private static void ApplyGlow(Dictionary<string, object> format, XElement glow, ThemeFonts theme)
        {
            if (glow == null)
            {
                return;
            }

            var map = new Dictionary<string, object>();
            if (double.TryParse((string)glow.Attribute(W14 + "rad"), NumberStyles.Float, CultureInfo.InvariantCulture, out double emu))
            {
                map["radius"] = (float)(emu / 12700.0);
            }

            XElement color = glow.Element(W14 + "srgbClr") ?? glow.Element(W14 + "schemeClr");
            if (color != null)
            {
                map["transparency"] = AlphaTransparency(color);
                Dictionary<string, object> obj = DrawingColor(color, null);
                if (obj != null)
                {
                    map["color"] = obj;
                }
            }

            if (map.Count > 0)
            {
                format["glow"] = map;
            }
        }

        private static void ApplyLine(Dictionary<string, object> format, XElement line, ThemeFonts theme)
        {
            if (line == null)
            {
                return;
            }

            var map = new Dictionary<string, object> { ["visible"] = true };
            if (double.TryParse((string)line.Attribute(W14 + "w"), NumberStyles.Float, CultureInfo.InvariantCulture, out double emu))
            {
                map["weight"] = (float)(emu / 12700.0);
            }

            string dash = (string)line.Element(W14 + "prstDash")?.Attribute(W14 + "val");
            if (!string.IsNullOrEmpty(dash))
            {
                map["dash_style"] = "msoLineDashStyle" + char.ToUpperInvariant(dash[0]) + dash.Substring(1);
            }

            XElement fill = line.Element(W14 + "solidFill");
            XElement color = fill?.Element(W14 + "srgbClr") ?? fill?.Element(W14 + "schemeClr");
            if (color != null)
            {
                map["transparency"] = AlphaTransparency(color);
                Dictionary<string, object> fore = DrawingColor(color, null);
                if (fore != null)
                {
                    map["fore_color"] = fore;
                }
            }

            format["line"] = map;
        }

        private static void ApplyReflection(Dictionary<string, object> format, XElement reflection)
        {
            if (reflection == null)
            {
                return;
            }

            var map = new Dictionary<string, object>();
            string prst = (string)reflection.Attribute(W14 + "prst") ?? "";
            if (ReflectionMap.TryGetValue(prst, out string type))
            {
                map["type"] = type;
            }
            else if (reflection.HasAttributes)
            {
                map["type"] = "msoReflectionType1";
            }

            if (double.TryParse((string)reflection.Attribute(W14 + "stA"), NumberStyles.Float, CultureInfo.InvariantCulture, out double startAlpha))
            {
                map["transparency"] = Math.Round(1.0 - startAlpha / 100000.0, 4);
            }

            if (double.TryParse((string)reflection.Attribute(W14 + "sy"), NumberStyles.Float, CultureInfo.InvariantCulture, out double sy))
            {
                map["size"] = (float)(Math.Abs(sy) / 1000.0);
            }

            if (map.Count > 0)
            {
                format["reflection"] = map;
            }
        }

        private static void ApplyTextShadow(Dictionary<string, object> format, XElement shadow, ThemeFonts theme)
        {
            if (shadow == null)
            {
                return;
            }

            var map = new Dictionary<string, object> { ["visible"] = true };
            string prst = (string)shadow.Attribute(W14 + "prst") ?? "";
            if (!string.IsNullOrEmpty(prst))
            {
                map["type"] = "msoShadow" + char.ToUpperInvariant(prst[0]) + prst.Substring(1);
            }

            if (double.TryParse((string)shadow.Attribute(W14 + "dir"), NumberStyles.Float, CultureInfo.InvariantCulture, out double dir)
                && double.TryParse((string)shadow.Attribute(W14 + "dist"), NumberStyles.Float, CultureInfo.InvariantCulture, out double dist))
            {
                double rad = dir / 60000.0 * Math.PI / 180.0;
                double points = dist / 12700.0;
                map["offset_x"] = (float)(points * Math.Cos(rad));
                map["offset_y"] = (float)(points * Math.Sin(rad));
            }

            XElement color = shadow.Element(W14 + "srgbClr") ?? shadow.Element(W14 + "schemeClr");
            if (color != null)
            {
                map["transparency"] = AlphaTransparency(color);
                Dictionary<string, object> fore = DrawingColor(color, null);
                if (fore != null)
                {
                    map["fore_color"] = fore;
                }
            }

            format["text_shadow"] = map;
        }

        private static double AlphaTransparency(XElement color)
        {
            string raw = (string)color?.Element(W14 + "alpha")?.Attribute(W14 + "val");
            if (!double.TryParse(raw, NumberStyles.Float, CultureInfo.InvariantCulture, out double alpha))
            {
                return 0;
            }

            return Math.Round(alpha / 100000.0, 4);
        }

        private static Dictionary<string, object> DrawingColor(XElement color, ThemeFonts theme)
        {
            if (color == null)
            {
                return null;
            }

            if (color.Name == W14 + "srgbClr" || color.Name == A + "srgbClr")
            {
                string rgb = NormalizeRgb((string)color.Attribute(W14 + "val") ?? (string)color.Attribute("val"));
                return string.IsNullOrEmpty(rgb)
                    ? null
                    : new Dictionary<string, object> { ["rgb"] = rgb, ["type"] = "wdColorTypeRGB" };
            }

            string scheme = (string)color.Attribute(W14 + "val") ?? "";
            string themeName = ThemeColorName(scheme);
            var map = new Dictionary<string, object> { ["type"] = "wdColorTypeScheme" };
            if (!string.IsNullOrEmpty(themeName))
            {
                map["object_theme_color"] = themeName;
            }

            if (theme != null && theme.Scheme.TryGetValue(SchemeKey(scheme), out string schemeRgb))
            {
                map["rgb"] = schemeRgb;
            }

            return map;
        }

        private static Dictionary<string, object> ColorObject(XElement color, ThemeFonts theme, XNamespace ns)
        {
            string themeAttr = (string)color.Attribute(ns + "themeColor");
            string rgb = ResolveRgb(color, theme, ns);
            if (string.IsNullOrEmpty(themeAttr) && string.IsNullOrEmpty(rgb))
            {
                return null;
            }

            var map = new Dictionary<string, object>();
            if (!string.IsNullOrEmpty(themeAttr))
            {
                map["type"] = "wdColorTypeScheme";
                string name = ThemeColorName(themeAttr);
                if (!string.IsNullOrEmpty(name))
                {
                    map["object_theme_color"] = name;
                }
            }
            else
            {
                map["type"] = "wdColorTypeRGB";
            }

            if (!string.IsNullOrEmpty(rgb))
            {
                map["rgb"] = rgb;
            }

            return map;
        }

        private static string ResolveRgb(XElement element, ThemeFonts theme, XNamespace ns, string attr = "val")
        {
            if (element == null)
            {
                return "";
            }

            string val = (string)element.Attribute(ns + attr) ?? "";
            if (!string.IsNullOrEmpty(val)
                && !string.Equals(val, "auto", StringComparison.OrdinalIgnoreCase)
                && val.Length >= 6
                && IsHex(val))
            {
                return NormalizeRgb(val);
            }

            string themeAttr = (string)element.Attribute(ns + "themeColor");
            if (string.IsNullOrEmpty(themeAttr) || theme == null)
            {
                return "";
            }

            if (!theme.Scheme.TryGetValue(SchemeKey(themeAttr), out string rgb))
            {
                return "";
            }

            return ApplyTintShade(rgb, (string)element.Attribute(ns + "themeTint"), (string)element.Attribute(ns + "themeShade"));
        }

        private static string SchemeKey(string themeColor)
        {
            switch ((themeColor ?? "").ToLowerInvariant())
            {
                case "dark1":
                case "text1": return "dk1";
                case "light1":
                case "background1": return "lt1";
                case "dark2":
                case "text2": return "dk2";
                case "light2":
                case "background2": return "lt2";
                case "accent1": return "accent1";
                case "accent2": return "accent2";
                case "accent3": return "accent3";
                case "accent4": return "accent4";
                case "accent5": return "accent5";
                case "accent6": return "accent6";
                case "hyperlink": return "hlink";
                case "followedhyperlink": return "folHlink";
                default: return themeColor ?? "";
            }
        }

        private static string ThemeColorName(string themeColor)
        {
            switch ((themeColor ?? "").ToLowerInvariant())
            {
                case "dark1":
                case "text1":
                case "dk1": return "wdThemeColorDark1";
                case "light1":
                case "background1":
                case "lt1": return "wdThemeColorLight1";
                case "dark2":
                case "text2":
                case "dk2": return "wdThemeColorDark2";
                case "light2":
                case "background2":
                case "lt2": return "wdThemeColorLight2";
                case "accent1": return "wdThemeColorAccent1";
                case "accent2": return "wdThemeColorAccent2";
                case "accent3": return "wdThemeColorAccent3";
                case "accent4": return "wdThemeColorAccent4";
                case "accent5": return "wdThemeColorAccent5";
                case "accent6": return "wdThemeColorAccent6";
                case "hyperlink":
                case "hlink": return "wdThemeColorHyperlink";
                case "followedhyperlink":
                case "folhlink": return "wdThemeColorFollowedHyperlink";
                default: return "";
            }
        }

        private static string ApplyTintShade(string rgb, string tintHex, string shadeHex)
        {
            if (!TryRgb(rgb, out int r, out int g, out int b))
            {
                return rgb;
            }

            if (!string.IsNullOrEmpty(shadeHex) && int.TryParse(shadeHex, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out int shade))
            {
                r = r * shade / 255;
                g = g * shade / 255;
                b = b * shade / 255;
            }

            if (!string.IsNullOrEmpty(tintHex) && int.TryParse(tintHex, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out int tint))
            {
                r = r + (255 - r) * (255 - tint) / 255;
                g = g + (255 - g) * (255 - tint) / 255;
                b = b + (255 - b) * (255 - tint) / 255;
            }

            return "#" + r.ToString("X2") + g.ToString("X2") + b.ToString("X2");
        }

        private static bool TryRgb(string rgb, out int r, out int g, out int b)
        {
            r = g = b = 0;
            string hex = (rgb ?? "").Trim().TrimStart('#');
            if (hex.Length < 6 || !IsHex(hex))
            {
                return false;
            }

            hex = hex.Substring(hex.Length - 6);
            r = int.Parse(hex.Substring(0, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture);
            g = int.Parse(hex.Substring(2, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture);
            b = int.Parse(hex.Substring(4, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture);
            return true;
        }

        private static bool IsHex(string text)
        {
            foreach (char c in text)
            {
                if (!Uri.IsHexDigit(c))
                {
                    return false;
                }
            }

            return text.Length > 0;
        }

        private static string NormalizeRgb(string raw)
        {
            if (string.IsNullOrWhiteSpace(raw))
            {
                return "";
            }

            string hex = raw.Trim().TrimStart('#');
            if (hex.Length < 6 || !IsHex(hex))
            {
                return "";
            }

            return "#" + hex.Substring(hex.Length - 6).ToUpperInvariant();
        }

        private static string ColorVal(XElement srgb)
        {
            return srgb == null ? null : (string)srgb.Attribute("val");
        }

        private static readonly Dictionary<string, string> UnderlineMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            ["none"] = "wdUnderlineNone",
            ["single"] = "wdUnderlineSingle",
            ["words"] = "wdUnderlineWords",
            ["double"] = "wdUnderlineDouble",
            ["dotted"] = "wdUnderlineDotted",
            ["thick"] = "wdUnderlineThick",
            ["dash"] = "wdUnderlineDash",
            ["dotDash"] = "wdUnderlineDotDash",
            ["dotDotDash"] = "wdUnderlineDotDotDash",
            ["wave"] = "wdUnderlineWavy",
            ["dottedHeavy"] = "wdUnderlineDottedHeavy",
            ["dashHeavy"] = "wdUnderlineDashHeavy",
            ["dashLong"] = "wdUnderlineDashLong",
            ["dashLongHeavy"] = "wdUnderlineDashLongHeavy",
            ["dotDashHeavy"] = "wdUnderlineDotDashHeavy",
            ["dotDotDashHeavy"] = "wdUnderlineDotDotDashHeavy",
            ["wavyHeavy"] = "wdUnderlineWavyHeavy",
            ["wavyDouble"] = "wdUnderlineWavyDouble"
        };

        private static readonly Dictionary<string, string> EmphasisMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            ["none"] = "wdEmphasisMarkNone",
            ["dot"] = "wdEmphasisMarkOverSolidCircle",
            ["comma"] = "wdEmphasisMarkOverComma",
            ["circle"] = "wdEmphasisMarkOverWhiteCircle",
            ["underDot"] = "wdEmphasisMarkUnderSolidCircle"
        };

        private static readonly Dictionary<string, string> HighlightMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            ["none"] = "wdNoHighlight",
            ["black"] = "wdBlack",
            ["blue"] = "wdBlue",
            ["cyan"] = "wdTurquoise",
            ["green"] = "wdBrightGreen",
            ["magenta"] = "wdPink",
            ["red"] = "wdRed",
            ["yellow"] = "wdYellow",
            ["white"] = "wdWhite",
            ["darkBlue"] = "wdDarkBlue",
            ["darkCyan"] = "wdTeal",
            ["darkGreen"] = "wdGreen",
            ["darkMagenta"] = "wdViolet",
            ["darkRed"] = "wdDarkRed",
            ["darkYellow"] = "wdDarkYellow",
            ["darkGray"] = "wdGray50",
            ["lightGray"] = "wdGray25"
        };

        private static readonly Dictionary<string, string> TextureMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            ["clear"] = "wdTextureNone",
            ["nil"] = "wdTextureNone",
            ["solid"] = "wdTextureSolid",
            ["pct5"] = "wdTexture5Percent",
            ["pct10"] = "wdTexture10Percent",
            ["pct12"] = "wdTexture12Pt5Percent",
            ["pct15"] = "wdTexture15Percent",
            ["pct20"] = "wdTexture20Percent",
            ["pct25"] = "wdTexture25Percent",
            ["pct30"] = "wdTexture30Percent",
            ["pct35"] = "wdTexture35Percent",
            ["pct37"] = "wdTexture37Pt5Percent",
            ["pct40"] = "wdTexture40Percent",
            ["pct45"] = "wdTexture45Percent",
            ["pct50"] = "wdTexture50Percent",
            ["pct55"] = "wdTexture55Percent",
            ["pct60"] = "wdTexture60Percent",
            ["pct62"] = "wdTexture62Pt5Percent",
            ["pct65"] = "wdTexture65Percent",
            ["pct70"] = "wdTexture70Percent",
            ["pct75"] = "wdTexture75Percent",
            ["pct80"] = "wdTexture80Percent",
            ["pct85"] = "wdTexture85Percent",
            ["pct87"] = "wdTexture87Pt5Percent",
            ["pct90"] = "wdTexture90Percent",
            ["pct95"] = "wdTexture95Percent"
        };

        private static readonly Dictionary<string, string> LineStyleMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            ["none"] = "wdLineStyleNone",
            ["nil"] = "wdLineStyleNone",
            ["single"] = "wdLineStyleSingle",
            ["double"] = "wdLineStyleDouble",
            ["dot"] = "wdLineStyleDot",
            ["dotted"] = "wdLineStyleDot",
            ["dash"] = "wdLineStyleDashSmallGap",
            ["dashSmallGap"] = "wdLineStyleDashSmallGap",
            ["dotDash"] = "wdLineStyleDotDash",
            ["dotDotDash"] = "wdLineStyleDotDotDash",
            ["triple"] = "wdLineStyleTriple",
            ["wave"] = "wdLineStyleWave"
        };

        private static readonly Dictionary<int, string> WidthMap = new Dictionary<int, string>
        {
            [2] = "wdLineWidth025pt",
            [4] = "wdLineWidth050pt",
            [6] = "wdLineWidth075pt",
            [8] = "wdLineWidth100pt",
            [12] = "wdLineWidth150pt",
            [18] = "wdLineWidth225pt",
            [24] = "wdLineWidth300pt",
            [36] = "wdLineWidth450pt",
            [48] = "wdLineWidth600pt"
        };

        private static readonly Dictionary<string, string> LangMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            ["en-US"] = "wdEnglishUS",
            ["en-GB"] = "wdEnglishUK",
            ["zh-CN"] = "wdSimplifiedChinese",
            ["zh-TW"] = "wdTraditionalChinese",
            ["ja-JP"] = "wdJapanese",
            ["ko-KR"] = "wdKorean",
            ["fr-FR"] = "wdFrench",
            ["de-DE"] = "wdGerman",
            ["es-ES"] = "wdSpanish",
            ["ru-RU"] = "wdRussian",
            ["x-none"] = "wdLanguageNone"
        };

        private static readonly Dictionary<string, string> LigatureMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            ["none"] = "wdLigaturesNone",
            ["standard"] = "wdLigaturesStandard",
            ["contextual"] = "wdLigaturesContextual",
            ["historical"] = "wdLigaturesHistorical",
            ["discretional"] = "wdLigaturesDiscretional",
            ["standardContextual"] = "wdLigaturesStandardContextual",
            ["standardHistorical"] = "wdLigaturesStandardHistorical",
            ["contextualHistorical"] = "wdLigaturesContextualHistorical",
            ["standardDiscretional"] = "wdLigaturesStandardDiscretional",
            ["contextualDiscretional"] = "wdLigaturesContextualDiscretional",
            ["historicalDiscretional"] = "wdLigaturesHistoricalDiscretional",
            ["standardContextualHistorical"] = "wdLigaturesStandardContextualHistorical",
            ["standardContextualDiscretional"] = "wdLigaturesStandardContextualDiscretional",
            ["standardHistoricalDiscretional"] = "wdLigaturesStandardHistoricalDiscretional",
            ["contextualHistoricalDiscretional"] = "wdLigaturesContextualHistoricalDiscretional",
            ["all"] = "wdLigaturesAll"
        };

        private static readonly Dictionary<string, string> NumberFormMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            ["default"] = "wdNumberFormDefault",
            ["lining"] = "wdNumberFormLining",
            ["oldStyle"] = "wdNumberFormOldStyle"
        };

        private static readonly Dictionary<string, string> NumberSpacingMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            ["default"] = "wdNumberSpacingDefault",
            ["proportional"] = "wdNumberSpacingProportional",
            ["tabular"] = "wdNumberSpacingTabular"
        };

        private static readonly Dictionary<string, string> ReflectionMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            ["tightReflectionTouching"] = "msoReflectionType1",
            ["halfReflectionTouching"] = "msoReflectionType2",
            ["fullReflectionTouching"] = "msoReflectionType3",
            ["tightReflection"] = "msoReflectionType4",
            ["halfReflection"] = "msoReflectionType5",
            ["fullReflection"] = "msoReflectionType6",
            ["none"] = "msoReflectionTypeNone"
        };
    }
}
