using System;
using System.Collections.Generic;
using System.Globalization;
using System.Xml.Linq;
using Newtonsoft.Json.Linq;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    /// <summary>改样式定义，或新建样式。写了文字格式才同时写段落样式和链接字符样式。</summary>
    public static class StructureStyleEditor
    {
        public static string ValidateCreate(bool create, string name, ICollection<string> existing)
        {
            if (!create)
            {
                return null;
            }

            if (string.IsNullOrWhiteSpace(name))
            {
                return "新建必须带 name。";
            }

            if (existing != null)
            {
                foreach (string item in existing)
                {
                    if (string.Equals(item, name.Trim(), StringComparison.OrdinalIgnoreCase))
                    {
                        return "已有同名样式，不覆盖。";
                    }
                }
            }

            return null;
        }

        public static Dictionary<string, object> Customize(
            Word.Document doc,
            Dictionary<string, object> args,
            IList<StructureStyleRecord> styles)
        {
            bool create = IsTrue(args, "create");
            string name = Arg(args, "name");
            string structure = Arg(args, "structure");
            Dictionary<string, object> text = AsDict(args, "text_format");
            Dictionary<string, object> paragraph = AsDict(args, "paragraph_format");
            var warnings = new List<string>();

            if (create)
            {
                if (!string.IsNullOrEmpty(structure))
                {
                    return Fail("新建不要带 structure，带 name。");
                }

                string known = ValidateCreate(true, name, StyleNames(doc));
                if (known != null)
                {
                    return Fail(known);
                }

                return Create(doc, name.Trim(), text, paragraph);
            }

            if (string.IsNullOrWhiteSpace(structure))
            {
                return Fail("改已有样式必须带 structure。");
            }

            StructureStyleRecord record = StructureStyleCatalog.Find(styles, structure);
            if (record == null)
            {
                return Fail("没有样式 " + structure + "。");
            }

            if (record.Support == "character" && paragraph != null && paragraph.Count > 0)
            {
                warnings.Add("该样式只支持字符格式，已忽略 paragraph_format。");
                paragraph = null;
            }

            try
            {
                Word.Style characterStyle = record.Support == "character"
                    ? null
                    : Resolve(doc, record, true);
                Word.Style paragraphStyle = record.Support == "character"
                    ? Resolve(doc, record, true)
                    : Resolve(doc, record, false);
                if (record.Support == "character")
                {
                    characterStyle = paragraphStyle;
                }

                if (paragraphStyle == null)
                {
                    return Fail("文档里找不到样式 " + record.Structure + "。");
                }

                if (paragraph != null
                    && paragraph.ContainsKey("outline_level")
                    && IsFixedHeading(record.Structure))
                {
                    return Fail("Word 中系统内置的标题 1 到标题 9 不允许改大纲层级。");
                }

                bool wantText = text != null && text.Count > 0 && record.Support != "character";
                bool createdCharacter = false;
                if (wantText)
                {
                    string ensureError = EnsureCharacter(doc, paragraphStyle, out characterStyle, out createdCharacter);
                    if (ensureError != null)
                    {
                        return Fail(ensureError);
                    }
                }

                FontSnap paragraphFont = wantText ? FontSnap.Capture(paragraphStyle, text) : null;
                FontSnap characterFont = wantText ? FontSnap.Capture(characterStyle, text) : null;
                ParagraphSnap paragraphSnap = record.Support != "character" && paragraph != null
                    ? ParagraphSnap.Capture(paragraphStyle, paragraph)
                    : null;
                try
                {
                    if (wantText)
                    {
                        WriteFont(paragraphStyle, text);
                        WriteFont(characterStyle, text);
                    }
                    else if (text != null && record.Support == "character")
                    {
                        WriteFont(paragraphStyle, text);
                    }

                    if (record.Support != "character" && paragraph != null)
                    {
                        WriteParagraph(paragraphStyle, paragraph);
                    }
                }
                catch (Exception ex)
                {
                    if (paragraphFont != null)
                    {
                        paragraphFont.Restore(paragraphStyle);
                    }

                    if (characterFont != null)
                    {
                        characterFont.Restore(characterStyle);
                    }

                    if (paragraphSnap != null)
                    {
                        paragraphSnap.Restore(paragraphStyle);
                    }

                    if (createdCharacter)
                    {
                        SafeDelete(characterStyle);
                    }

                    return Fail(ex.Message);
                }
            }
            catch (Exception ex)
            {
                return Fail(ex.Message);
            }

            var ok = new Dictionary<string, object>
            {
                ["ok"] = true,
                ["structure"] = record.Structure
            };
            if (warnings.Count > 0)
            {
                ok["format_warnings"] = warnings;
            }

            return ok;
        }

        private static Dictionary<string, object> Create(
            Word.Document doc,
            string name,
            Dictionary<string, object> text,
            Dictionary<string, object> paragraph)
        {
            bool wantText = text != null && text.Count > 0;
            bool wantParagraph = paragraph != null && paragraph.Count > 0;
            if (wantText && !wantParagraph)
            {
                return CreateCharacter(doc, name, text);
            }

            return CreateParagraph(doc, name, wantText ? text : null, paragraph);
        }

        private static Dictionary<string, object> CreateCharacter(
            Word.Document doc,
            string name,
            Dictionary<string, object> text)
        {
            Word.Style style = null;
            try
            {
                style = doc.Styles.Add(name, Word.WdStyleType.wdStyleTypeCharacter);
                WriteFont(style, text);
                HideCharacter(style);
            }
            catch (Exception ex)
            {
                SafeDelete(style);
                return Fail(ex.Message);
            }

            return new Dictionary<string, object>
            {
                ["ok"] = true,
                ["structure"] = name,
                ["support"] = "character"
            };
        }

        private static Dictionary<string, object> CreateParagraph(
            Word.Document doc,
            string name,
            Dictionary<string, object> text,
            Dictionary<string, object> paragraph)
        {
            Word.Style paragraphStyle = null;
            Word.Style characterStyle = null;
            bool createdCharacter = false;
            try
            {
                paragraphStyle = doc.Styles.Add(name, Word.WdStyleType.wdStyleTypeParagraph);
                TryBaseStyle(paragraphStyle);
                if (text != null)
                {
                    string ensureError = EnsureCharacter(doc, paragraphStyle, out characterStyle, out createdCharacter);
                    if (ensureError != null)
                    {
                        SafeDelete(paragraphStyle);
                        return Fail(ensureError);
                    }

                    WriteFont(paragraphStyle, text);
                    WriteFont(characterStyle, text);
                }

                if (paragraph != null)
                {
                    WriteParagraph(paragraphStyle, paragraph);
                }
            }
            catch (Exception ex)
            {
                if (createdCharacter)
                {
                    SafeDelete(characterStyle);
                }

                SafeDelete(paragraphStyle);
                return Fail(ex.Message);
            }

            return new Dictionary<string, object>
            {
                ["ok"] = true,
                ["structure"] = name,
                ["support"] = text != null ? "both" : "paragraph"
            };
        }

        private static string EnsureCharacter(Word.Document doc, Word.Style paragraphStyle, out Word.Style characterStyle, out bool created)
        {
            characterStyle = FindCharacter(doc, paragraphStyle);
            created = false;
            if (characterStyle == null)
            {
                string characterName = paragraphStyle.NameLocal + " 字符";
                try
                {
                    characterStyle = doc.Styles.Add(characterName, Word.WdStyleType.wdStyleTypeCharacter);
                    created = true;
                }
                catch (Exception)
                {
                    characterStyle = Find(doc, characterName);
                }
            }

            if (characterStyle == null)
            {
                return "没能建出字符样式。";
            }

            try
            {
                object link = characterStyle;
                paragraphStyle.set_LinkStyle(ref link);
            }
            catch (Exception)
            {
                // 下面核对；连不上就回滚。
            }

            if (!Linked(paragraphStyle, characterStyle))
            {
                if (created)
                {
                    SafeDelete(characterStyle);
                    characterStyle = null;
                    created = false;
                }

                return "字符样式没能连上，已撤掉这次写入。";
            }

            HideCharacter(characterStyle);
            return null;
        }

        private static Word.Style FindCharacter(Word.Document doc, Word.Style paragraphStyle)
        {
            string local = paragraphStyle.NameLocal;
            Word.Style found = Find(doc, local + " 字符");
            if (found != null && found.Type == Word.WdStyleType.wdStyleTypeCharacter)
            {
                return found;
            }

            found = Find(doc, local + " Char");
            if (found != null && found.Type == Word.WdStyleType.wdStyleTypeCharacter)
            {
                return found;
            }

            try
            {
                Word.Style link = paragraphStyle.get_LinkStyle();
                if (link != null && link.Type == Word.WdStyleType.wdStyleTypeCharacter)
                {
                    return link;
                }
            }
            catch (Exception)
            {
                // 还没有链接。
            }

            return null;
        }

        private static bool Linked(Word.Style paragraphStyle, Word.Style characterStyle)
        {
            try
            {
                Word.Style paragraphLink = paragraphStyle.get_LinkStyle();
                Word.Style characterLink = characterStyle.get_LinkStyle();
                return paragraphLink != null
                    && characterLink != null
                    && paragraphLink.NameLocal == characterStyle.NameLocal
                    && characterLink.NameLocal == paragraphStyle.NameLocal;
            }
            catch (Exception)
            {
                return false;
            }
        }

        private static void HideCharacter(Word.Style style)
        {
            if (style == null)
            {
                return;
            }

            try
            {
                style.GetType().InvokeMember(
                    "Visibility",
                    System.Reflection.BindingFlags.SetProperty,
                    null,
                    style,
                    new object[] { 1 });
            }
            catch (Exception)
            {
                // 宿主不认隐藏时，样式仍在，快速样式里不一定收起。
            }

            try
            {
                style.QuickStyle = false;
            }
            catch (Exception)
            {
                // 没有快速样式开关就跳过。
            }
        }

        private static void DropAutoCharacter(Word.Document doc, Word.Style paragraphStyle)
        {
            Word.Style characterStyle = FindCharacter(doc, paragraphStyle);
            if (characterStyle == null)
            {
                return;
            }

            try
            {
                object missing = Type.Missing;
                paragraphStyle.set_LinkStyle(ref missing);
            }
            catch (Exception)
            {
                // 断开失败也继续删自动带出来的字符样式。
            }

            SafeDelete(characterStyle);
        }

        private static void TryBaseStyle(Word.Style paragraphStyle)
        {
            try
            {
                object normal = "正文";
                paragraphStyle.set_BaseStyle(ref normal);
            }
            catch (Exception)
            {
                try
                {
                    object normal = "Normal";
                    paragraphStyle.set_BaseStyle(ref normal);
                }
                catch (Exception)
                {
                    // 基于正文失败时，样式仍按新建的段落样式留下。
                }
            }
        }

        private sealed class FontSnap
        {
            private bool hasSize;
            private float size;
            private bool hasBold;
            private int bold;
            private bool hasItalic;
            private int italic;
            private bool hasUnderline;
            private Word.WdUnderline underline;
            private bool hasColor;
            private Word.WdColor color;
            private bool hasName;
            private string name;
            private string nameFarEast;

            public static FontSnap Capture(Word.Style style, Dictionary<string, object> text)
            {
                var snap = new FontSnap();
                if (style == null || text == null)
                {
                    return snap;
                }

                Word.Font font = style.Font;
                if (text.ContainsKey("font_size"))
                {
                    snap.hasSize = true;
                    snap.size = font.Size;
                }

                if (text.ContainsKey("is_bold"))
                {
                    snap.hasBold = true;
                    snap.bold = font.Bold;
                }

                if (text.ContainsKey("is_italic"))
                {
                    snap.hasItalic = true;
                    snap.italic = font.Italic;
                }

                if (text.ContainsKey("has_underline"))
                {
                    snap.hasUnderline = true;
                    snap.underline = font.Underline;
                }

                if (text.ContainsKey("font_color"))
                {
                    snap.hasColor = true;
                    snap.color = font.Color;
                }

                if (text.ContainsKey("font_name"))
                {
                    snap.hasName = true;
                    snap.name = font.Name;
                    snap.nameFarEast = font.NameFarEast;
                }

                return snap;
            }

            public void Restore(Word.Style style)
            {
                if (style == null)
                {
                    return;
                }

                Word.Font font = style.Font;
                if (hasSize)
                {
                    font.Size = size;
                }

                if (hasBold)
                {
                    font.Bold = bold;
                }

                if (hasItalic)
                {
                    font.Italic = italic;
                }

                if (hasUnderline)
                {
                    font.Underline = underline;
                }

                if (hasColor)
                {
                    font.Color = color;
                }

                if (hasName)
                {
                    font.Name = name;
                    font.NameFarEast = nameFarEast;
                }
            }
        }

        private sealed class ParagraphSnap
        {
            private bool hasBefore;
            private float before;
            private bool hasAfter;
            private float after;
            private bool hasLeft;
            private float left;
            private bool hasOutline;
            private Word.WdOutlineLevel outline;

            public static ParagraphSnap Capture(Word.Style style, Dictionary<string, object> paragraph)
            {
                var snap = new ParagraphSnap();
                if (style == null || paragraph == null)
                {
                    return snap;
                }

                Word.ParagraphFormat format = style.ParagraphFormat;
                if (paragraph.ContainsKey("space_before"))
                {
                    snap.hasBefore = true;
                    snap.before = format.SpaceBefore;
                }

                if (paragraph.ContainsKey("space_after"))
                {
                    snap.hasAfter = true;
                    snap.after = format.SpaceAfter;
                }

                if (paragraph.ContainsKey("left_indent"))
                {
                    snap.hasLeft = true;
                    snap.left = format.LeftIndent;
                }

                if (paragraph.ContainsKey("outline_level"))
                {
                    snap.hasOutline = true;
                    snap.outline = format.OutlineLevel;
                }

                return snap;
            }

            public void Restore(Word.Style style)
            {
                if (style == null)
                {
                    return;
                }

                Word.ParagraphFormat format = style.ParagraphFormat;
                if (hasBefore)
                {
                    format.SpaceBefore = before;
                }

                if (hasAfter)
                {
                    format.SpaceAfter = after;
                }

                if (hasLeft)
                {
                    format.LeftIndent = left;
                }

                if (hasOutline)
                {
                    format.OutlineLevel = outline;
                }
            }
        }

        private static bool LinksMatch(Word.Document doc, string paragraphName, string characterName)
        {
            OpenXmlPackage package;
            try
            {
                package = OpenXmlPackage.LoadFull(doc);
            }
            catch (Exception)
            {
                return false;
            }

            XElement styles = package.FindPartRoot("styles.xml", OpenXmlPackage.W + "styles");
            if (styles == null)
            {
                return false;
            }

            XNamespace ns = OpenXmlPackage.W;
            string paraId = null;
            string charId = null;
            string paraLink = null;
            string charLink = null;
            foreach (XElement style in styles.Elements(ns + "style"))
            {
                string styleName = (string)style.Element(ns + "name")?.Attribute(ns + "val");
                if (styleName == paragraphName)
                {
                    paraId = (string)style.Attribute(ns + "styleId");
                    paraLink = (string)style.Element(ns + "link")?.Attribute(ns + "val");
                }
                else if (styleName == characterName)
                {
                    charId = (string)style.Attribute(ns + "styleId");
                    charLink = (string)style.Element(ns + "link")?.Attribute(ns + "val");
                }
            }

            return !string.IsNullOrEmpty(paraId)
                && !string.IsNullOrEmpty(charId)
                && paraLink == charId
                && charLink == paraId;
        }

        private static void WriteFont(Word.Style style, Dictionary<string, object> text)
        {
            if (style == null || text == null)
            {
                return;
            }

            Word.Font font = style.Font;
            if (text.TryGetValue("font_name", out object face) && face != null)
            {
                font.Name = face.ToString();
                font.NameFarEast = face.ToString();
            }

            if (text.TryGetValue("font_size", out object size))
            {
                font.Size = Convert.ToSingle(size, CultureInfo.InvariantCulture);
            }

            if (text.TryGetValue("font_color", out object color) && color != null)
            {
                font.Color = (Word.WdColor)RgbToBgr(color.ToString());
            }

            if (text.TryGetValue("is_bold", out object bold))
            {
                font.Bold = IsTrueValue(bold) ? -1 : 0;
            }

            if (text.TryGetValue("is_italic", out object italic))
            {
                font.Italic = IsTrueValue(italic) ? -1 : 0;
            }

            if (text.TryGetValue("has_underline", out object underline))
            {
                font.Underline = IsTrueValue(underline)
                    ? Word.WdUnderline.wdUnderlineSingle
                    : Word.WdUnderline.wdUnderlineNone;
            }
        }

        private static bool IsFixedHeading(string structure)
        {
            if (string.IsNullOrEmpty(structure) || !structure.StartsWith("heading_", StringComparison.Ordinal))
            {
                return false;
            }

            int level;
            string n = structure.Substring("heading_".Length);
            return int.TryParse(n, NumberStyles.Integer, CultureInfo.InvariantCulture, out level)
                && level >= 1
                && level <= 9
                && n == level.ToString(CultureInfo.InvariantCulture);
        }

        private static void WriteParagraph(Word.Style style, Dictionary<string, object> paragraph)
        {
            if (style == null || paragraph == null)
            {
                return;
            }

            Word.ParagraphFormat format = style.ParagraphFormat;
            if (paragraph.TryGetValue("space_before", out object before))
            {
                format.SpaceBefore = Convert.ToSingle(before, CultureInfo.InvariantCulture);
            }

            if (paragraph.TryGetValue("space_after", out object after))
            {
                format.SpaceAfter = Convert.ToSingle(after, CultureInfo.InvariantCulture);
            }

            if (paragraph.TryGetValue("left_indent", out object left))
            {
                format.LeftIndent = Convert.ToSingle(left, CultureInfo.InvariantCulture);
            }

            if (paragraph.TryGetValue("outline_level", out object level) && level != null)
            {
                Word.WdOutlineLevel expected = (Word.WdOutlineLevel)Enum.Parse(
                    typeof(Word.WdOutlineLevel),
                    level.ToString(),
                    true);
                format.OutlineLevel = expected;
                if (format.OutlineLevel != expected)
                {
                    throw new InvalidOperationException("这个样式的大纲级别写不进去。");
                }
            }
        }

        private static int RgbToBgr(string raw)
        {
            string hex = raw.Trim().TrimStart('#');
            int rgb = int.Parse(hex, NumberStyles.HexNumber, CultureInfo.InvariantCulture);
            return ((rgb & 0xFF) << 16) | (rgb & 0xFF00) | ((rgb >> 16) & 0xFF);
        }

        private static Word.Style Find(Word.Document doc, string name)
        {
            if (string.IsNullOrWhiteSpace(name))
            {
                return null;
            }

            try
            {
                return doc.Styles[name];
            }
            catch (Exception)
            {
                return null;
            }
        }

        private static List<string> StyleNames(Word.Document doc)
        {
            var names = new List<string>();
            foreach (Word.Style style in doc.Styles)
            {
                try
                {
                    names.Add(style.NameLocal);
                }
                catch (Exception)
                {
                    // 个别内置样式取不到名字。
                }
            }

            return names;
        }

        private static void SafeDelete(Word.Style style)
        {
            if (style == null)
            {
                return;
            }

            try
            {
                style.Delete();
            }
            catch (Exception)
            {
                // 删不掉就留在失败回包里，由调用方看到新建失败。
            }
        }

        private static bool IsTrue(Dictionary<string, object> args, string key)
        {
            return args != null && args.TryGetValue(key, out object value) && IsTrueValue(value);
        }

        private static bool IsTrueValue(object value)
        {
            if (value is bool bit)
            {
                return bit;
            }

            string text = value == null ? "" : value.ToString();
            return text == "1" || text.Equals("true", StringComparison.OrdinalIgnoreCase);
        }

        private static string Arg(Dictionary<string, object> args, string key)
        {
            if (args == null || !args.TryGetValue(key, out object value) || value == null)
            {
                return null;
            }

            return value.ToString().Trim();
        }

        private static Word.Style Resolve(Word.Document doc, StructureStyleRecord record, bool character)
        {
            if (record == null)
            {
                return null;
            }

            var names = new List<string>();
            if (character)
            {
                AddName(names, record.CharacterStyleName);
            }

            AddName(names, record.StyleName);
            foreach (string name in StructureStyleCatalog.LookupNames(record))
            {
                AddName(names, name);
            }

            foreach (string name in names)
            {
                Word.Style style = Find(doc, name);
                if (style == null)
                {
                    continue;
                }

                bool characterStyle = style.Type == Word.WdStyleType.wdStyleTypeCharacter;
                if (character ? characterStyle : !characterStyle)
                {
                    return style;
                }
            }

            return null;
        }

        private static void AddName(List<string> names, string name)
        {
            if (!string.IsNullOrWhiteSpace(name) && !names.Contains(name))
            {
                names.Add(name);
            }
        }

        private static Dictionary<string, object> AsDict(Dictionary<string, object> args, string key)
        {
            if (args == null || !args.TryGetValue(key, out object value) || value == null)
            {
                return null;
            }

            if (value is Dictionary<string, object> dict)
            {
                return dict;
            }

            if (value is JObject jo)
            {
                return jo.ToObject<Dictionary<string, object>>();
            }

            if (value is JToken token && token.Type == JTokenType.Object)
            {
                return token.ToObject<Dictionary<string, object>>();
            }

            return null;
        }

        private static Dictionary<string, object> Fail(string error)
        {
            return new Dictionary<string, object> { ["ok"] = false, ["error"] = error };
        }
    }
}
