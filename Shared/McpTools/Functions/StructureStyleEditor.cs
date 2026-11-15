using System;
using System.Collections.Generic;
using System.Globalization;
using System.Xml.Linq;
using Newtonsoft.Json.Linq;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    /// <summary>改样式定义，或新建一对双支持样式。</summary>
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

                return CreatePair(doc, name.Trim(), text, paragraph);
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

            if (record.Support == "paragraph" && text != null && text.Count > 0)
            {
                warnings.Add("该样式只支持段落格式，已忽略 text_format。");
                text = null;
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
                    && IsFixedHeading(record.Structure)
                    && !WpsChartExtract.IsWpsDocument(doc))
                {
                    return Fail("Word 中系统内置的标题 1 到标题 9 不允许改大纲层级。");
                }

                if (record.Support != "paragraph" && text != null)
                {
                    WriteFont(paragraphStyle, text);
                    if (characterStyle != null && !object.ReferenceEquals(characterStyle, paragraphStyle))
                    {
                        WriteFont(characterStyle, text);
                    }
                }

                if (record.Support != "character" && paragraph != null)
                {
                    WriteParagraph(paragraphStyle, paragraph);
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

        private static Dictionary<string, object> CreatePair(
            Word.Document doc,
            string name,
            Dictionary<string, object> text,
            Dictionary<string, object> paragraph)
        {
            string characterName = name + " 字符";
            Word.Style paragraphStyle = null;
            Word.Style characterStyle = null;
            try
            {
                paragraphStyle = doc.Styles.Add(name, Word.WdStyleType.wdStyleTypeParagraph);
                characterStyle = Find(doc, characterName);
                if (characterStyle == null)
                {
                    try
                    {
                        characterStyle = doc.Styles.Add(characterName, Word.WdStyleType.wdStyleTypeCharacter);
                    }
                    catch (Exception)
                    {
                        characterStyle = Find(doc, characterName);
                    }
                }

                if (characterStyle == null)
                {
                    SafeDelete(paragraphStyle);
                    return Fail("没能建出成对的字符样式。");
                }
                try
                {
                    object normal = "正文";
                    paragraphStyle.set_BaseStyle(ref normal);
                }
                catch (Exception)
                {
                    object normal = "Normal";
                    paragraphStyle.set_BaseStyle(ref normal);
                }

                if (text != null)
                {
                    WriteFont(paragraphStyle, text);
                    WriteFont(characterStyle, text);
                }

                if (paragraph != null)
                {
                    WriteParagraph(paragraphStyle, paragraph);
                }

                try
                {
                    object link = characterStyle;
                    paragraphStyle.set_LinkStyle(ref link);
                }
                catch (Exception)
                {
                    // 下面用 OOXML 核对；连不上就删掉。
                }

                if (!LinksMatch(doc, name, characterName))
                {
                    SafeDelete(characterStyle);
                    SafeDelete(paragraphStyle);
                    return Fail("新建样式没能连成一对，已删除。");
                }
            }
            catch (Exception ex)
            {
                SafeDelete(characterStyle);
                SafeDelete(paragraphStyle);
                return Fail(ex.Message);
            }

            return new Dictionary<string, object>
            {
                ["ok"] = true,
                ["structure"] = name,
                ["support"] = "both"
            };
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
