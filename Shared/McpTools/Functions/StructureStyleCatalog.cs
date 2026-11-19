using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text.RegularExpressions;
using System.Xml.Linq;

namespace WordAddIn1
{
    /// <summary>样式目录。只读 styles.xml 里写上的项。正文先读 docDefaults，缺的键再读正文样式自己写下的项。</summary>
    public sealed class StructureStyleRecord
    {
        public string Structure { get; set; }
        public string Origin { get; set; }
        public string Support { get; set; }
        public string StyleId { get; set; }
        public string StyleName { get; set; }
        public string CharacterStyleId { get; set; }
        public string CharacterStyleName { get; set; }
        public Dictionary<string, object> TextFormat { get; set; }
        public Dictionary<string, object> ParagraphFormat { get; set; }

        public Dictionary<string, object> ToCatalogRow()
        {
            var row = new Dictionary<string, object>
            {
                ["structure"] = Structure,
                ["origin"] = Origin,
                ["support"] = Support,
                ["apply"] = Support
            };
            if (TextFormat != null && TextFormat.Count > 0)
            {
                row["text_format"] = TextFormat;
            }

            if (ParagraphFormat != null && ParagraphFormat.Count > 0)
            {
                row["paragraph_format"] = ParagraphFormat;
            }

            return row;
        }

        public Dictionary<string, object> FormatsForScope(string scope)
        {
            var row = new Dictionary<string, object>();
            if (Support == "character" || scope == "text")
            {
                if (TextFormat != null && TextFormat.Count > 0)
                {
                    row["text_format"] = TextFormat;
                }

                return row;
            }

            if (Support == "paragraph")
            {
                if (ParagraphFormat != null && ParagraphFormat.Count > 0)
                {
                    row["paragraph_format"] = ParagraphFormat;
                }

                return row;
            }

            if (TextFormat != null && TextFormat.Count > 0)
            {
                row["text_format"] = TextFormat;
            }

            if (ParagraphFormat != null && ParagraphFormat.Count > 0)
            {
                row["paragraph_format"] = ParagraphFormat;
            }

            return row;
        }
    }

    public static class StructureStyleCatalog
    {
        private static readonly XNamespace W = OpenXmlPackage.W;

        private static readonly Dictionary<string, string> Aliases =
            new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                ["Normal"] = "正文",
                ["正文"] = "Normal",
                ["Title"] = "标题",
                ["标题"] = "Title",
                ["Subtitle"] = "副标题",
                ["副标题"] = "Subtitle",
                ["List Paragraph"] = "列表段落",
                ["列表段落"] = "List Paragraph",
                ["Hyperlink"] = "超链接",
                ["超链接"] = "Hyperlink",
                ["Caption"] = "题注",
                ["题注"] = "Caption",
                ["Header"] = "页眉",
                ["页眉"] = "Header",
                ["Footer"] = "页脚",
                ["页脚"] = "Footer",
                ["Intense Emphasis"] = "明显强调",
                ["明显强调"] = "Intense Emphasis",
                ["Default Paragraph Font"] = "默认段落字体",
                ["默认段落字体"] = "Default Paragraph Font",
                ["Quote"] = "引用",
                ["引用"] = "Quote",
                ["Intense Quote"] = "明显引用",
                ["明显引用"] = "Intense Quote"
            };

        public static readonly Dictionary<string, string> ApplyRules =
            new Dictionary<string, string>
            {
                ["both"] = "文字格式作用在定位得到的整个范围上。段落格式只在被整段盖住的段上生效。只盖住半截时，段落格式没有发生作用，要想段落格式生效，需要套整段。一次盖住多段、两头又多出半截时，段落格式只在盖住完整的段落上生效，半截不生效，要想半截的段落格式生效，需要套整段。",
                ["paragraph"] = "只支持段落。被整段盖住时只改段落格式，字上已有的字符样式不动。只盖住半截时段落格式没有发生作用，要想段落格式生效，需要套整段。没有字符格式可套，所以这次套用什么也不会发生。",
                ["character"] = "只支持字符。文字格式作用在定位得到的整个范围上，段落格式不动。"
            };

        public static List<StructureStyleRecord> Read(XElement stylesRoot)
        {
            var raw = new List<RawStyle>();
            if (stylesRoot == null)
            {
                return new List<StructureStyleRecord>();
            }

            XElement defaults = stylesRoot.Element(W + "docDefaults");
            foreach (XElement style in stylesRoot.Elements(W + "style"))
            {
                raw.Add(RawStyle.Read(style));
            }

            var byId = new Dictionary<string, RawStyle>(StringComparer.Ordinal);
            foreach (RawStyle style in raw)
            {
                if (!string.IsNullOrEmpty(style.Id) && !byId.ContainsKey(style.Id))
                {
                    byId[style.Id] = style;
                }
            }

            var consumed = new HashSet<string>(StringComparer.Ordinal);
            var rows = new List<StructureStyleRecord>();
            foreach (RawStyle style in raw)
            {
                if (style.Type != "paragraph" || consumed.Contains(style.Id))
                {
                    continue;
                }

                RawStyle linked = null;
                if (!string.IsNullOrEmpty(style.Link)
                    && byId.TryGetValue(style.Link, out RawStyle other)
                    && other.Type == "character"
                    && string.Equals(other.Link, style.Id, StringComparison.Ordinal))
                {
                    linked = other;
                    consumed.Add(other.Id);
                }

                rows.Add(ToRecord(style, linked, defaults));
            }

            foreach (RawStyle style in raw)
            {
                if (style.Type == "character" && !consumed.Contains(style.Id))
                {
                    rows.Add(ToCharacter(style));
                }
            }

            return rows;
        }

        public static StructureStyleRecord Find(IList<StructureStyleRecord> rows, string structure)
        {
            if (rows == null || string.IsNullOrWhiteSpace(structure))
            {
                return null;
            }

            string key = structure.Trim();
            foreach (StructureStyleRecord row in rows)
            {
                if (string.Equals(row.Structure, key, StringComparison.OrdinalIgnoreCase)
                    || string.Equals(row.StyleName, key, StringComparison.OrdinalIgnoreCase)
                    || string.Equals(row.CharacterStyleName, key, StringComparison.OrdinalIgnoreCase)
                    || string.Equals(Alias(row.StyleName), key, StringComparison.OrdinalIgnoreCase))
                {
                    return row;
                }
            }

            string builtin = BuiltinStructure(key);
            if (builtin == null)
            {
                return null;
            }

            foreach (StructureStyleRecord row in rows)
            {
                if (string.Equals(row.Structure, builtin, StringComparison.Ordinal))
                {
                    return row;
                }
            }

            return null;
        }

        public static string BuiltinStructure(string name)
        {
            if (string.IsNullOrWhiteSpace(name))
            {
                return null;
            }

            string n = name.Trim();
            if (n.Equals("Normal", StringComparison.OrdinalIgnoreCase) || n == "正文")
            {
                return "body";
            }

            if (n.Equals("Title", StringComparison.OrdinalIgnoreCase) || n == "标题")
            {
                return "title";
            }

            if (n.Equals("Subtitle", StringComparison.OrdinalIgnoreCase) || n == "副标题")
            {
                return "subtitle";
            }

            Match heading = Regex.Match(n, @"^(?:heading|标题)\s*([1-9])$", RegexOptions.IgnoreCase);
            if (heading.Success)
            {
                return "heading_" + heading.Groups[1].Value;
            }

            return null;
        }

        public static IEnumerable<string> LookupNames(StructureStyleRecord row)
        {
            if (row == null)
            {
                yield break;
            }

            if (!string.IsNullOrEmpty(row.StyleName))
            {
                yield return row.StyleName;
                string alias = Alias(row.StyleName);
                if (!string.IsNullOrEmpty(alias))
                {
                    yield return alias;
                }
            }

            if (!string.IsNullOrEmpty(row.CharacterStyleName))
            {
                yield return row.CharacterStyleName;
            }

            if (row.Structure == "body")
            {
                yield return "正文";
                yield return "Normal";
            }
            else if (row.Structure == "title")
            {
                yield return "标题";
                yield return "Title";
            }
            else if (row.Structure == "subtitle")
            {
                yield return "副标题";
                yield return "Subtitle";
            }
            else if (!string.IsNullOrEmpty(row.Structure) && row.Structure.StartsWith("heading_", StringComparison.Ordinal))
            {
                string n = row.Structure.Substring("heading_".Length);
                yield return "标题 " + n;
                yield return "heading " + n;
                if (!string.IsNullOrEmpty(row.CharacterStyleName))
                {
                    yield return "标题 " + n + " 字符";
                    yield return "heading " + n + " Char";
                }
            }
        }

        private static StructureStyleRecord ToRecord(RawStyle style, RawStyle linked, XElement defaults)
        {
            string builtin = BuiltinStructure(style.Name);
            bool body = builtin == "body" && linked == null;
            var record = new StructureStyleRecord
            {
                Structure = builtin ?? DisplayName(style.Name),
                Origin = style.Custom && builtin == null ? "custom" : "builtin",
                Support = "both",
                StyleId = style.Id,
                StyleName = style.Name,
                CharacterStyleId = linked == null ? null : linked.Id,
                CharacterStyleName = linked == null ? null : linked.Name
            };
            if (body)
            {
                record.Support = "both";
                record.TextFormat = ReadText(defaults == null ? null : defaults.Element(W + "rPrDefault")?.Element(W + "rPr"));
                record.ParagraphFormat = ReadParagraph(defaults == null ? null : defaults.Element(W + "pPrDefault")?.Element(W + "pPr"));
                FillMissing(record.TextFormat, ReadText(style.RPr));
                FillMissing(record.ParagraphFormat, ReadParagraph(style.PPr));
                return record;
            }

            Dictionary<string, object> text = ReadText(linked != null ? linked.RPr : null);
            if (text.Count == 0)
            {
                text = ReadText(style.RPr);
            }

            record.ParagraphFormat = ReadParagraph(style.PPr);
            if (text.Count > 0)
            {
                record.Support = "both";
                record.TextFormat = text;
            }
            else
            {
                record.Support = "paragraph";
                record.TextFormat = null;
            }

            return record;
        }

        private static StructureStyleRecord ToCharacter(RawStyle style)
        {
            return new StructureStyleRecord
            {
                Structure = DisplayName(style.Name),
                Origin = style.Custom ? "custom" : "builtin",
                Support = "character",
                StyleId = style.Id,
                StyleName = style.Name,
                CharacterStyleId = style.Id,
                CharacterStyleName = style.Name,
                TextFormat = ReadText(style.RPr)
            };
        }

        private static string DisplayName(string wName)
        {
            if (string.IsNullOrEmpty(wName))
            {
                return wName;
            }

            if (Aliases.TryGetValue(wName, out string alias) && !IsEnglishToken(alias))
            {
                return alias;
            }

            return wName;
        }

        private static bool IsEnglishToken(string value)
        {
            foreach (char c in value)
            {
                if (c > 127)
                {
                    return false;
                }
            }

            return true;
        }

        private static string Alias(string name)
        {
            if (string.IsNullOrEmpty(name))
            {
                return null;
            }

            string alias;
            return Aliases.TryGetValue(name, out alias) ? alias : null;
        }

        /// <summary>同一支样式的中英文名。styles.xml 用英文，样式库用中文。</summary>
        internal static IEnumerable<string> EquivalentNames(string name)
        {
            if (string.IsNullOrEmpty(name))
            {
                yield break;
            }

            yield return name;
            string alias = Alias(name);
            if (!string.IsNullOrEmpty(alias))
            {
                yield return alias;
            }

            Match heading = Regex.Match(
                name,
                @"^(标题|heading)\s+(\d+)(\s*字符|\s+Char)?$",
                RegexOptions.IgnoreCase);
            if (!heading.Success)
            {
                yield break;
            }

            string index = heading.Groups[2].Value;
            bool character = heading.Groups[3].Success;
            bool chinese = heading.Groups[1].Value.StartsWith("标", StringComparison.Ordinal);
            yield return chinese
                ? (character ? "heading " + index + " Char" : "heading " + index)
                : (character ? "标题 " + index + " 字符" : "标题 " + index);
        }

        private static Dictionary<string, object> ReadText(XElement rPr)
        {
            var format = new Dictionary<string, object>();
            if (rPr == null)
            {
                return format;
            }

            XElement size = rPr.Element(W + "sz");
            int half;
            if (size != null && int.TryParse((string)size.Attribute(W + "val"), out half) && half > 0)
            {
                format["font_size"] = half / 2.0;
            }

            XElement color = rPr.Element(W + "color");
            if (color != null)
            {
                string val = (string)color.Attribute(W + "val");
                if (!string.IsNullOrEmpty(val) && !val.Equals("auto", StringComparison.OrdinalIgnoreCase))
                {
                    format["font_color"] = val;
                }
            }

            XElement fonts = rPr.Element(W + "rFonts");
            if (fonts != null)
            {
                string face = (string)fonts.Attribute(W + "eastAsia")
                    ?? (string)fonts.Attribute(W + "ascii")
                    ?? (string)fonts.Attribute(W + "hAnsi");
                if (!string.IsNullOrEmpty(face))
                {
                    format["font_name"] = face;
                }
            }

            if (Flag(rPr.Element(W + "b")))
            {
                format["is_bold"] = true;
            }

            if (Flag(rPr.Element(W + "i")))
            {
                format["is_italic"] = true;
            }

            XElement underline = rPr.Element(W + "u");
            if (underline != null)
            {
                string val = (string)underline.Attribute(W + "val");
                if (!string.IsNullOrEmpty(val) && !val.Equals("none", StringComparison.OrdinalIgnoreCase))
                {
                    format["has_underline"] = true;
                }
            }

            return format;
        }

        private static Dictionary<string, object> ReadParagraph(XElement pPr)
        {
            var format = new Dictionary<string, object>();
            if (pPr == null)
            {
                return format;
            }

            XElement spacing = pPr.Element(W + "spacing");
            if (spacing != null)
            {
                AddTwips(format, "space_before", (string)spacing.Attribute(W + "before"));
                AddTwips(format, "space_after", (string)spacing.Attribute(W + "after"));
            }

            XElement ind = pPr.Element(W + "ind");
            if (ind != null)
            {
                AddTwips(format, "left_indent", (string)ind.Attribute(W + "left"));
                AddTwips(format, "right_indent", (string)ind.Attribute(W + "right"));
                AddTwips(format, "first_line_indent", (string)ind.Attribute(W + "firstLine"));
            }

            XElement outline = pPr.Element(W + "outlineLvl");
            int level;
            if (outline != null && int.TryParse((string)outline.Attribute(W + "val"), NumberStyles.Integer, CultureInfo.InvariantCulture, out level)
                && level >= 0 && level <= 8)
            {
                format["outline_level"] = "wdOutlineLevel" + (level + 1);
            }

            return format;
        }

        private static void FillMissing(Dictionary<string, object> target, Dictionary<string, object> extra)
        {
            if (target == null || extra == null)
            {
                return;
            }

            foreach (KeyValuePair<string, object> pair in extra)
            {
                if (!target.ContainsKey(pair.Key))
                {
                    target[pair.Key] = pair.Value;
                }
            }
        }

        private static void AddTwips(Dictionary<string, object> format, string key, string raw)
        {
            int twips;
            if (int.TryParse(raw, NumberStyles.Integer, CultureInfo.InvariantCulture, out twips))
            {
                format[key] = twips / 20.0;
            }
        }

        private static bool Flag(XElement element)
        {
            if (element == null)
            {
                return false;
            }

            string val = (string)element.Attribute(W + "val");
            return string.IsNullOrEmpty(val) || val == "1" || val.Equals("true", StringComparison.OrdinalIgnoreCase) || val == "on";
        }

        private sealed class RawStyle
        {
            public string Id;
            public string Type;
            public string Name;
            public string Link;
            public bool Custom;
            public XElement RPr;
            public XElement PPr;

            public static RawStyle Read(XElement style)
            {
                return new RawStyle
                {
                    Id = (string)style.Attribute(W + "styleId"),
                    Type = (string)style.Attribute(W + "type"),
                    Name = (string)style.Element(W + "name")?.Attribute(W + "val"),
                    Link = (string)style.Element(W + "link")?.Attribute(W + "val"),
                    Custom = style.Attribute(W + "customStyle") != null,
                    RPr = style.Element(W + "rPr"),
                    PPr = style.Element(W + "pPr")
                };
            }
        }
    }
}
