using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    /// <summary>
    /// 从已打开文档的 WordOpenXML 读段版式（对齐 KB pPr + 样式链），不逐段 COM Find。
    /// </summary>
    public static class ParaFormatOoxmlReader
    {
        private static readonly XNamespace W =
            "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

        private static readonly XNamespace Pkg =
            "http://schemas.microsoft.com/office/2006/xmlPackage";

        private static readonly Regex TagRe = new Regex("<[^>]+>", RegexOptions.Compiled);

        public sealed class BodyParagraph
        {
            public XElement Paragraph { get; set; }

            public bool Eligible { get; set; }

            public string VisibleText { get; set; }
        }

        public static bool TryReadOpenXml(Word.Document document, out XDocument xmlDoc, out string error)
        {
            xmlDoc = null;
            error = null;
            try
            {
                string raw = document?.Content?.WordOpenXML;
                if (string.IsNullOrEmpty(raw))
                {
                    error = "WordOpenXML 为空";
                    return false;
                }

                xmlDoc = XDocument.Parse(raw);
                return true;
            }
            catch (Exception ex)
            {
                error = ex.Message;
                return false;
            }
        }

        public static List<BodyParagraph> CollectBodyParagraphs(XDocument xmlDoc)
        {
            var list = new List<BodyParagraph>();
            if (xmlDoc == null)
            {
                return list;
            }

            XElement body = xmlDoc.Descendants(W + "body").FirstOrDefault();
            if (body == null)
            {
                return list;
            }

            Walk(body, eligible: true, inTable: false, list);
            return list;
        }

        public static Dictionary<string, object> ExtractEffective(
            XElement paragraph,
            Dictionary<string, StyleRec> styles,
            Dictionary<string, object> docDefaults)
        {
            var layers = new List<Dictionary<string, object>>();
            if (docDefaults != null && docDefaults.Count > 0)
            {
                layers.Add(docDefaults);
            }

            string styleId = GetParagraphStyleId(paragraph);
            if (!string.IsNullOrEmpty(styleId) && styles != null)
            {
                foreach (Dictionary<string, object> layer in StyleChainParaProps(styleId, styles))
                {
                    layers.Add(layer);
                }
            }

            XElement pPr = paragraph?.Element(W + "pPr");
            Dictionary<string, object> direct = PprToPartial(pPr);
            if (direct.Count > 0)
            {
                layers.Add(direct);
            }

            Dictionary<string, object> merged = MergeLayers(layers);
            NormalizeLineSpacing(merged);
            if (!merged.ContainsKey("alignment"))
            {
                merged["alignment"] = "inherit";
            }

            return merged;
        }

        public static Dictionary<string, StyleRec> LoadStyles(XDocument xmlDoc)
        {
            var map = new Dictionary<string, StyleRec>(StringComparer.Ordinal);
            XElement stylesRoot = FindStylesRoot(xmlDoc);
            if (stylesRoot == null)
            {
                return map;
            }

            foreach (XElement style in stylesRoot.Elements(W + "style"))
            {
                string type = (string)style.Element(W + "type")?.Attribute(W + "val")
                    ?? (string)style.Attribute(W + "type")
                    ?? "";
                if (!string.IsNullOrEmpty(type)
                    && !string.Equals(type, "paragraph", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                string id = (string)style.Attribute(W + "styleId");
                if (string.IsNullOrEmpty(id))
                {
                    continue;
                }

                string basedOn = (string)style.Element(W + "basedOn")?.Attribute(W + "val") ?? "";
                map[id] = new StyleRec
                {
                    StyleId = id,
                    BasedOn = basedOn,
                    Props = PprToPartial(style.Element(W + "pPr"))
                };
            }

            return map;
        }

        public static Dictionary<string, object> LoadDocDefaults(XDocument xmlDoc)
        {
            XElement stylesRoot = FindStylesRoot(xmlDoc);
            XElement pPr = stylesRoot
                ?.Element(W + "docDefaults")
                ?.Element(W + "pPrDefault")
                ?.Element(W + "pPr");
            return PprToPartial(pPr);
        }

        public static string NormalizeAlignText(string text)
        {
            if (string.IsNullOrEmpty(text))
            {
                return "";
            }

            string visible = text.Replace("\r", " ").Replace("\n", " ").Replace("\a", " ");
            visible = TagRe.Replace(visible, " ");
            return string.Join(" ", visible.Split((char[])null, StringSplitOptions.RemoveEmptyEntries));
        }

        public static string VisibleTextFromParagraph(XElement paragraph)
        {
            if (paragraph == null)
            {
                return "";
            }

            var sb = new StringBuilder();
            foreach (XElement t in paragraph.Descendants(W + "t"))
            {
                if (t.Ancestors(W + "del").Any() || t.Ancestors(W + "moveFrom").Any())
                {
                    continue;
                }

                sb.Append((string)t);
            }

            return NormalizeAlignText(sb.ToString());
        }

        public sealed class StyleRec
        {
            public string StyleId;
            public string BasedOn;
            public Dictionary<string, object> Props;
        }

        private static XElement FindStylesRoot(XDocument xmlDoc)
        {
            if (xmlDoc == null)
            {
                return null;
            }

            foreach (XElement part in xmlDoc.Descendants(Pkg + "part"))
            {
                string name = (string)part.Attribute(Pkg + "name") ?? "";
                if (name.IndexOf("styles.xml", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    return part.Descendants(W + "styles").FirstOrDefault();
                }
            }

            return xmlDoc.Descendants(W + "styles").FirstOrDefault();
        }

        private static void Walk(XElement element, bool eligible, bool inTable, List<BodyParagraph> list)
        {
            if (element == null)
            {
                return;
            }

            foreach (XElement child in element.Elements())
            {
                if (child.Name == W + "p")
                {
                    list.Add(new BodyParagraph
                    {
                        Paragraph = child,
                        Eligible = eligible,
                        VisibleText = VisibleTextFromParagraph(child)
                    });
                }
                else if (child.Name == W + "tbl")
                {
                    bool single = IsSingleCellTable(child);
                    Walk(child, eligible: single, inTable: true, list);
                }
                else if (child.Name == W + "sectPr")
                {
                }
                else
                {
                    Walk(child, eligible, inTable, list);
                }
            }
        }

        private static bool IsSingleCellTable(XElement tbl)
        {
            List<XElement> rows = tbl.Elements(W + "tr").ToList();
            if (rows.Count != 1)
            {
                return false;
            }

            return rows[0].Elements(W + "tc").Count() == 1;
        }

        private static string GetParagraphStyleId(XElement paragraph)
        {
            return (string)paragraph
                ?.Element(W + "pPr")
                ?.Element(W + "pStyle")
                ?.Attribute(W + "val") ?? "";
        }

        private static List<Dictionary<string, object>> StyleChainParaProps(
            string styleId,
            Dictionary<string, StyleRec> styles)
        {
            var ids = new List<string>();
            var seen = new HashSet<string>(StringComparer.Ordinal);
            string cur = styleId;
            while (!string.IsNullOrEmpty(cur) && seen.Add(cur))
            {
                ids.Add(cur);
                if (!styles.TryGetValue(cur, out StyleRec rec) || string.IsNullOrEmpty(rec.BasedOn))
                {
                    break;
                }

                cur = rec.BasedOn;
            }

            ids.Reverse();
            var layers = new List<Dictionary<string, object>>();
            foreach (string id in ids)
            {
                if (styles.TryGetValue(id, out StyleRec rec) && rec.Props != null && rec.Props.Count > 0)
                {
                    layers.Add(rec.Props);
                }
            }

            return layers;
        }

        private static Dictionary<string, object> MergeLayers(List<Dictionary<string, object>> layers)
        {
            var result = new Dictionary<string, object>(StringComparer.Ordinal);
            foreach (Dictionary<string, object> layer in layers)
            {
                if (layer == null)
                {
                    continue;
                }

                foreach (KeyValuePair<string, object> kv in layer)
                {
                    if (kv.Value == null)
                    {
                        continue;
                    }

                    if (kv.Value is string s && string.IsNullOrWhiteSpace(s) && kv.Key != "alignment")
                    {
                        continue;
                    }

                    result[kv.Key] = kv.Value;
                }
            }

            return result;
        }

        private static void NormalizeLineSpacing(Dictionary<string, object> format)
        {
            string rule = format.TryGetValue("line_spacing_rule", out object r) ? r?.ToString() : null;
            double ls = 1.0;
            if (format.TryGetValue("line_spacing", out object rawLs) && rawLs != null)
            {
                try
                {
                    ls = Convert.ToDouble(rawLs);
                }
                catch
                {
                    ls = 1.0;
                }
            }

            Dictionary<string, object> pair = ParaFormatReader.NormalizeLineSpacingPair(
                string.IsNullOrEmpty(rule) ? "single" : rule,
                ls);
            format["line_spacing_rule"] = pair["line_spacing_rule"];
            format["line_spacing"] = pair["line_spacing"];
        }

        private static Dictionary<string, object> PprToPartial(XElement pPr)
        {
            var partial = new Dictionary<string, object>(StringComparer.Ordinal);
            if (pPr == null)
            {
                return partial;
            }

            string jc = (string)pPr.Element(W + "jc")?.Attribute(W + "val");
            if (!string.IsNullOrEmpty(jc))
            {
                partial["alignment"] = MapJc(jc);
            }

            XElement ind = pPr.Element(W + "ind");
            if (ind != null)
            {
                TryAddTwip(partial, "left_indent", (string)ind.Attribute(W + "left"));
                TryAddTwip(partial, "right_indent", (string)ind.Attribute(W + "right"));
                double? first = TwipToPt((string)ind.Attribute(W + "firstLine"));
                double? hanging = TwipToPt((string)ind.Attribute(W + "hanging"));
                if (hanging.HasValue)
                {
                    partial["hanging_indent"] = hanging.Value;
                }
                else if (first.HasValue)
                {
                    partial["first_line_indent"] = first.Value;
                }
            }

            XElement spacing = pPr.Element(W + "spacing");
            if (spacing != null)
            {
                TryAddTwip(partial, "space_before", (string)spacing.Attribute(W + "before"));
                TryAddTwip(partial, "space_after", (string)spacing.Attribute(W + "after"));
                string lineRule = (string)spacing.Attribute(W + "lineRule");
                string line = (string)spacing.Attribute(W + "line");
                if (!string.IsNullOrEmpty(lineRule))
                {
                    partial["line_spacing_rule"] = MapLineRule(lineRule);
                }

                if (int.TryParse(line, out int lineVal) && lineVal > 0)
                {
                    string rule = partial.TryGetValue("line_spacing_rule", out object rr)
                        ? rr?.ToString()
                        : null;
                    if (rule == "multiple" || string.Equals(lineRule, "auto", StringComparison.OrdinalIgnoreCase)
                        || string.IsNullOrEmpty(lineRule))
                    {
                        partial["line_spacing"] = Math.Round(lineVal / 240.0, 2);
                        partial["line_spacing_rule"] = "multiple";
                    }
                    else if (rule == "exact" || rule == "at_least")
                    {
                        partial["line_spacing"] = Math.Round(lineVal / 20.0, 1);
                    }
                    else
                    {
                        partial["line_spacing"] = Math.Round(lineVal / 240.0, 2);
                        if (!partial.ContainsKey("line_spacing_rule"))
                        {
                            partial["line_spacing_rule"] = "multiple";
                        }
                    }
                }
            }

            return partial;
        }

        private static void TryAddTwip(Dictionary<string, object> dest, string key, string raw)
        {
            double? pt = TwipToPt(raw);
            if (pt.HasValue)
            {
                dest[key] = pt.Value;
            }
        }

        private static double? TwipToPt(string raw)
        {
            if (string.IsNullOrEmpty(raw) || !int.TryParse(raw, out int twip) || twip == 0)
            {
                return null;
            }

            return Math.Round(twip / 20.0, 1);
        }

        private static string MapJc(string val)
        {
            switch ((val ?? "").Trim())
            {
                case "center":
                    return "center";
                case "right":
                case "end":
                    return "right";
                case "both":
                case "numTab":
                    return "justify";
                case "distribute":
                case "mediumKashida":
                case "highKashida":
                case "lowKashida":
                case "thaiDistribute":
                    return "distribute";
                case "left":
                case "start":
                    return "left";
                default:
                    return val;
            }
        }

        private static string MapLineRule(string val)
        {
            switch ((val ?? "").Trim())
            {
                case "auto":
                    return "single";
                case "exact":
                    return "exact";
                case "atLeast":
                    return "at_least";
                default:
                    return val;
            }
        }
    }
}
