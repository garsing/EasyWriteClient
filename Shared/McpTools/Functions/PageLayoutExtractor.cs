using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    /// <summary>
    /// I13：从 WordOpenXML 读各节 sectPr + 页眉脚部件。失败直接抛，不回退 COM。
    /// </summary>
    public static class PageLayoutExtractor
    {
        private static readonly XNamespace W = OpenXmlPackage.W;
        private static readonly XNamespace R = OpenXmlPackage.R;

        public static Dictionary<string, object> Extract(Word.Document document)
        {
            OpenXmlPackage pkg = OpenXmlPackage.Load(document);
            var warnings = new List<string>();
            List<XElement> sectPrs = CollectSectPr(pkg.Body);
            if (sectPrs.Count == 0)
            {
                throw new InvalidOperationException("WordOpenXML 中没有 w:sectPr");
            }

            Dictionary<string, string> rels = pkg.LoadDocumentRels();
            bool docOddEven = HasEvenAndOddHeaders(pkg);
            var sections = new List<Dictionary<string, object>>();
            for (int i = 0; i < sectPrs.Count; i++)
            {
                sections.Add(ExtractSection(pkg, sectPrs[i], i + 1, rels, docOddEven, warnings));
            }

            var payload = new Dictionary<string, object>
            {
                ["version"] = 1,
                ["section_count"] = sections.Count,
                ["sections"] = sections
            };

            if (sections.Count > 0 && sections[0].TryGetValue("page_setup", out object firstPs))
            {
                payload["mode"] = new Dictionary<string, object>
                {
                    ["page_setup"] = CloneDict(firstPs as Dictionary<string, object>),
                    ["headers"] = CloneDict(GetDict(sections[0], "headers")),
                    ["footers"] = CloneDict(GetDict(sections[0], "footers")),
                    ["mode_source_section_index"] = 1,
                    ["mode_occurrence_count"] = sections.Count
                };
            }

            if (warnings.Count > 0)
            {
                payload["extract_warnings"] = warnings;
            }

            return payload;
        }

        private static List<XElement> CollectSectPr(XElement body)
        {
            var list = new List<XElement>();
            foreach (XElement child in body.Elements())
            {
                if (child.Name == W + "sectPr")
                {
                    list.Add(child);
                    continue;
                }

                if (child.Name == W + "p")
                {
                    XElement nested = child.Element(W + "pPr")?.Element(W + "sectPr");
                    if (nested != null)
                    {
                        list.Add(nested);
                    }
                }
            }

            return list;
        }

        private static bool HasEvenAndOddHeaders(OpenXmlPackage pkg)
        {
            XElement settings = pkg.FindPartRoot("settings.xml", W + "settings");
            return settings?.Element(W + "evenAndOddHeaders") != null;
        }

        private static Dictionary<string, object> ExtractSection(
            OpenXmlPackage pkg,
            XElement sectPr,
            int index,
            Dictionary<string, string> rels,
            bool docOddEven,
            List<string> warnings)
        {
            XElement pgSz = sectPr.Element(W + "pgSz");
            XElement pgMar = sectPr.Element(W + "pgMar");
            string orient = ((string)pgSz?.Attribute(W + "orient") ?? "").Trim();
            var pageSetup = new Dictionary<string, object>
            {
                ["orientation"] = string.Equals(orient, "landscape", StringComparison.OrdinalIgnoreCase)
                    ? "landscape"
                    : "portrait",
                ["page_width"] = TwipAttr(pgSz, "w") ?? 595.3,
                ["page_height"] = TwipAttr(pgSz, "h") ?? 841.9,
                ["top_margin"] = TwipAttr(pgMar, "top") ?? 72.0,
                ["bottom_margin"] = TwipAttr(pgMar, "bottom") ?? 72.0,
                ["left_margin"] = TwipAttr(pgMar, "left") ?? 90.0,
                ["right_margin"] = TwipAttr(pgMar, "right") ?? 90.0,
                ["header_distance"] = TwipAttr(pgMar, "header") ?? 42.5,
                ["footer_distance"] = TwipAttr(pgMar, "footer") ?? 49.6,
                ["different_first_page_header_footer"] = sectPr.Element(W + "titlePg") != null,
                ["odd_and_even_pages_header_footer"] =
                    docOddEven || sectPr.Element(W + "evenAndOddHeaders") != null
            };

            int pages = 1;
            return new Dictionary<string, object>
            {
                ["section_index"] = index,
                ["pages_in_section"] = pages,
                ["extract_origin"] = "ooxml",
                ["page_setup"] = pageSetup,
                ["headers"] = ExtractStories(pkg, sectPr, rels, header: true, warnings, index),
                ["footers"] = ExtractStories(pkg, sectPr, rels, header: false, warnings, index)
            };
        }

        private static Dictionary<string, object> ExtractStories(
            OpenXmlPackage pkg,
            XElement sectPr,
            Dictionary<string, string> rels,
            bool header,
            List<string> warnings,
            int sectionIndex)
        {
            var stories = new Dictionary<string, object>(StringComparer.Ordinal);
            string tag = header ? "headerReference" : "footerReference";
            foreach (XElement href in sectPr.Elements(W + tag))
            {
                string type = ((string)href.Attribute(W + "type") ?? "default").Trim();
                string key = type == "first" ? "first" : type == "even" ? "even" : "primary";
                string rid = (string)href.Attribute(R + "id") ?? "";
                if (string.IsNullOrEmpty(rid) || !rels.TryGetValue(rid, out string target))
                {
                    warnings.Add("section " + sectionIndex + " " + key + " 缺少 rel");
                    continue;
                }

                Dictionary<string, object> story = StoryFromPart(pkg, target);
                if (story != null)
                {
                    stories[key] = story;
                }
            }

            return stories.Count == 0 ? null : stories;
        }

        private static Dictionary<string, object> StoryFromPart(OpenXmlPackage pkg, string target)
        {
            XElement part = pkg.FindWordPartByTarget(target);
            if (part == null)
            {
                return null;
            }

            XElement root = part.Name == W + "hdr" || part.Name == W + "ftr"
                ? part
                : part.Descendants().FirstOrDefault(e => e.Name == W + "hdr" || e.Name == W + "ftr");
            if (root == null)
            {
                return null;
            }

            var sb = new System.Text.StringBuilder();
            foreach (XElement t in root.Descendants(W + "t"))
            {
                sb.Append((string)t);
            }

            string text = sb.ToString().Replace("\r", "").Trim();
            var story = new Dictionary<string, object>();
            if (!string.IsNullOrEmpty(text))
            {
                story["text"] = text;
            }

            XElement firstRun = root.Descendants(W + "r").FirstOrDefault(r =>
                r.Descendants(W + "t").Any(x => !string.IsNullOrWhiteSpace((string)x)));
            if (firstRun != null)
            {
                var charFormat = new Dictionary<string, object>();
                XElement fonts = firstRun.Element(W + "rPr")?.Element(W + "rFonts");
                string name = (string)fonts?.Attribute(W + "eastAsia")
                    ?? (string)fonts?.Attribute(W + "ascii")
                    ?? (string)fonts?.Attribute(W + "hAnsi");
                if (!string.IsNullOrEmpty(name))
                {
                    charFormat["font_name"] = name;
                }

                string sz = (string)firstRun.Element(W + "rPr")?.Element(W + "sz")?.Attribute(W + "val");
                if (int.TryParse(sz, out int half) && half > 0)
                {
                    charFormat["font_size"] = half / 2.0;
                }

                if (charFormat.Count > 0)
                {
                    story["char_format"] = charFormat;
                }
            }

            return story.Count == 0 ? null : story;
        }

        private static double? TwipAttr(XElement el, string localName)
        {
            if (el == null)
            {
                return null;
            }

            string raw = (string)el.Attribute(W + localName);
            if (string.IsNullOrEmpty(raw) || !int.TryParse(raw, out int twip))
            {
                return null;
            }

            return Math.Round(twip / 20.0, 2);
        }

        private static Dictionary<string, object> GetDict(Dictionary<string, object> row, string key)
        {
            if (row != null && row.TryGetValue(key, out object raw) && raw is Dictionary<string, object> dict)
            {
                return dict;
            }

            return null;
        }

        private static Dictionary<string, object> CloneDict(Dictionary<string, object> src)
        {
            if (src == null)
            {
                return null;
            }

            return new Dictionary<string, object>(src, StringComparer.Ordinal);
        }
    }
}
