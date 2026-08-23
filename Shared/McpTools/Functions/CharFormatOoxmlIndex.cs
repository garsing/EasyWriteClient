using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    /// <summary>用一次 WordOpenXML 按文本对齐抽取字符格式（替代 FormatExtractor COM Find）。</summary>
    public sealed class CharFormatOoxmlIndex
    {
        private static readonly XNamespace W = OpenXmlPackage.W;

        private readonly List<ParaRec> _paras = new List<ParaRec>();
        private readonly Dictionary<string, XElement> _styleRpr =
            new Dictionary<string, XElement>(StringComparer.Ordinal);
        private readonly XElement _docDefaultRpr;

        private sealed class ParaRec
        {
            public string Visible;
            public XElement Paragraph;
            public XElement FirstRun;
        }

        public static CharFormatOoxmlIndex Load(Word.Document document)
        {
            OpenXmlPackage pkg = OpenXmlPackage.Load(document);
            return new CharFormatOoxmlIndex(pkg);
        }

        private CharFormatOoxmlIndex(OpenXmlPackage pkg)
        {
            XElement styles = pkg.FindPartRoot("styles.xml", W + "styles");
            _docDefaultRpr = styles
                ?.Element(W + "docDefaults")
                ?.Element(W + "rPrDefault")
                ?.Element(W + "rPr");
            if (styles != null)
            {
                foreach (XElement style in styles.Elements(W + "style"))
                {
                    string id = (string)style.Attribute(W + "styleId");
                    if (string.IsNullOrEmpty(id))
                    {
                        continue;
                    }

                    XElement rPr = style.Element(W + "rPr");
                    if (rPr != null)
                    {
                        _styleRpr[id] = rPr;
                    }
                }
            }

            foreach (XElement p in pkg.Body.Descendants(W + "p"))
            {
                string visible = ParaFormatOoxmlReader.VisibleTextFromParagraph(p);
                XElement firstRun = p.Descendants(W + "r").FirstOrDefault(r =>
                    r.Descendants(W + "t").Any(t => !string.IsNullOrWhiteSpace((string)t)));
                _paras.Add(new ParaRec
                {
                    Visible = visible,
                    Paragraph = p,
                    FirstRun = firstRun
                });
            }
        }

        public List<Dictionary<string, object>> ExtractAllMatches(string sentenceText)
        {
            var list = new List<Dictionary<string, object>>();
            string target = ParaFormatOoxmlReader.NormalizeAlignText(sentenceText);
            if (string.IsNullOrEmpty(target))
            {
                return list;
            }

            foreach (ParaRec para in _paras)
            {
                if (string.IsNullOrEmpty(para.Visible))
                {
                    continue;
                }

                if (string.Equals(para.Visible, target, StringComparison.Ordinal)
                    || para.Visible.IndexOf(target, StringComparison.Ordinal) >= 0)
                {
                    list.Add(ExtractEffective(para));
                }
            }

            return list;
        }

        private Dictionary<string, object> ExtractEffective(ParaRec para)
        {
            var layers = new List<XElement>();
            if (_docDefaultRpr != null)
            {
                layers.Add(_docDefaultRpr);
            }

            string pStyle = (string)para.Paragraph
                ?.Element(W + "pPr")
                ?.Element(W + "pStyle")
                ?.Attribute(W + "val");
            if (!string.IsNullOrEmpty(pStyle) && _styleRpr.TryGetValue(pStyle, out XElement styleRpr))
            {
                layers.Add(styleRpr);
            }

            XElement run = para.FirstRun;
            if (run != null)
            {
                string rStyle = (string)run.Element(W + "rPr")?.Element(W + "rStyle")?.Attribute(W + "val");
                if (!string.IsNullOrEmpty(rStyle) && _styleRpr.TryGetValue(rStyle, out XElement crPr))
                {
                    layers.Add(crPr);
                }

                XElement runRpr = run.Element(W + "rPr");
                if (runRpr != null)
                {
                    layers.Add(runRpr);
                }
            }

            var format = new Dictionary<string, object>
            {
                ["font_name"] = "",
                ["font_size"] = 0f,
                ["font_color"] = "automatic",
                ["background_color"] = "automatic",
                ["is_bold"] = false,
                ["has_underline"] = false,
                ["has_strikethrough"] = false
            };

            foreach (XElement rPr in layers)
            {
                ApplyRpr(format, rPr);
            }

            return format;
        }

        private static void ApplyRpr(Dictionary<string, object> format, XElement rPr)
        {
            if (rPr == null)
            {
                return;
            }

            XElement fonts = rPr.Element(W + "rFonts");
            if (fonts != null)
            {
                string name = (string)fonts.Attribute(W + "eastAsia")
                    ?? (string)fonts.Attribute(W + "ascii")
                    ?? (string)fonts.Attribute(W + "hAnsi")
                    ?? (string)fonts.Attribute(W + "cs");
                if (!string.IsNullOrEmpty(name))
                {
                    format["font_name"] = name;
                }
            }

            string sz = (string)rPr.Element(W + "szCs")?.Attribute(W + "val")
                ?? (string)rPr.Element(W + "sz")?.Attribute(W + "val");
            if (int.TryParse(sz, out int half) && half > 0)
            {
                format["font_size"] = half / 2f;
            }

            XElement color = rPr.Element(W + "color");
            if (color != null)
            {
                string val = ((string)color.Attribute(W + "val") ?? "").Trim();
                if (!string.IsNullOrEmpty(val)
                    && !string.Equals(val, "auto", StringComparison.OrdinalIgnoreCase)
                    && !string.Equals(val, "automatic", StringComparison.OrdinalIgnoreCase))
                {
                    val = val.TrimStart('#');
                    format["font_color"] = val.Length >= 6 ? "#" + val.Substring(val.Length - 6).ToUpperInvariant() : "automatic";
                }
            }

            XElement shd = rPr.Element(W + "shd");
            if (shd != null)
            {
                string fill = ((string)shd.Attribute(W + "fill") ?? "").Trim();
                if (!string.IsNullOrEmpty(fill)
                    && !string.Equals(fill, "auto", StringComparison.OrdinalIgnoreCase)
                    && !string.Equals(fill, "clear", StringComparison.OrdinalIgnoreCase))
                {
                    fill = fill.TrimStart('#');
                    format["background_color"] = fill.Length >= 6
                        ? "#" + fill.Substring(fill.Length - 6).ToUpperInvariant()
                        : "automatic";
                }
            }

            XElement b = rPr.Element(W + "b");
            if (b != null)
            {
                format["is_bold"] = OnVal(b);
            }

            XElement strike = rPr.Element(W + "strike") ?? rPr.Element(W + "dstrike");
            if (strike != null)
            {
                format["has_strikethrough"] = OnVal(strike);
            }

            XElement u = rPr.Element(W + "u");
            if (u != null)
            {
                string uval = ((string)u.Attribute(W + "val") ?? "").ToLowerInvariant();
                format["has_underline"] = uval != "none" && uval != "0" && uval != "false";
            }
        }

        private static bool OnVal(XElement el)
        {
            string val = (string)el.Attribute(W + "val");
            if (string.IsNullOrEmpty(val))
            {
                return true;
            }

            return !string.Equals(val, "0", StringComparison.OrdinalIgnoreCase)
                && !string.Equals(val, "false", StringComparison.OrdinalIgnoreCase)
                && !string.Equals(val, "off", StringComparison.OrdinalIgnoreCase);
        }
    }
}
