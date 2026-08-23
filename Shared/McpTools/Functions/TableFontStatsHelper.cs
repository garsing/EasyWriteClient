using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    public sealed class TableFontStatsResult
    {
        public string TableFontName { get; set; } = "";
        public float TableFontSize { get; set; }
        public string TableFontColor { get; set; } = "000000";
        public List<string> Warnings { get; set; } = new List<string>();
    }

    /// <summary>
    /// 表内 run 字体众数统计，语义对齐 Backend table_font_stats.py。
    /// </summary>
    public static class TableFontStatsHelper
    {
        private static readonly XNamespace W =
            "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

        public static TableFontStatsResult ComputeModeTripletFromTableElement(XElement tableElement)
        {
            var nameCounts = new Dictionary<string, int>(StringComparer.Ordinal);
            var sizeCounts = new Dictionary<float, int>();
            var colorCounts = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            if (tableElement != null)
            {
                AccumulateRunStats(tableElement, nameCounts, sizeCounts, colorCounts);
            }

            return FinishModeTriplet(nameCounts, sizeCounts, colorCounts, "", 0f, "000000");
        }

        public static TableFontStatsResult ComputeModeTriplet(Word.Table table, Word.Document document)
        {
            var nameCounts = new Dictionary<string, int>(StringComparer.Ordinal);
            var sizeCounts = new Dictionary<float, int>();
            var colorCounts = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

            try
            {
                string tableXml = table?.Range?.WordOpenXML;
                if (!string.IsNullOrEmpty(tableXml))
                {
                    XDocument xmlDoc = XDocument.Parse(tableXml);
                    AccumulateRunStats(xmlDoc.Root, nameCounts, sizeCounts, colorCounts);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[TableFontStatsHelper] 统计失败：{ex.Message}");
            }

            GetBodyFormat(document, out string bodyName, out float bodySize, out string bodyColor);
            return FinishModeTriplet(nameCounts, sizeCounts, colorCounts, bodyName, bodySize, bodyColor);
        }

        private static void AccumulateRunStats(
            XElement root,
            Dictionary<string, int> nameCounts,
            Dictionary<float, int> sizeCounts,
            Dictionary<string, int> colorCounts)
        {
            if (root == null)
            {
                return;
            }

            foreach (XElement run in root.Descendants(W + "r"))
            {
                string text = string.Join("", run.Descendants(W + "t").Select(t => (string)t));
                if (string.IsNullOrWhiteSpace(text))
                {
                    continue;
                }

                XElement rPr = run.Element(W + "rPr");
                string name = ExtractRunFontName(rPr);
                if (!string.IsNullOrWhiteSpace(name))
                {
                    if (!nameCounts.ContainsKey(name))
                    {
                        nameCounts[name] = 0;
                    }

                    nameCounts[name]++;
                }

                float size = ExtractRunFontSize(rPr);
                if (size > 0)
                {
                    if (!sizeCounts.ContainsKey(size))
                    {
                        sizeCounts[size] = 0;
                    }

                    sizeCounts[size]++;
                }

                string colorKey = ExtractRunColorKey(rPr);
                if (!colorCounts.ContainsKey(colorKey))
                {
                    colorCounts[colorKey] = 0;
                }

                colorCounts[colorKey]++;
            }
        }

        private static TableFontStatsResult FinishModeTriplet(
            Dictionary<string, int> nameCounts,
            Dictionary<float, int> sizeCounts,
            Dictionary<string, int> colorCounts,
            string bodyName,
            float bodySize,
            string bodyColor)
        {
            var result = new TableFontStatsResult();
            var warnings = new List<string>();

            string nameMode = nameCounts.Count > 0
                ? nameCounts.OrderByDescending(kvp => kvp.Value).First().Key
                : "";
            if (string.IsNullOrWhiteSpace(nameMode))
            {
                nameMode = bodyName ?? "";
                if (!string.IsNullOrWhiteSpace(bodyName))
                {
                    warnings.Add("table_font_name_from_body_fallback");
                }
            }

            float sizeMode;
            if (sizeCounts.Count > 0)
            {
                sizeMode = sizeCounts.OrderByDescending(kvp => kvp.Value).First().Key;
            }
            else
            {
                sizeMode = bodySize;
                warnings.Add("table_font_size_from_body_fallback");
            }

            var explicitColors = colorCounts
                .Where(kvp => !string.Equals(kvp.Key, "auto", StringComparison.OrdinalIgnoreCase))
                .ToList();
            string colorMode;
            if (explicitColors.Count > 0)
            {
                colorMode = explicitColors.OrderByDescending(kvp => kvp.Value).First().Key;
            }
            else
            {
                colorMode = bodyColor;
            }

            result.TableFontName = nameMode ?? "";
            result.TableFontSize = sizeMode;
            result.TableFontColor = NormalizeFontColor(colorMode, bodyColor);
            result.Warnings = warnings;
            return result;
        }

        private static void GetBodyFormat(
            Word.Document document,
            out string fontName,
            out float fontSize,
            out string fontColor)
        {
            fontName = "";
            fontSize = 10.5f;
            fontColor = "000000";

            try
            {
                foreach (var entry in DocumentState.SubtypeFormatMapping)
                {
                    if (entry == null)
                    {
                        continue;
                    }

                    if (!entry.TryGetValue("detailed_subtype", out object subtypeObj))
                    {
                        continue;
                    }

                    if (!string.Equals(subtypeObj?.ToString(), "body", StringComparison.OrdinalIgnoreCase))
                    {
                        continue;
                    }

                    if (entry.TryGetValue("format", out object formatObj)
                        && formatObj is Dictionary<string, object> format)
                    {
                        fontName = format.ContainsKey("font_name")
                            ? format["font_name"]?.ToString() ?? ""
                            : "";
                        if (format.ContainsKey("font_size"))
                        {
                            object sizeObj = format["font_size"];
                            if (sizeObj is float f)
                            {
                                fontSize = f;
                            }
                            else if (float.TryParse(sizeObj?.ToString(), out float parsed))
                            {
                                fontSize = parsed;
                            }
                        }

                        fontColor = NormalizeFontColor(
                            format.ContainsKey("font_color") ? format["font_color"]?.ToString() : null,
                            "000000");
                        return;
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[TableFontStatsHelper] mapping body 回退失败：{ex.Message}");
            }

            try
            {
                if (document != null)
                {
                    Word.Style normal = document.Styles[Word.WdBuiltinStyle.wdStyleNormal];
                    fontName = normal?.Font?.Name ?? "";
                    if (normal?.Font?.Size > 0)
                    {
                        fontSize = (float)normal.Font.Size;
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[TableFontStatsHelper] Normal 样式回退失败：{ex.Message}");
            }
        }

        private static string ExtractRunFontName(XElement rPr)
        {
            if (rPr == null)
            {
                return "";
            }

            var rFonts = rPr.Element(W + "rFonts");
            if (rFonts == null)
            {
                return "";
            }

            foreach (string attr in new[] { "ascii", "eastAsia", "hAnsi", "cs" })
            {
                string val = GetValAttribute(rFonts, attr);
                if (!string.IsNullOrWhiteSpace(val))
                {
                    return val.Trim();
                }
            }

            return "";
        }

        private static float ExtractRunFontSize(XElement rPr)
        {
            if (rPr == null)
            {
                return 0f;
            }

            var sz = rPr.Element(W + "sz");
            string val = GetValAttribute(sz, "val");
            if (int.TryParse(val, out int halfPoints) && halfPoints > 0)
            {
                return halfPoints / 2f;
            }

            return 0f;
        }

        private static string ExtractRunColorKey(XElement rPr)
        {
            if (rPr == null)
            {
                return "auto";
            }

            var colorEl = rPr.Element(W + "color");
            string val = GetValAttribute(colorEl, "val");
            if (string.IsNullOrWhiteSpace(val)
                || string.Equals(val, "auto", StringComparison.OrdinalIgnoreCase))
            {
                return "auto";
            }

            return TableFormatExtractor.ConvertBgrToRgb(val.ToUpperInvariant());
        }

        private static string GetValAttribute(XElement element, string localName)
        {
            if (element == null)
            {
                return null;
            }

            var attr = element.Attribute(W + localName) ?? element.Attribute("val");
            return attr?.Value;
        }

        private static string NormalizeFontColor(string color, string fallback = "000000")
        {
            string c = (color ?? "").Trim().TrimStart('#');
            if (string.IsNullOrEmpty(c)
                || string.Equals(c, "auto", StringComparison.OrdinalIgnoreCase)
                || string.Equals(c, "automatic", StringComparison.OrdinalIgnoreCase))
            {
                string fb = (fallback ?? "000000").Trim().TrimStart('#');
                return fb.Length >= 6 ? fb.Substring(fb.Length - 6).ToUpperInvariant() : "000000";
            }

            if (c.Length >= 6)
            {
                return c.Substring(c.Length - 6).ToUpperInvariant();
            }

            return "000000";
        }
    }
}
