using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Xml.Linq;

namespace WordAddIn1.ChartOoxmlHost
{
    /// <summary>宿主无关：chart 部件 c:ser / c:pt / 色快照。禁止引用 Word/PPT Interop。</summary>
    internal static class ChartPartXml
    {
        public static readonly XNamespace C = "http://schemas.openxmlformats.org/drawingml/2006/chart";
        public static readonly XNamespace A = "http://schemas.openxmlformats.org/drawingml/2006/main";

        public sealed class SeriesCol
        {
            public string Name { get; set; }
            public string AxisY { get; set; }
            public string Color { get; set; }
            public string ChartType { get; set; }
            public List<string> Categories { get; set; }
            public List<string> Values { get; set; }
        }

        public sealed class SeriesGrid
        {
            public string Type { get; set; }
            public string Title { get; set; }
            public List<SeriesCol> Series { get; set; }
        }

        public static bool TryParseGrid(XDocument chartDoc, out SeriesGrid grid, out string error)
        {
            grid = null;
            error = null;
            if (chartDoc == null || chartDoc.Root == null)
            {
                error = "chart 部件空";
                return false;
            }

            var series = new List<SeriesCol>();
            foreach (XElement ser in chartDoc.Descendants(C + "ser"))
            {
                series.Add(ParseSer(ser));
            }

            grid = new SeriesGrid
            {
                Type = DetectType(chartDoc),
                Title = ReadChartTitle(chartDoc),
                Series = series
            };
            return true;
        }

        public static string BuildDisplayChartXml(SeriesGrid grid)
        {
            if (grid == null || grid.Series == null || grid.Series.Count == 0)
            {
                return "<Chart>\n  <DataTable>\n  </DataTable>\n</Chart>\n";
            }

            var sb = new StringBuilder();
            sb.Append("<Chart>\n");
            sb.Append("  <DataTable>\n");
            sb.Append("    <Header>\n");
            sb.Append("      <Column type=\"category\">类别</Column>\n");
            foreach (SeriesCol col in grid.Series)
            {
                sb.Append("      <Column type=\"value\"");
                sb.Append(" axisY=\"").Append(EscapeXml(col.AxisY ?? ChartConfigSyntax.AxisYPrimary)).Append("\"");
                if (!string.IsNullOrEmpty(col.Color))
                {
                    sb.Append(" color=\"").Append(EscapeXml(col.Color)).Append("\"");
                }

                if (!string.IsNullOrEmpty(col.ChartType))
                {
                    sb.Append(" chartType=\"").Append(EscapeXml(col.ChartType)).Append("\"");
                }

                sb.Append(">").Append(EscapeXml(col.Name ?? "")).Append("</Column>\n");
            }

            sb.Append("    </Header>\n");
            sb.Append("    <Rows>\n");
            int rowCount = grid.Series[0].Values != null ? grid.Series[0].Values.Count : 0;
            if (grid.Series[0].Categories != null && grid.Series[0].Categories.Count > rowCount)
            {
                rowCount = grid.Series[0].Categories.Count;
            }

            for (int r = 0; r < rowCount; r++)
            {
                var cells = new List<string>();
                string x = "";
                if (grid.Series[0].Categories != null && r < grid.Series[0].Categories.Count)
                {
                    x = grid.Series[0].Categories[r] ?? "";
                }

                cells.Add(x);
                foreach (SeriesCol col in grid.Series)
                {
                    string raw = col.Values != null && r < col.Values.Count ? col.Values[r] : "";
                    if (double.TryParse(raw, NumberStyles.Any, CultureInfo.InvariantCulture, out double y))
                    {
                        cells.Add(y.ToString("F2", CultureInfo.InvariantCulture));
                    }
                    else
                    {
                        cells.Add(string.IsNullOrEmpty(raw) ? "0" : raw);
                    }
                }

                sb.Append("      <Row>").Append(EscapeXml(string.Join(",", cells))).Append("</Row>\n");
            }

            sb.Append("    </Rows>\n");
            sb.Append("  </DataTable>\n");
            sb.Append("</Chart>\n");
            return sb.ToString();
        }

        public static bool TryRewriteGrid(XDocument chartDoc, SeriesGrid grid, out string error)
        {
            error = null;
            if (chartDoc == null || grid == null || grid.Series == null || grid.Series.Count == 0)
            {
                error = "改包网格空";
                return false;
            }

            List<XElement> sers = chartDoc.Descendants(C + "ser").ToList();
            if (sers.Count == 0)
            {
                error = "chart 无 c:ser";
                return false;
            }

            XElement parent = sers[0].Parent;
            while (parent != null && sers.Count < grid.Series.Count)
            {
                XElement clone = new XElement(sers[0]);
                SetIdxOrder(clone, sers.Count);
                parent.Add(clone);
                sers.Add(clone);
            }

            while (sers.Count > grid.Series.Count)
            {
                int last = sers.Count - 1;
                sers[last].Remove();
                sers.RemoveAt(last);
            }

            for (int i = 0; i < grid.Series.Count && i < sers.Count; i++)
            {
                SeriesCol col = grid.Series[i];
                XElement ser = sers[i];
                SetIdxOrder(ser, i);
                RewriteTx(ser, col.Name ?? "");
                List<string> cats = col.Categories ?? new List<string>();
                List<string> vals = col.Values ?? new List<string>();
                RewritePts(ser, C + "cat", cats);
                RewritePts(ser, C + "val", vals);
            }

            return true;
        }

        public static List<string> SnapshotSeriesColors(XDocument chartDoc)
        {
            var colors = new List<string>();
            if (chartDoc == null)
            {
                return colors;
            }

            foreach (XElement ser in chartDoc.Descendants(C + "ser"))
            {
                colors.Add(FirstSrgb(ser) ?? "");
            }

            return colors;
        }

        public static bool ColorsMatch(IList<string> expected, IList<string> actual)
        {
            if (expected == null || actual == null || expected.Count != actual.Count)
            {
                return false;
            }

            for (int i = 0; i < expected.Count; i++)
            {
                if (!string.Equals(NormColor(expected[i]), NormColor(actual[i]), StringComparison.OrdinalIgnoreCase))
                {
                    return false;
                }
            }

            return true;
        }

        public static bool TryPinSeriesColors(XDocument chartDoc, IList<string> colors)
        {
            if (chartDoc == null || colors == null)
            {
                return false;
            }

            bool any = false;
            List<XElement> sers = chartDoc.Descendants(C + "ser").ToList();
            for (int i = 0; i < sers.Count && i < colors.Count; i++)
            {
                string want = NormColor(colors[i]);
                if (string.IsNullOrEmpty(want))
                {
                    continue;
                }

                if (PinSerColor(sers[i], want))
                {
                    any = true;
                }
            }

            return any;
        }

        public static string DetectType(XDocument chartDoc)
        {
            if (chartDoc == null)
            {
                return ChartConfigSyntax.TypeColumn;
            }

            if (chartDoc.Descendants(C + "pie3DChart").Any())
            {
                return ChartConfigSyntax.TypePie3d;
            }

            if (chartDoc.Descendants(C + "pieChart").Any())
            {
                return "pie2d";
            }

            if (chartDoc.Descendants(C + "lineChart").Any() || chartDoc.Descendants(C + "line3DChart").Any())
            {
                return ChartConfigSyntax.TypeLine;
            }

            XElement bar = chartDoc.Descendants(C + "barChart").FirstOrDefault()
                ?? chartDoc.Descendants(C + "bar3DChart").FirstOrDefault();
            if (bar != null)
            {
                XElement dir = bar.Element(C + "barDir");
                string val = dir?.Attribute("val")?.Value ?? "col";
                if (string.Equals(val, "bar", StringComparison.OrdinalIgnoreCase))
                {
                    return "bar";
                }

                return ChartConfigSyntax.TypeColumn;
            }

            return ChartConfigSyntax.TypeColumn;
        }

        public static string ReadChartTitle(XDocument chartDoc)
        {
            if (chartDoc == null)
            {
                return "";
            }

            XElement title = chartDoc.Descendants(C + "title").FirstOrDefault();
            if (title == null)
            {
                return "";
            }

            XElement v = title.Descendants(C + "v").FirstOrDefault();
            if (v != null)
            {
                return v.Value ?? "";
            }

            XElement t = title.Descendants(A + "t").FirstOrDefault();
            return t != null ? (t.Value ?? "") : "";
        }

        public static byte[] XmlPartBytes(XDocument doc)
        {
            var sb = new StringBuilder();
            sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n");
            sb.Append(doc.ToString(SaveOptions.DisableFormatting));
            return Encoding.UTF8.GetBytes(sb.ToString());
        }

        private static SeriesCol ParseSer(XElement ser)
        {
            var col = new SeriesCol
            {
                Name = ReadTx(ser),
                AxisY = ChartConfigSyntax.AxisYPrimary,
                Color = FirstSrgb(ser) ?? "",
                ChartType = "",
                Categories = ReadPts(ser.Element(C + "cat")),
                Values = ReadPts(ser.Element(C + "val"))
            };
            if (string.IsNullOrEmpty(col.Name))
            {
                col.Name = "系列1";
            }

            return col;
        }

        private static string ReadTx(XElement ser)
        {
            XElement tx = ser.Element(C + "tx");
            if (tx == null)
            {
                return "";
            }

            XElement v = tx.Element(C + "v");
            if (v != null)
            {
                return v.Value ?? "";
            }

            XElement pt = tx.Descendants(C + "pt").FirstOrDefault();
            XElement pv = pt != null ? pt.Element(C + "v") : null;
            return pv != null ? (pv.Value ?? "") : "";
        }

        private static List<string> ReadPts(XElement parent)
        {
            var list = new List<string>();
            if (parent == null)
            {
                return list;
            }

            foreach (XElement pt in parent.Descendants(C + "pt"))
            {
                XElement v = pt.Element(C + "v");
                list.Add(v != null ? (v.Value ?? "") : "");
            }

            return list;
        }

        private static void RewriteTx(XElement ser, string name)
        {
            XElement tx = ser.Element(C + "tx");
            if (tx == null)
            {
                tx = new XElement(C + "tx");
                ser.AddFirst(tx);
            }

            XElement v = tx.Element(C + "v");
            if (v != null)
            {
                v.Value = name ?? "";
                return;
            }

            XElement pt = tx.Descendants(C + "pt").FirstOrDefault();
            if (pt != null)
            {
                XElement pv = pt.Element(C + "v") ?? new XElement(C + "v");
                if (pv.Parent == null)
                {
                    pt.Add(pv);
                }

                pv.Value = name ?? "";
                return;
            }

            tx.RemoveNodes();
            tx.Add(new XElement(C + "v", name ?? ""));
        }

        private static void RewritePts(XElement ser, XName parentName, List<string> points)
        {
            XElement parent = ser.Element(parentName);
            if (parent == null)
            {
                return;
            }

            foreach (XElement cache in parent.Elements().ToList())
            {
                if (cache.Name != C + "strRef" && cache.Name != C + "numRef"
                    && cache.Name != C + "strLit" && cache.Name != C + "numLit")
                {
                    continue;
                }

                XElement inner = cache.Element(C + "strCache")
                    ?? cache.Element(C + "numCache")
                    ?? cache;
                WriteCachePts(inner, points);
            }
        }

        private static void WriteCachePts(XElement cache, List<string> points)
        {
            if (cache == null)
            {
                return;
            }

            XElement ptCount = cache.Element(C + "ptCount");
            if (ptCount != null)
            {
                ptCount.SetAttributeValue("val", points.Count.ToString(CultureInfo.InvariantCulture));
            }

            foreach (XElement pt in cache.Elements(C + "pt").ToList())
            {
                pt.Remove();
            }

            for (int i = 0; i < points.Count; i++)
            {
                var pt = new XElement(C + "pt");
                pt.SetAttributeValue("idx", i.ToString(CultureInfo.InvariantCulture));
                pt.Add(new XElement(C + "v", points[i] ?? ""));
                cache.Add(pt);
            }
        }

        private static void SetIdxOrder(XElement ser, int idx)
        {
            string text = idx.ToString(CultureInfo.InvariantCulture);
            XElement idxEl = ser.Element(C + "idx");
            if (idxEl != null)
            {
                idxEl.SetAttributeValue("val", text);
            }

            XElement orderEl = ser.Element(C + "order");
            if (orderEl != null)
            {
                orderEl.SetAttributeValue("val", text);
            }
        }

        private static string FirstSrgb(XElement root)
        {
            if (root == null)
            {
                return null;
            }

            foreach (XElement el in root.Descendants(A + "srgbClr"))
            {
                string val = el.Attribute("val")?.Value;
                if (!string.IsNullOrEmpty(val))
                {
                    return val;
                }
            }

            return null;
        }

        private static bool PinSerColor(XElement ser, string hex)
        {
            XElement spPr = ser.Element(C + "spPr");
            if (spPr == null)
            {
                spPr = new XElement(C + "spPr");
                ser.Add(spPr);
            }

            XElement srgb = spPr.Descendants(A + "srgbClr").FirstOrDefault();
            if (srgb != null)
            {
                srgb.SetAttributeValue("val", hex);
                foreach (XElement tint in srgb.Elements(A + "tint").ToList())
                {
                    tint.Remove();
                }

                return true;
            }

            XElement solid = new XElement(A + "solidFill");
            solid.Add(new XElement(A + "srgbClr", new XAttribute("val", hex)));
            XElement oldFill = spPr.Element(A + "solidFill");
            if (oldFill != null)
            {
                oldFill.ReplaceWith(solid);
            }
            else
            {
                spPr.AddFirst(solid);
            }

            return true;
        }

        private static string NormColor(string color)
        {
            if (string.IsNullOrEmpty(color))
            {
                return "";
            }

            string t = color.Trim();
            if (t.StartsWith("#", StringComparison.Ordinal))
            {
                t = t.Substring(1);
            }

            return t.ToUpperInvariant();
        }

        private static string EscapeXml(string value)
        {
            if (string.IsNullOrEmpty(value))
            {
                return "";
            }

            return value
                .Replace("&", "&amp;")
                .Replace("<", "&lt;")
                .Replace(">", "&gt;")
                .Replace("\"", "&quot;");
        }
    }
}
