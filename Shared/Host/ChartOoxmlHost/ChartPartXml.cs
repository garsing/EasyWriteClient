using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Xml.Linq;

namespace WordAddIn1.ChartOoxmlHost
{
    /// <summary>宿主无关：chart 部件 c:ser / c:pt / 色快照。禁止引用 Word/PPT Interop。</summary>
    internal static partial class ChartPartXml
    {
        public static readonly XNamespace C = "http://schemas.openxmlformats.org/drawingml/2006/chart";
        public static readonly XNamespace A = "http://schemas.openxmlformats.org/drawingml/2006/main";

        public sealed class SeriesCol
        {
            public string Name { get; set; }
            public string AxisY { get; set; }
            public string Color { get; set; }
            public string ChartType { get; set; }
            public string Line { get; set; }
            public string LineWeight { get; set; }
            public string Marker { get; set; }
            public string MarkerSize { get; set; }
            public string MarkerColor { get; set; }
            public string MarkerFill { get; set; }
            public bool? ShowDataLabels { get; set; }
            public bool? ShowValue { get; set; }
            public bool? ShowPercentage { get; set; }
            public string LabelPosition { get; set; }
            public string LabelFont { get; set; }
            public string LabelSize { get; set; }
            public string LabelColor { get; set; }
            public string LabelFormat { get; set; }
            public string FillGradient { get; set; }
            public string FillAngle { get; set; }
            public List<string> Categories { get; set; }
            public List<string> Values { get; set; }
        }

        public sealed class SeriesGrid
        {
            public string Type { get; set; }
            public string Title { get; set; }
            /// <summary>稿里写了 Title 且正文为空。换数时删掉 c:title；没写则留下旧标题。</summary>
            public bool ClearTitle { get; set; }
            public string Legend { get; set; }
            public string LegendFontColor { get; set; }
            public string AxisYTitle { get; set; }
            public string AxisY2Title { get; set; }
            public string AxisXTitle { get; set; }
            public AxisXInfo AxisX { get; set; }
            public string Gridlines { get; set; }
            public int ChartStyle { get; set; }
            public AxisYInfo AxisY { get; set; }
            public string PlotColor { get; set; }
            public string ChartAreaColor { get; set; }
            public string TitleFontColor { get; set; }
            public int TitleFontSize { get; set; }
            public bool TitleFontBold { get; set; }
            public int? Explosion { get; set; }
            public double? YMin { get; set; }
            public double? YMax { get; set; }
            public double? YMajor { get; set; }
            public int? GapWidth { get; set; }
            public int? Overlap { get; set; }
            public AxisYInfo AxisY2 { get; set; }
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
                TitleFontColor = ReadTitleFontColor(chartDoc),
                TitleFontSize = ReadTitleFontSize(chartDoc),
                TitleFontBold = ReadTitleFontBold(chartDoc),
                Explosion = ReadExplosion(chartDoc),
                Series = series,
                Legend = ReadLegend(chartDoc),
                LegendFontColor = ReadLegendFontColor(chartDoc),
                AxisYTitle = ReadValAxisTitle(chartDoc, secondary: false),
                AxisY2Title = ReadValAxisTitle(chartDoc, secondary: true),
                AxisY = ReadAxisChrome(FindValAx(chartDoc, false)),
                AxisXTitle = ReadCatAxisTitle(chartDoc),
                AxisX = ReadCatAxis(chartDoc),
                Gridlines = ReadGridlines(chartDoc),
                ChartStyle = ReadChartStyle(chartDoc),
                AxisY2 = ReadValAxisChrome(chartDoc),
                PlotColor = ReadAreaFill(chartDoc.Descendants(C + "plotArea").FirstOrDefault()),
                ChartAreaColor = ReadAreaFill(chartDoc.Root),
                GapWidth = ReadIntVal(chartDoc.Descendants(C + "barChart").FirstOrDefault()?.Element(C + "gapWidth")),
                Overlap = ReadIntVal(chartDoc.Descendants(C + "barChart").FirstOrDefault()?.Element(C + "overlap"))
            };
            ReadPrimaryScale(chartDoc, grid);
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
                WriteSeriesChrome(ser, col);
            }

            PlaceSecondarySeries(chartDoc, grid);
            WriteChartChrome(chartDoc, grid);

            if (!string.IsNullOrEmpty(grid.Title))
            {
                UpsertTitle(chartDoc, grid);
            }
            else if (grid.ClearTitle)
            {
                RemoveTitle(chartDoc);
            }

            if (grid.Explosion.HasValue)
            {
                WriteExplosion(chartDoc, grid.Explosion.Value);
            }

            return true;
        }

        /// <summary>只写类型以外的皮和标题。不改 c:f，不改点缓存。</summary>
        public static bool TryApplyChrome(XDocument chartDoc, SeriesGrid grid, out string error)
        {
            error = null;
            if (chartDoc == null || grid == null)
            {
                error = "改包网格空";
                return false;
            }

            WriteChartChrome(chartDoc, grid);
            if (grid.Series != null)
            {
                List<XElement> sers = chartDoc.Descendants(C + "ser").ToList();
                for (int i = 0; i < grid.Series.Count && i < sers.Count; i++)
                {
                    WriteSeriesChrome(sers[i], grid.Series[i]);
                }

                PlaceSecondarySeries(chartDoc, grid);
            }

            if (!string.IsNullOrEmpty(grid.Title))
            {
                UpsertTitle(chartDoc, grid);
            }
            else if (grid.ClearTitle)
            {
                RemoveTitle(chartDoc);
            }

            if (grid.Explosion.HasValue)
            {
                WriteExplosion(chartDoc, grid.Explosion.Value);
            }

            return true;
        }

        /// <summary>
        /// 次轴系列从柱图组挪到折线图组，并挂到右侧数值轴。WPS 抽图只读包，不认 COM 的 AxisGroup。
        /// </summary>
        private static void PlaceSecondarySeries(XDocument chartDoc, SeriesGrid grid)
        {
            if (chartDoc == null || grid == null || grid.Series == null)
            {
                return;
            }

            bool any = false;
            for (int i = 0; i < grid.Series.Count; i++)
            {
                if (ChartConfigSyntax.IsSecondaryAxis(grid.Series[i].AxisY))
                {
                    any = true;
                    break;
                }
            }

            if (!any)
            {
                return;
            }

            XElement plot = chartDoc.Descendants(C + "plotArea").FirstOrDefault();
            if (plot == null)
            {
                return;
            }

            string catId = AxisId(plot, "catAx");
            if (string.IsNullOrEmpty(catId))
            {
                catId = "1";
            }

            string secId = EnsureRightValueAxis(plot, catId);
            var valueAxisIds = new HashSet<string>(StringComparer.Ordinal);
            foreach (XElement val in plot.Elements(C + "valAx"))
            {
                string id = val.Element(C + "axId")?.Attribute("val")?.Value;
                if (!string.IsNullOrEmpty(id))
                {
                    valueAxisIds.Add(id);
                }
            }

            List<XElement> sers = chartDoc.Descendants(C + "ser").ToList();
            var secondary = new HashSet<XElement>();
            for (int i = 0; i < grid.Series.Count && i < sers.Count; i++)
            {
                if (ChartConfigSyntax.IsSecondaryAxis(grid.Series[i].AxisY))
                {
                    secondary.Add(sers[i]);
                }
            }

            XElement lineChart = null;
            foreach (XElement line in plot.Elements(C + "lineChart"))
            {
                if (string.Equals(ChartValueAxisId(line, valueAxisIds), secId, StringComparison.Ordinal))
                {
                    lineChart = line;
                    break;
                }
            }

            if (lineChart == null)
            {
                foreach (XElement line in plot.Elements(C + "lineChart"))
                {
                    List<XElement> kids = line.Elements(C + "ser").ToList();
                    if (kids.Count == 0 || !kids.All(secondary.Contains))
                    {
                        continue;
                    }

                    PointChartAtValueAxis(line, valueAxisIds, secId);
                    lineChart = line;
                    break;
                }
            }

            if (lineChart == null)
            {
                lineChart = new XElement(
                    C + "lineChart",
                    new XElement(C + "grouping", new XAttribute("val", "standard")),
                    new XElement(C + "axId", new XAttribute("val", catId)),
                    new XElement(C + "axId", new XAttribute("val", secId)));
                XElement bar = plot.Elements(C + "barChart").FirstOrDefault()
                    ?? plot.Elements(C + "bar3DChart").FirstOrDefault();
                if (bar != null)
                {
                    bar.AddAfterSelf(lineChart);
                }
                else
                {
                    plot.AddFirst(lineChart);
                }
            }

            foreach (XElement ser in secondary)
            {
                if (ser.Parent == lineChart)
                {
                    continue;
                }

                ser.Remove();
                XElement firstAx = lineChart.Elements(C + "axId").FirstOrDefault();
                if (firstAx != null)
                {
                    firstAx.AddBeforeSelf(ser);
                }
                else
                {
                    lineChart.Add(ser);
                }
            }

            foreach (XElement line in plot.Elements(C + "lineChart").ToList())
            {
                if (!line.Elements(C + "ser").Any())
                {
                    line.Remove();
                }
            }
        }

        private static string ChartValueAxisId(XElement chartEl, HashSet<string> valueAxisIds)
        {
            if (chartEl == null)
            {
                return null;
            }

            foreach (XElement ax in chartEl.Elements(C + "axId"))
            {
                string id = ax.Attribute("val")?.Value;
                if (!string.IsNullOrEmpty(id) && valueAxisIds.Contains(id))
                {
                    return id;
                }
            }

            return null;
        }

        private static void PointChartAtValueAxis(XElement chartEl, HashSet<string> valueAxisIds, string valueAxisId)
        {
            foreach (XElement ax in chartEl.Elements(C + "axId"))
            {
                string id = ax.Attribute("val")?.Value;
                if (!string.IsNullOrEmpty(id) && valueAxisIds.Contains(id))
                {
                    ax.SetAttributeValue("val", valueAxisId);
                    return;
                }
            }

            chartEl.Add(new XElement(C + "axId", new XAttribute("val", valueAxisId)));
        }

        private static string EnsureRightValueAxis(XElement plot, string catId)
        {
            foreach (XElement val in plot.Elements(C + "valAx"))
            {
                string pos = val.Element(C + "axPos")?.Attribute("val")?.Value;
                if (string.Equals(pos, "r", StringComparison.OrdinalIgnoreCase))
                {
                    XElement id = val.Element(C + "axId");
                    string existing = id?.Attribute("val")?.Value;
                    if (!string.IsNullOrEmpty(existing))
                    {
                        return existing;
                    }
                }
            }

            int next = 3;
            foreach (XElement idEl in plot.Descendants(C + "axId"))
            {
                int n;
                if (int.TryParse(idEl.Attribute("val")?.Value, out n) && n >= next)
                {
                    next = n + 1;
                }
            }

            string secId = next.ToString(CultureInfo.InvariantCulture);
            XElement primary = null;
            foreach (XElement val in plot.Elements(C + "valAx"))
            {
                string pos = val.Element(C + "axPos")?.Attribute("val")?.Value;
                if (!string.Equals(pos, "r", StringComparison.OrdinalIgnoreCase))
                {
                    primary = val;
                    break;
                }
            }

            var sec = new XElement(
                C + "valAx",
                new XElement(C + "axId", new XAttribute("val", secId)),
                new XElement(
                    C + "scaling",
                    new XElement(C + "orientation", new XAttribute("val", "minMax"))),
                new XElement(C + "delete", new XAttribute("val", "0")),
                new XElement(C + "axPos", new XAttribute("val", "r")));
            CopyAxisChild(primary, sec, "numFmt");
            CopyAxisChild(primary, sec, "majorTickMark");
            CopyAxisChild(primary, sec, "minorTickMark");
            CopyAxisChild(primary, sec, "tickLblPos");
            sec.Add(new XElement(C + "crossAx", new XAttribute("val", catId)));
            sec.Add(new XElement(C + "crosses", new XAttribute("val", "max")));
            CopyAxisChild(primary, sec, "crossBetween");
            plot.Add(sec);
            return secId;
        }

        private static void CopyAxisChild(XElement from, XElement to, string localName)
        {
            XElement child = from?.Element(C + localName);
            if (child != null)
            {
                to.Add(new XElement(child));
            }
        }

        private static string AxisId(XElement plot, string axisName)
        {
            XElement axis = plot.Elements(C + axisName).FirstOrDefault();
            return axis?.Element(C + "axId")?.Attribute("val")?.Value;
        }

        private static bool SeriesOnSecondary(XElement ser)
        {
            XElement chartEl = ser?.Parent;
            XElement plot = chartEl?.Parent;
            if (chartEl == null || plot == null)
            {
                return false;
            }

            var ids = new HashSet<string>(StringComparer.Ordinal);
            foreach (XElement ax in chartEl.Elements(C + "axId"))
            {
                string id = ax.Attribute("val")?.Value;
                if (!string.IsNullOrEmpty(id))
                {
                    ids.Add(id);
                }
            }

            foreach (XElement val in plot.Elements(C + "valAx"))
            {
                string id = val.Element(C + "axId")?.Attribute("val")?.Value;
                if (string.IsNullOrEmpty(id) || !ids.Contains(id))
                {
                    continue;
                }

                string pos = val.Element(C + "axPos")?.Attribute("val")?.Value;
                if (string.Equals(pos, "r", StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
            }

            return false;
        }

        /// <summary>稿明确关掉标题：去掉 c:title，并标 autoTitleDeleted，避免宿主把系列名补回标题。</summary>
        private static void RemoveTitle(XDocument chartDoc)
        {
            XElement chart = chartDoc.Descendants(C + "chart").FirstOrDefault();
            if (chart == null)
            {
                return;
            }

            foreach (XElement old in chart.Elements(C + "title").ToList())
            {
                old.Remove();
            }

            XElement auto = chart.Element(C + "autoTitleDeleted");
            if (auto == null)
            {
                var deleted = new XElement(C + "autoTitleDeleted", new XAttribute("val", "1"));
                XElement plot = chart.Element(C + "plotArea");
                if (plot != null)
                {
                    plot.AddBeforeSelf(deleted);
                }
                else
                {
                    chart.AddFirst(deleted);
                }
            }
            else
            {
                auto.SetAttributeValue("val", "1");
            }
        }

        /// <summary>有标题才写 c:title。没写标题时不进这里，原节点留着。</summary>
        private static void UpsertTitle(XDocument chartDoc, SeriesGrid grid)
        {
            XElement chart = chartDoc.Descendants(C + "chart").FirstOrDefault();
            if (chart == null)
            {
                return;
            }

            foreach (XElement old in chart.Elements(C + "title").ToList())
            {
                old.Remove();
            }

            var bodyPr = new XElement(
                A + "bodyPr",
                new XAttribute("rot", "0"),
                new XAttribute("spcFirstLastPara", "0"),
                new XAttribute("vertOverflow", "ellipsis"),
                new XAttribute("vert", "horz"),
                new XAttribute("wrap", "square"),
                new XAttribute("anchor", "ctr"),
                new XAttribute("anchorCtr", "1"));
            var rPr = new XElement(A + "rPr", new XAttribute("lang", "zh-CN"));
            if (grid.TitleFontSize > 0)
            {
                rPr.SetAttributeValue("sz", ((int)Math.Round(grid.TitleFontSize * 100.0)).ToString(CultureInfo.InvariantCulture));
            }

            if (grid.TitleFontBold)
            {
                rPr.SetAttributeValue("b", "1");
            }

            string titleColor = NormColor(grid.TitleFontColor);
            if (!string.IsNullOrEmpty(titleColor))
            {
                rPr.Add(new XElement(A + "solidFill", new XElement(A + "srgbClr", new XAttribute("val", titleColor))));
            }

            var titleEl = new XElement(
                C + "title",
                new XElement(
                    C + "tx",
                    new XElement(
                        C + "rich",
                        new XElement(bodyPr),
                        new XElement(A + "lstStyle"),
                        new XElement(
                            A + "p",
                            new XElement(
                                A + "r",
                                rPr,
                                new XElement(A + "t", grid.Title ?? ""))))),
                new XElement(C + "overlay", new XAttribute("val", "0")));
            XElement plot = chart.Element(C + "plotArea");
            if (plot != null)
            {
                plot.AddBeforeSelf(titleEl);
            }
            else
            {
                chart.AddFirst(titleEl);
            }

            XElement auto = chart.Element(C + "autoTitleDeleted");
            if (auto == null)
            {
                titleEl.AddAfterSelf(new XElement(C + "autoTitleDeleted", new XAttribute("val", "0")));
            }
            else
            {
                auto.SetAttributeValue("val", "0");
            }
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
                colors.Add(SnapshotFill(ser));
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
                if (IsNone(colors[i]))
                {
                    WriteSeriesNoFill(sers[i]);
                    any = true;
                    continue;
                }

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

            bool hasLine = chartDoc.Descendants(C + "lineChart").Any()
                || chartDoc.Descendants(C + "line3DChart").Any();
            XElement bar = chartDoc.Descendants(C + "barChart").FirstOrDefault()
                ?? chartDoc.Descendants(C + "bar3DChart").FirstOrDefault();
            if (hasLine && bar == null)
            {
                return ChartConfigSyntax.TypeLine;
            }

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

        private static XElement TitleRunProps(XDocument chartDoc)
        {
            return chartDoc?.Descendants(C + "title").FirstOrDefault()?.Descendants(A + "rPr").FirstOrDefault();
        }

        private static string ReadTitleFontColor(XDocument chartDoc)
        {
            return TitleRunProps(chartDoc)?.Element(A + "solidFill")?.Element(A + "srgbClr")?.Attribute("val")?.Value;
        }

        private static int ReadTitleFontSize(XDocument chartDoc)
        {
            int hundredths;
            string sz = TitleRunProps(chartDoc)?.Attribute("sz")?.Value;
            if (!int.TryParse(sz, NumberStyles.Integer, CultureInfo.InvariantCulture, out hundredths) || hundredths <= 0)
            {
                return 0;
            }

            return (int)Math.Round(hundredths / 100.0);
        }

        private static bool ReadTitleFontBold(XDocument chartDoc)
        {
            string bold = TitleRunProps(chartDoc)?.Attribute("b")?.Value;
            return bold == "1" || string.Equals(bold, "true", StringComparison.OrdinalIgnoreCase);
        }

        private static int? ReadExplosion(XDocument chartDoc)
        {
            string val = chartDoc?.Descendants(C + "explosion").FirstOrDefault()?.Attribute("val")?.Value;
            int n;
            return int.TryParse(val, NumberStyles.Integer, CultureInfo.InvariantCulture, out n) ? (int?)n : null;
        }

        private static void WriteExplosion(XDocument chartDoc, int explosion)
        {
            XElement ser = chartDoc?.Descendants(C + "ser").FirstOrDefault();
            if (ser == null)
            {
                return;
            }

            SetVal(ser, "explosion", explosion.ToString(CultureInfo.InvariantCulture));
        }

        public static byte[] XmlPartBytes(XDocument doc)
        {
            var sb = new StringBuilder();
            sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n");
            sb.Append(doc.ToString(SaveOptions.DisableFormatting));
            return Encoding.UTF8.GetBytes(sb.ToString());
        }

        private static string SeriesChartType(XElement ser)
        {
            string parent = ser == null || ser.Parent == null ? "" : ser.Parent.Name.LocalName;
            if (parent == "lineChart")
            {
                return "line";
            }

            if (parent == "pieChart")
            {
                return "pie2d";
            }

            if (parent == "pie3DChart")
            {
                return "pie3d";
            }

            if (parent == "barChart")
            {
                string dir = ser.Parent.Element(C + "barDir")?.Attribute("val")?.Value;
                return string.Equals(dir, "bar", StringComparison.OrdinalIgnoreCase) ? "bar" : "column";
            }

            return "";
        }

        private static SeriesCol ParseSer(XElement ser)
        {
            var col = new SeriesCol
            {
                Name = ReadTx(ser),
                AxisY = SeriesOnSecondary(ser) ? ChartConfigSyntax.AxisYSecondary : ChartConfigSyntax.AxisYPrimary,
                Color = ReadFillColor(ser) ?? "",
                Line = ReadLineColor(ser),
                LineWeight = ReadLineWeight(ser),
                Marker = ReadMarker(ser),
                MarkerSize = ReadMarkerSize(ser),
                MarkerColor = ReadMarkerColor(ser),
                MarkerFill = ReadMarkerFill(ser),
                ShowDataLabels = ReadShowDataLabels(ser),
                ShowValue = ReadShowVal(ser),
                ShowPercentage = ReadShowPercent(ser),
                LabelPosition = ReadLabelPos(ser),
                LabelFont = ReadLabelFont(ser),
                LabelSize = ReadLabelSize(ser),
                LabelColor = ReadLabelColor(ser),
                LabelFormat = ReadLabelFormat(ser),
                FillGradient = ReadGradient(ser),
                FillAngle = ReadGradientAngle(ser),
                ChartType = SeriesChartType(ser),
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

        private static string SnapshotFill(XElement ser)
        {
            XElement spPr = ser?.Element(C + "spPr");
            if (spPr?.Element(A + "noFill") != null)
            {
                return "none";
            }

            string hex = spPr?.Element(A + "solidFill")?.Element(A + "srgbClr")?.Attribute("val")?.Value;
            return hex ?? "";
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
