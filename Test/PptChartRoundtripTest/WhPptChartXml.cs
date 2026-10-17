using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using WordAddIn1.PresentationHost;

namespace PptChartRoundtripTest
{
    /// <summary>
    /// PPT 图表用例的 HTML 改写成文字 ChartConfig。断言读抽出的 XML，不走幻灯片 COM。
    /// </summary>
    internal static class WhPptChartXml
    {
        public static bool TryExport(string html, out string xml, out string error)
        {
            xml = null;
            if (!ChartHtml.TryParseFirstChart(html, out PptHtmlApplyNode node, out error))
            {
                return false;
            }

            xml = Build(node);
            return true;
        }

        public static string Match(ChartCase one, string extracted)
        {
            if (one == null)
            {
                return "用例空";
            }

            if (string.IsNullOrWhiteSpace(extracted))
            {
                return "抽出 XML 空";
            }

            XDocument doc;
            try
            {
                doc = XDocument.Parse(extracted);
            }
            catch (Exception ex)
            {
                return "抽出 XML 解析失败: " + ex.Message;
            }

            XElement chart = doc.Descendants().FirstOrDefault(e => e.Name.LocalName == "Chart");
            if (chart == null)
            {
                return "抽出 XML 没有 Chart";
            }

            string type = Text(chart, "Type");
            if (string.Equals(type, "pie", StringComparison.OrdinalIgnoreCase))
            {
                type = "pie2d";
            }

            string expectType = one.ExpectType;
            if (string.Equals(expectType, "pie", StringComparison.OrdinalIgnoreCase))
            {
                expectType = "pie2d";
            }

            if (!string.IsNullOrEmpty(expectType)
                && !string.Equals(expectType, type, StringComparison.OrdinalIgnoreCase))
            {
                return "类型期望 " + one.ExpectType + " 实际 " + (type ?? "(null)");
            }

            List<XElement> rows = chart.Descendants().Where(e => e.Name.LocalName == "Row").ToList();
            if (rows.Count < one.ExpectRows)
            {
                return "行数期望 ≥" + one.ExpectRows + " 实际 " + rows.Count;
            }

            List<XElement> series = ValueColumns(chart);
            if (series.Count < one.ExpectSeries)
            {
                return "系列数期望 ≥" + one.ExpectSeries + " 实际 " + series.Count;
            }

            if (!string.IsNullOrEmpty(one.ExpectLineAxis))
            {
                string axis = series.Count > 1 ? Attr(series[1], "axisY") : null;
                if (string.IsNullOrEmpty(axis))
                {
                    axis = "y";
                }

                if (!string.Equals(one.ExpectLineAxis, axis, StringComparison.OrdinalIgnoreCase))
                {
                    return "折线轴期望 " + one.ExpectLineAxis + " 实际 " + axis;
                }
            }

            if (rows.Count > 0 && !string.IsNullOrEmpty(one.FirstCategory))
            {
                string row0 = (rows[0].Value ?? "").Trim();
                if (row0.IndexOf(one.FirstCategory, StringComparison.Ordinal) < 0)
                {
                    return "首类期望 " + one.FirstCategory + " 实际 " + row0;
                }
            }

            if (one.ExpectRows > 0 && rows.Count >= one.ExpectRows)
            {
                string last = (rows[one.ExpectRows - 1].Value ?? "").Trim();
                if (!ContainsNumber(last, one.LastValue))
                {
                    return "末值期望 " + one.LastValue.ToString(CultureInfo.InvariantCulture) + " 实际 " + last;
                }
            }

            return MatchAttrs(one.Attrs, chart, series);
        }

        private static string Build(PptHtmlApplyNode node)
        {
            PptHtmlChartFormat fmt = node.ChartFormat;
            PptHtmlChartGrid grid = node.ChartGrid;
            var sb = new StringBuilder();
            sb.Append("<ChartConfig><Chart>");
            string type = First(node.ChartType, fmt == null ? null : fmt.ChartType, "column");
            if (string.Equals(type, "pie", StringComparison.OrdinalIgnoreCase))
            {
                type = "pie2d";
            }

            El(sb, "Type", type);
            if (fmt != null)
            {
                El(sb, "Title", fmt.Title);
                El(sb, "Theme", fmt.Theme);
                El(sb, "LegendPosition", fmt.Legend);
                El(sb, "ShowDataLabels", fmt.ShowDataLabels);
                El(sb, "ShowValue", fmt.ShowValue);
                El(sb, "ShowPercentage", fmt.ShowPercentage);
                El(sb, "PlotColor", fmt.PlotColor);
                El(sb, "TitleFontSize", fmt.TitleFontSize);
                El(sb, "TitleFontBold", fmt.TitleFontBold);
                El(sb, "TitleFontColor", fmt.TitleFontColor);
                El(sb, "Gridlines", fmt.Gridlines);
                El(sb, "GapWidth", fmt.GapWidth);
                El(sb, "Overlap", fmt.Overlap);
                El(sb, "PlotBox", fmt.PlotBox);
                El(sb, "PlotInside", fmt.PlotInside);
                El(sb, "ChartAreaColor", fmt.ChartAreaColor);
                El(sb, "ChartStyle", fmt.ChartStyle);
                El(sb, "LegendFontColor", fmt.LegendFontColor);
                El(sb, "DataMarkers", fmt.DataMarkers);
                El(sb, "MarkerSize", fmt.MarkerSize);
                El(sb, "LineWeight", fmt.ChartLineWeight);
                El(sb, "YAxisMin", fmt.AxisYMin);
                El(sb, "YAxisMax", fmt.AxisYMax);
                El(sb, "YAxisMajorUnit", fmt.AxisYMajorUnit);
                El(sb, "Explosion", fmt.Explosion);
                if (!string.IsNullOrEmpty(fmt.FillMissing))
                {
                    El(sb, "FillMissingValues", string.Equals(fmt.FillMissing, "zero", StringComparison.OrdinalIgnoreCase) ? "true" : "false");
                }

                Axis(sb, "AxisX", fmt.AxisX, fmt.AxisXType, fmt.AxisXFormat, fmt.AxisXTickCount, fmt.AxisXTickSpacing, fmt.AxisXBetween, fmt.AxisXStyle);
                Axis(sb, "AxisY", fmt.AxisY, null, null, null, null, null, fmt.AxisYStyle);
                Axis(sb, "AxisY2", fmt.AxisYSecondary, null, null, null, null, null, fmt.AxisY2Style);
            }

            sb.Append("<DataTable><Header>");
            sb.Append("<Column type=\"category\">");
            string catName = "类别";
            if (grid != null && grid.Columns != null)
            {
                for (int i = 0; i < grid.Columns.Count; i++)
                {
                    if (grid.Columns[i] != null && grid.Columns[i].Role == "category" && !string.IsNullOrEmpty(grid.Columns[i].Name))
                    {
                        catName = grid.Columns[i].Name;
                        break;
                    }
                }
            }

            sb.Append(Esc(catName)).Append("</Column>");
            if (grid != null && grid.Columns != null)
            {
                for (int i = 0; i < grid.Columns.Count; i++)
                {
                    PptHtmlChartColumn col = grid.Columns[i];
                    if (col == null || col.Role == "category")
                    {
                        continue;
                    }

                    sb.Append("<Column type=\"value\"");
                    AttrEl(sb, "axisY", string.IsNullOrEmpty(col.AxisY) ? "y" : col.AxisY);
                    AttrEl(sb, "chartType", col.SeriesType);
                    AttrEl(sb, "color", col.Color);
                    AttrEl(sb, "showDataLabels", col.ShowDataLabels);
                    AttrEl(sb, "showValue", col.ShowValue);
                    AttrEl(sb, "showPercentage", col.ShowPercentage);
                    AttrEl(sb, "fillGradient", col.FillGradient);
                    AttrEl(sb, "fillAngle", col.FillAngle);
                    AttrEl(sb, "line", col.Line);
                    AttrEl(sb, "lineWeight", col.LineWeight);
                    AttrEl(sb, "marker", col.Marker);
                    AttrEl(sb, "markerSize", col.MarkerSize);
                    AttrEl(sb, "markerColor", col.MarkerColor);
                    AttrEl(sb, "markerFill", col.MarkerFill);
                    AttrEl(sb, "labelPosition", col.LabelPosition);
                    AttrEl(sb, "labelFont", col.LabelFont);
                    AttrEl(sb, "labelSize", col.LabelSize);
                    AttrEl(sb, "labelColor", col.LabelColor);
                    AttrEl(sb, "labelFormat", col.LabelFormat);
                    sb.Append(">").Append(Esc(col.Name ?? "系列")).Append("</Column>");
                }
            }

            sb.Append("</Header><Rows>");
            if (grid != null && grid.Rows != null)
            {
                for (int r = 0; r < grid.Rows.Count; r++)
                {
                    sb.Append("<Row>");
                    List<string> cells = grid.Rows[r];
                    if (cells != null)
                    {
                        for (int c = 0; c < cells.Count; c++)
                        {
                            if (c > 0)
                            {
                                sb.Append(",");
                            }

                            sb.Append(Esc(cells[c] ?? ""));
                        }
                    }

                    sb.Append("</Row>");
                }
            }

            sb.Append("</Rows></DataTable></Chart></ChartConfig>");
            return sb.ToString();
        }

        private static string MatchAttrs(IList<AttrCheck> checks, XElement chart, List<XElement> series)
        {
            if (checks == null)
            {
                return null;
            }

            var fails = new List<string>();
            for (int i = 0; i < checks.Count; i++)
            {
                AttrCheck one = checks[i];
                if (one == null || string.Equals(one.Mode, "applyonly", StringComparison.Ordinal))
                {
                    continue;
                }

                string actual = Lookup(chart, series, one.Key);
                if (!Same(one, actual))
                {
                    fails.Add(one.Key + " 期望 " + (one.Expected ?? "(null)") + " 实际 " + (actual ?? "(null)"));
                }
            }

            if (fails.Count == 0)
            {
                return null;
            }

            int n = Math.Min(4, fails.Count);
            return string.Join("；", fails.GetRange(0, n));
        }

        private static string Lookup(XElement chart, List<XElement> series, string key)
        {
            if (string.IsNullOrEmpty(key))
            {
                return null;
            }

            if (key.StartsWith("S", StringComparison.Ordinal) && key.IndexOf('.') > 0)
            {
                int dot = key.IndexOf('.');
                if (int.TryParse(key.Substring(1, dot - 1), NumberStyles.Integer, CultureInfo.InvariantCulture, out int idx)
                    && idx >= 0 && idx < series.Count)
                {
                    return SeriesProp(series[idx], key.Substring(dot + 1));
                }

                return null;
            }

            if (key.StartsWith("AxisXStyle.", StringComparison.Ordinal))
            {
                return AxisProp(chart, "AxisX", key.Substring("AxisXStyle.".Length));
            }

            if (key.StartsWith("AxisYStyle.", StringComparison.Ordinal))
            {
                return AxisProp(chart, "AxisY", key.Substring("AxisYStyle.".Length));
            }

            if (key.StartsWith("AxisY2Style.", StringComparison.Ordinal))
            {
                return AxisProp(chart, "AxisY2", key.Substring("AxisY2Style.".Length));
            }

            switch (key)
            {
                case "Legend":
                    return Text(chart, "LegendPosition");
                case "AxisX":
                    return AxisProp(chart, "AxisX", "Title");
                case "AxisY":
                    return AxisProp(chart, "AxisY", "Title");
                case "AxisYSecondary":
                    return AxisProp(chart, "AxisY2", "Title");
                case "FillMissing":
                    return Text(chart, "FillMissingValues");
                case "ChartLineWeight":
                    return Text(chart, "LineWeight");
                default:
                    return Text(chart, key);
            }
        }

        private static string SeriesProp(XElement col, string name)
        {
            switch (name)
            {
                case "Color":
                    return Attr(col, "color");
                case "AxisY":
                    return Attr(col, "axisY");
                case "ShowDataLabels":
                    return Attr(col, "showDataLabels");
                case "ShowValue":
                    return Attr(col, "showValue");
                case "ShowPercentage":
                    return Attr(col, "showPercentage");
                case "FillGradient":
                    return Attr(col, "fillGradient");
                case "FillAngle":
                    return Attr(col, "fillAngle");
                case "Line":
                    return Attr(col, "line");
                case "LineWeight":
                    return Attr(col, "lineWeight");
                case "Marker":
                    return Attr(col, "marker");
                case "MarkerSize":
                    return Attr(col, "markerSize");
                case "MarkerColor":
                    return Attr(col, "markerColor");
                case "MarkerFill":
                    return Attr(col, "markerFill");
                case "LabelPosition":
                    return Attr(col, "labelPosition");
                case "LabelFont":
                    return Attr(col, "labelFont");
                case "LabelSize":
                    return Attr(col, "labelSize");
                case "LabelColor":
                    return Attr(col, "labelColor");
                case "LabelFormat":
                    return Attr(col, "labelFormat");
                default:
                    return null;
            }
        }

        private static string AxisProp(XElement chart, string axisName, string child)
        {
            XElement axis = chart.Elements().FirstOrDefault(e => e.Name.LocalName == axisName);
            if (axis == null)
            {
                return null;
            }

            return Text(axis, child);
        }

        private static bool Same(AttrCheck one, string actual)
        {
            string expected = one.Expected;
            if (expected == null)
            {
                return string.IsNullOrEmpty(actual);
            }

            if (actual == null)
            {
                return false;
            }

            string mode = one.Mode ?? "exact";
            if (mode == "hex")
            {
                return string.Equals(NormHex(expected), NormHex(actual), StringComparison.OrdinalIgnoreCase);
            }

            if (mode == "num")
            {
                if (double.TryParse(expected, NumberStyles.Float, CultureInfo.InvariantCulture, out double e)
                    && double.TryParse(actual, NumberStyles.Float, CultureInfo.InvariantCulture, out double a))
                {
                    return Math.Abs(e - a) < 0.001;
                }
            }

            return string.Equals(expected.Trim(), actual.Trim(), StringComparison.OrdinalIgnoreCase);
        }

        private static string NormHex(string raw)
        {
            if (string.IsNullOrEmpty(raw))
            {
                return "";
            }

            string s = raw.Trim();
            if (s.StartsWith("#", StringComparison.Ordinal))
            {
                s = s.Substring(1);
            }

            return s;
        }

        private static bool ContainsNumber(string row, double value)
        {
            if (string.IsNullOrEmpty(row))
            {
                return false;
            }

            string[] parts = row.Split(',');
            for (int i = 0; i < parts.Length; i++)
            {
                if (double.TryParse(parts[i].Trim(), NumberStyles.Float, CultureInfo.InvariantCulture, out double n)
                    && Math.Abs(n - value) < 0.001)
                {
                    return true;
                }
            }

            return false;
        }

        private static List<XElement> ValueColumns(XElement chart)
        {
            return chart.Descendants()
                .Where(e => e.Name.LocalName == "Column"
                    && string.Equals(Attr(e, "type"), "value", StringComparison.OrdinalIgnoreCase))
                .ToList();
        }

        private static void Axis(
            StringBuilder sb,
            string name,
            string title,
            string type,
            string format,
            string tickCount,
            string tickSpacing,
            string between,
            PptHtmlAxisExtras style)
        {
            if (string.IsNullOrEmpty(title) && string.IsNullOrEmpty(type) && string.IsNullOrEmpty(format)
                && string.IsNullOrEmpty(tickCount) && string.IsNullOrEmpty(tickSpacing) && string.IsNullOrEmpty(between)
                && style == null)
            {
                return;
            }

            sb.Append("<").Append(name).Append(">");
            El(sb, "Title", title);
            El(sb, "Type", type);
            El(sb, "Format", format);
            El(sb, "TickLabelCount", tickCount);
            El(sb, "TickLabelSpacing", tickSpacing);
            El(sb, "BetweenCategories", between);
            if (style != null)
            {
                El(sb, "Visible", style.Visible);
                El(sb, "TickFont", style.TickFont);
                El(sb, "TickColor", style.TickColor);
                El(sb, "TickSize", style.TickSize);
                El(sb, "TickPosition", style.TickPosition);
                El(sb, "MajorTick", style.MajorTick);
                El(sb, "MinorTick", style.MinorTick);
                El(sb, "Grid", style.Grid);
                El(sb, "GridColor", style.GridColor);
                El(sb, "Line", style.Line);
                El(sb, "LineWeight", style.LineWeight);
            }

            sb.Append("</").Append(name).Append(">");
        }

        private static void El(StringBuilder sb, string name, string value)
        {
            if (string.IsNullOrEmpty(value))
            {
                return;
            }

            sb.Append("<").Append(name).Append(">").Append(Esc(value)).Append("</").Append(name).Append(">");
        }

        private static void AttrEl(StringBuilder sb, string name, string value)
        {
            if (string.IsNullOrEmpty(value))
            {
                return;
            }

            sb.Append(" ").Append(name).Append("=\"").Append(Esc(value)).Append("\"");
        }

        private static string Text(XElement parent, string child)
        {
            XElement el = parent.Elements().FirstOrDefault(e => e.Name.LocalName == child);
            return el == null ? null : (el.Value ?? "").Trim();
        }

        private static string Attr(XElement el, string name)
        {
            XAttribute a = el.Attributes().FirstOrDefault(x => x.Name.LocalName == name);
            return a == null ? null : a.Value;
        }

        private static string First(params string[] values)
        {
            for (int i = 0; i < values.Length; i++)
            {
                if (!string.IsNullOrEmpty(values[i]))
                {
                    return values[i];
                }
            }

            return "column";
        }

        private static string Esc(string raw)
        {
            return (raw ?? "")
                .Replace("&", "&amp;")
                .Replace("<", "&lt;")
                .Replace(">", "&gt;")
                .Replace("\"", "&quot;");
        }
    }
}
