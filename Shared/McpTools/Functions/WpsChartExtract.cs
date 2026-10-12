using System;
using System.Collections.Generic;
using System.Globalization;
using System.Reflection;
using System.Text;
using WordAddIn1.OpenFiles;

namespace WordAddIn1
{
    /// <summary>
    /// WPS 文字图表晚绑定抽取：不走 Word.InlineShape / 内嵌 Excel Chart COM。
    /// Series 读取对齐 WPP（LateBoundChartSeries），产出 WordReader 相同的 &lt;Chart&gt;&lt;DataTable&gt;。
    /// </summary>
    internal static class WpsChartExtract
    {
        private const int WdInlineShapeChart = 12;
        private const int MsoChart = 3;

        public static bool IsWpsDocument(object document)
        {
            if (document == null)
            {
                return false;
            }

            try
            {
                object app = WpsCom.GetProperty(document, "Application");
                if (app == null)
                {
                    return false;
                }

                string name = Convert.ToString(WpsCom.GetProperty(app, "Name") ?? "");
                string path = Convert.ToString(WpsCom.GetProperty(app, "Path") ?? "");
                string blob = (name + " " + path).ToLowerInvariant();
                return blob.IndexOf("wps", StringComparison.Ordinal) >= 0
                    || blob.IndexOf("kingsoft", StringComparison.Ordinal) >= 0
                    || blob.IndexOf("金山", StringComparison.Ordinal) >= 0;
            }
            catch
            {
                return false;
            }
        }

        public static List<string> CollectTaggedXml(object document)
        {
            var tagged = new List<string>();
            if (document == null)
            {
                return tagged;
            }

            var found = new List<(int Start, object Chart)>();
            CollectFromCollection(WpsCom.GetProperty(document, "InlineShapes"), found, preferChartType: true);
            CollectFromCollection(WpsCom.GetProperty(document, "Shapes"), found, preferChartType: false);
            found.Sort((a, b) => a.Start.CompareTo(b.Start));

            Log("shapes=" + found.Count);
            var occupied = new HashSet<string>(StringComparer.Ordinal);
            foreach (var item in found)
            {
                string raw = ExtractDataTableXml(item.Chart);
                if (string.IsNullOrEmpty(raw))
                {
                    raw = EmptyChartXml();
                }

                string body = ExtractChartBodyContent(raw);
                string chartId = DocumentState.FindOrCreateChartId(body, occupied);
                occupied.Add(chartId);
                tagged.Add(BuildChartXmlWithId(chartId, raw));
                Log("tagged " + chartId + " start=" + item.Start);
            }

            return tagged;
        }

        private static void CollectFromCollection(
            object collection,
            List<(int Start, object Chart)> found,
            bool preferChartType)
        {
            if (collection == null)
            {
                return;
            }

            int count;
            try
            {
                count = Convert.ToInt32(WpsCom.GetProperty(collection, "Count"));
            }
            catch (Exception ex)
            {
                Log("collection Count 失败: " + ex.Message);
                return;
            }

            for (int i = 1; i <= count; i++)
            {
                object shape = WpsCom.TryGetItem(collection, i);
                if (shape == null)
                {
                    continue;
                }

                int type = TryInt(WpsCom.GetProperty(shape, "Type"), -1);
                object chart = TryGetChart(shape);
                if (chart == null)
                {
                    if (preferChartType && type != WdInlineShapeChart && type != MsoChart)
                    {
                        continue;
                    }

                    Log("shape[" + i + "] type=" + type + " 无 Chart");
                    continue;
                }

                int start = TryRangeStart(shape, i);
                found.Add((start, chart));
                Log("shape[" + i + "] type=" + type + " start=" + start);
            }
        }

        private static object TryGetChart(object shape)
        {
            try
            {
                object has = WpsCom.GetProperty(shape, "HasChart");
                if (has != null && has is bool b && !b)
                {
                    return null;
                }
            }
            catch
            {
            }

            try
            {
                object chart = WpsCom.GetProperty(shape, "Chart");
                if (chart != null)
                {
                    return chart;
                }
            }
            catch (Exception ex)
            {
                Log("Chart 属性失败: " + ex.Message);
            }

            try
            {
                object ole = WpsCom.GetProperty(shape, "OLEFormat");
                object obj = ole == null ? null : WpsCom.GetProperty(ole, "Object");
                if (obj != null && LooksLikeChart(obj))
                {
                    return obj;
                }
            }
            catch (Exception ex)
            {
                Log("OLEFormat.Object 失败: " + ex.Message);
            }

            return null;
        }

        private static bool LooksLikeChart(object obj)
        {
            try
            {
                return WpsCom.GetProperty(obj, "SeriesCollection") != null
                    || CanInvokeSeriesCollection(obj);
            }
            catch
            {
                return CanInvokeSeriesCollection(obj);
            }
        }

        private static bool CanInvokeSeriesCollection(object chart)
        {
            try
            {
                return GetSeriesCollection(chart) != null;
            }
            catch
            {
                return false;
            }
        }

        private static object GetSeriesCollection(object chart)
        {
            try
            {
                return WpsCom.Invoke(chart, "SeriesCollection");
            }
            catch
            {
            }

            try
            {
                return WpsCom.GetProperty(chart, "SeriesCollection");
            }
            catch
            {
                return null;
            }
        }

        private static int TryRangeStart(object shape, int fallback)
        {
            try
            {
                object range = WpsCom.GetProperty(shape, "Range");
                if (range != null)
                {
                    return Convert.ToInt32(WpsCom.GetProperty(range, "Start"));
                }
            }
            catch
            {
            }

            try
            {
                return Convert.ToInt32(WpsCom.GetProperty(shape, "ZOrderPosition"));
            }
            catch
            {
                return fallback;
            }
        }

        private static string ExtractDataTableXml(object chart)
        {
            try
            {
                if (LateBoundChartSeries.TryRead(chart, out LateBoundChartSeries.Grid raw, out string seriesError)
                    && raw != null
                    && raw.Series != null
                    && raw.Series.Count > 0)
                {
                    return BuildDataTableXml(ToSeriesRows(raw));
                }

                Log("Series 读失败: " + (seriesError ?? "empty"));
                string fromSheet = TryExtractFromChartDataWorkbook(chart);
                return string.IsNullOrEmpty(fromSheet) ? EmptyChartXml() : fromSheet;
            }
            catch (Exception ex)
            {
                Log("ExtractDataTableXml 失败: " + ex.Message);
                return TryExtractFromChartDataWorkbook(chart) ?? "";
            }
        }

        private static List<SeriesRow> ToSeriesRows(LateBoundChartSeries.Grid raw)
        {
            var list = new List<SeriesRow>();
            List<string> cats = raw.Categories ?? new List<string>();
            for (int i = 0; i < raw.Series.Count; i++)
            {
                LateBoundChartSeries.Col one = raw.Series[i];
                var row = new SeriesRow
                {
                    Name = string.IsNullOrEmpty(one.Name) ? ("系列" + (i + 1)) : one.Name,
                    AxisY = "primary",
                    Color = NormalizeWordColor(one.Color),
                    ChartType = "",
                    ShowDataLabels = "false",
                    Points = new List<PointRow>()
                };
                List<string> values = one.Values ?? new List<string>();
                int n = Math.Max(cats.Count, values.Count);
                for (int j = 0; j < n; j++)
                {
                    double y = 0;
                    if (j < values.Count)
                    {
                        double.TryParse(values[j], NumberStyles.Any, CultureInfo.InvariantCulture, out y);
                    }

                    row.Points.Add(new PointRow
                    {
                        X = j < cats.Count ? (cats[j] ?? "") : "",
                        Y = y
                    });
                }

                list.Add(row);
            }

            return list;
        }

        private static string NormalizeWordColor(string color)
        {
            if (string.IsNullOrEmpty(color) || color == "none")
            {
                return "";
            }

            return color[0] == '#' ? color.Substring(1) : color;
        }

        private static string TryExtractFromChartDataWorkbook(object chart)
        {
            try
            {
                object chartData = WpsCom.GetProperty(chart, "ChartData");
                if (chartData == null)
                {
                    return null;
                }

                object workbook = null;
                try
                {
                    workbook = WpsCom.GetProperty(chartData, "Workbook");
                }
                catch
                {
                }

                if (workbook == null)
                {
                    return null;
                }

                object sheets = WpsCom.GetProperty(workbook, "Worksheets")
                    ?? WpsCom.GetProperty(workbook, "Sheets");
                object sheet = WpsCom.TryGetItem(sheets, 1);
                if (sheet == null)
                {
                    return null;
                }

                object used = WpsCom.GetProperty(sheet, "UsedRange");
                object raw = used == null ? null : (TryGet(used, "Value2") ?? TryGet(used, "Value"));
                Array table = ToArray(raw);
                if (table == null)
                {
                    return null;
                }

                return BuildDataTableFromSheet(table);
            }
            catch (Exception ex)
            {
                Log("ChartData.Workbook 失败: " + ex.Message);
                return null;
            }
        }

        private static string BuildDataTableFromSheet(Array table)
        {
            int r0 = table.GetLowerBound(0);
            int r1 = table.GetUpperBound(0);
            int c0 = table.Rank > 1 ? table.GetLowerBound(1) : 0;
            int c1 = table.Rank > 1 ? table.GetUpperBound(1) : 0;
            if (r1 < r0)
            {
                return EmptyChartXml();
            }

            var seriesInfoList = new List<SeriesRow>();
            if (table.Rank == 1)
            {
                return EmptyChartXml();
            }

            for (int c = c0 + 1; c <= c1; c++)
            {
                string name = Convert.ToString(table.GetValue(r0, c)) ?? ("系列" + (c - c0));
                var series = new SeriesRow
                {
                    Name = name,
                    AxisY = "primary",
                    Color = "",
                    ChartType = "",
                    ShowDataLabels = "false",
                    Points = new List<PointRow>()
                };
                for (int r = r0 + 1; r <= r1; r++)
                {
                    string x = Convert.ToString(table.GetValue(r, c0)) ?? "";
                    double y;
                    double.TryParse(
                        Convert.ToString(table.GetValue(r, c)),
                        NumberStyles.Any,
                        CultureInfo.InvariantCulture,
                        out y);
                    series.Points.Add(new PointRow { X = x, Y = y });
                }

                seriesInfoList.Add(series);
            }

            return seriesInfoList.Count == 0 ? EmptyChartXml() : BuildDataTableXml(seriesInfoList);
        }

        private static string BuildDataTableXml(List<SeriesRow> seriesInfoList)
        {
            var dataTableXml = new StringBuilder();
            dataTableXml.Append("<Chart>\n");
            dataTableXml.Append("  <DataTable>\n");
            dataTableXml.Append("    <Header>\n");
            dataTableXml.Append("      <Column type=\"category\">类别</Column>\n");
            foreach (SeriesRow seriesInfo in seriesInfoList)
            {
                dataTableXml.Append("      <Column type=\"value\"");
                dataTableXml.Append(" axisY=\"" + seriesInfo.AxisY + "\"");
                if (!string.IsNullOrEmpty(seriesInfo.Color))
                {
                    dataTableXml.Append(" color=\"" + seriesInfo.Color + "\"");
                }

                if (!string.IsNullOrEmpty(seriesInfo.ChartType))
                {
                    dataTableXml.Append(" chartType=\"" + seriesInfo.ChartType + "\"");
                }

                if (!string.IsNullOrEmpty(seriesInfo.ShowDataLabels))
                {
                    dataTableXml.Append(" showDataLabels=\"" + seriesInfo.ShowDataLabels + "\"");
                }

                dataTableXml.Append(">" + EscapeXmlText(seriesInfo.Name) + "</Column>\n");
            }

            dataTableXml.Append("    </Header>\n");
            dataTableXml.Append("    <Rows>\n");
            if (seriesInfoList.Count > 0 && seriesInfoList[0].Points.Count > 0)
            {
                int rowCount = seriesInfoList[0].Points.Count;
                for (int rowIndex = 0; rowIndex < rowCount; rowIndex++)
                {
                    var rowValues = new List<string>();
                    rowValues.Add(seriesInfoList[0].Points[rowIndex].X ?? "");
                    foreach (SeriesRow seriesInfo in seriesInfoList)
                    {
                        if (rowIndex < seriesInfo.Points.Count)
                        {
                            rowValues.Add(seriesInfo.Points[rowIndex].Y.ToString("F2", CultureInfo.InvariantCulture));
                        }
                        else
                        {
                            rowValues.Add("0");
                        }
                    }

                    dataTableXml.Append("      <Row>" + EscapeXmlText(string.Join(",", rowValues)) + "</Row>\n");
                }
            }

            dataTableXml.Append("    </Rows>\n");
            dataTableXml.Append("  </DataTable>\n");
            dataTableXml.Append("</Chart>\n");
            return dataTableXml.ToString();
        }

        private static string EmptyChartXml()
        {
            return "<Chart>\n  <DataTable>\n  </DataTable>\n</Chart>\n";
        }

        private static string ExtractChartBodyContent(string rawChartXml)
        {
            if (string.IsNullOrEmpty(rawChartXml))
            {
                return "";
            }

            int start = rawChartXml.IndexOf("<DataTable", StringComparison.OrdinalIgnoreCase);
            if (start < 0)
            {
                return rawChartXml.Trim();
            }

            int end = rawChartXml.LastIndexOf("</DataTable>", StringComparison.OrdinalIgnoreCase);
            if (end < start)
            {
                return rawChartXml.Trim();
            }

            return rawChartXml.Substring(start, end + "</DataTable>".Length - start).Trim();
        }

        private static string BuildChartXmlWithId(string chartId, string rawChartXml)
        {
            if (string.IsNullOrEmpty(rawChartXml))
            {
                return "";
            }

            if (rawChartXml.StartsWith("<Chart", StringComparison.OrdinalIgnoreCase))
            {
                int closeBracket = rawChartXml.IndexOf('>');
                if (closeBracket > 0)
                {
                    return "<Chart id=\"" + chartId + "\">" + rawChartXml.Substring(closeBracket + 1);
                }
            }

            return "<Chart id=\"" + chartId + "\">\n" + rawChartXml + "</Chart>\n";
        }

        private static Array ToArray(object value)
        {
            if (value == null)
            {
                return null;
            }

            if (value is Array arr)
            {
                return arr;
            }

            object inner = TryGet(value, "Value2") ?? TryGet(value, "Value");
            return inner as Array;
        }

        private static object TryGet(object target, string name)
        {
            if (target == null)
            {
                return null;
            }

            try
            {
                return WpsCom.GetProperty(target, name);
            }
            catch
            {
                try
                {
                    return target.GetType().InvokeMember(
                        name,
                        BindingFlags.GetProperty | BindingFlags.Public | BindingFlags.Instance,
                        null,
                        target,
                        null);
                }
                catch
                {
                    return null;
                }
            }
        }

        private static int TryInt(object value, int fallback)
        {
            if (value == null)
            {
                return fallback;
            }

            try
            {
                return Convert.ToInt32(value);
            }
            catch
            {
                return fallback;
            }
        }

        private static string EscapeXmlText(string value)
        {
            if (string.IsNullOrEmpty(value))
            {
                return "";
            }

            return value
                .Replace("&", "&amp;")
                .Replace("<", "&lt;")
                .Replace(">", "&gt;");
        }

        private static void Log(string message)
        {
            System.Diagnostics.Debug.WriteLine("[WpsChartExtract] " + message);
        }

        private sealed class SeriesRow
        {
            public string Name { get; set; }

            public string AxisY { get; set; }

            public string Color { get; set; }

            public string ChartType { get; set; }

            public string ShowDataLabels { get; set; }

            public List<PointRow> Points { get; set; }
        }

        private sealed class PointRow
        {
            public string X { get; set; }

            public double Y { get; set; }
        }
    }
}
