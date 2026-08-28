using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Xml.Linq;
using WordAddIn1.OpenFiles;

namespace WordAddIn1.PresentationHost
{
    internal sealed class PptHtmlChartColumn
    {
        public string Role { get; set; }

        public string Name { get; set; }

        public string Color { get; set; }

        public string SeriesType { get; set; }

        public string AxisY { get; set; }

        public string ShowDataLabels { get; set; }
    }

    internal sealed class PptHtmlChartGrid
    {
        public List<PptHtmlChartColumn> Columns { get; set; }

        public List<List<string>> Rows { get; set; }

        public bool Truncated { get; set; }

        public bool IsPourable
        {
            get
            {
                if (Columns == null || Columns.Count < 2 || Rows == null || Rows.Count < 1)
                {
                    return false;
                }

                return Columns.Any(c => c != null && c.Role != "category");
            }
        }
    }

    internal sealed class PptHtmlChartFormat
    {
        public string ChartType { get; set; }

        public string Title { get; set; }

        public string Theme { get; set; }

        public string Legend { get; set; }

        public string ShowDataLabels { get; set; }

        public string PlotColor { get; set; }

        public string TitleFontSize { get; set; }

        public string TitleFontBold { get; set; }

        public string TitleFontColor { get; set; }

        public string Gridlines { get; set; }

        public string GapWidth { get; set; }

        public string DataMarkers { get; set; }

        public string MarkerSize { get; set; }

        public string ChartLineWeight { get; set; }

        public string AxisYMin { get; set; }

        public string AxisYMax { get; set; }

        public string AxisYMajorUnit { get; set; }

        public string DataLabelType { get; set; }

        public string Explosion { get; set; }

        public string FillMissing { get; set; }

        public string AxisX { get; set; }

        public string AxisXType { get; set; }

        public string AxisXFormat { get; set; }

        public string AxisXTickCount { get; set; }

        public string AxisXTickSpacing { get; set; }

        public string AxisXBetween { get; set; }

        public string AxisY { get; set; }

        public string AxisYSecondary { get; set; }
    }

    internal sealed class PptHtmlChartReadModel
    {
        public PptHtmlChartFormat Format { get; set; }

        public PptHtmlChartGrid Grid { get; set; }
    }

    internal static class PptHtmlChartIo
    {
        public const int MaxRows = 30;

        public const int MaxCols = 8;

        private const int XlColumnClustered = 51;
        private const int XlBarClustered = 57;
        private const int XlLine = 4;
        private const int XlLineMarkers = 65;
        private const int XlDataLabelsShowValue = 2;
        private const int XlPie = 5;
        private const int Xl3DPie = -4102;

        private const int XlCategory = 1;
        private const int XlValue = 2;
        private const int XlPrimary = 1;
        private const int XlSecondary = 2;
        private const int XlCategoryScale = 2;
        private const int XlTickLabelPositionNone = -4142;
        private const int XlTickLabelPositionNextToAxis = 4;
        private const int XlTickMarkNone = -4142;
        private const int MsoFillGradient = 3;

        private sealed class ChartStyleSnap
        {
            public int? ChartStyle { get; set; }

            public bool? HasTitle { get; set; }

            public string TitleFontColor { get; set; }

            public string TitleFontSize { get; set; }

            public bool? TitleFontBold { get; set; }

            public bool? HasLegend { get; set; }

            public int? LegendPosition { get; set; }

            public string LegendFontColor { get; set; }

            public bool? ChartAreaFillVisible { get; set; }

            public int? ChartAreaFillRgb { get; set; }

            public bool? PlotFillVisible { get; set; }

            public int? PlotFillRgb { get; set; }

            public float? PlotLeft { get; set; }

            public float? PlotTop { get; set; }

            public float? PlotWidth { get; set; }

            public float? PlotHeight { get; set; }

            public float? PlotInsideLeft { get; set; }

            public float? PlotInsideTop { get; set; }

            public float? PlotInsideWidth { get; set; }

            public float? PlotInsideHeight { get; set; }

            public int? GapWidth { get; set; }

            public int? Overlap { get; set; }

            public AxisStyleSnap Category { get; set; }

            public AxisStyleSnap Value { get; set; }

            public AxisStyleSnap ValueSecondary { get; set; }

            public List<SeriesStyleSnap> Series { get; set; }
        }

        private sealed class AxisStyleSnap
        {
            public bool? Deleted { get; set; }

            public bool? HasTitle { get; set; }

            public string Title { get; set; }

            public string TickFontName { get; set; }

            public int? TickFontColor { get; set; }

            public double? TickFontSize { get; set; }

            public int? TickLabelPosition { get; set; }

            public int? MajorTickMark { get; set; }

            public int? MinorTickMark { get; set; }

            public string NumberFormat { get; set; }

            public bool? HasMajorGridlines { get; set; }

            public int? MajorGridlineRgb { get; set; }

            public bool? LineVisible { get; set; }

            public int? LineRgb { get; set; }

            public double? LineWeight { get; set; }
        }

        private sealed class SeriesStyleSnap
        {
            public int? ChartType { get; set; }

            public int? AxisGroup { get; set; }

            public FillSnap Fill { get; set; }

            public LineSnap Line { get; set; }

            public int? MarkerStyle { get; set; }

            public int? MarkerSize { get; set; }

            public int? MarkerForeRgb { get; set; }

            public int? MarkerBackRgb { get; set; }

            public bool? HasDataLabels { get; set; }

            public int? DataLabelPosition { get; set; }

            public string DataLabelFontName { get; set; }

            public double? DataLabelFontSize { get; set; }

            public int? DataLabelFontColor { get; set; }

            public string DataLabelNumberFormat { get; set; }
        }

        private sealed class FillSnap
        {
            public bool? Visible { get; set; }

            public int? FillType { get; set; }

            public int? SolidRgb { get; set; }

            public double? Angle { get; set; }

            public List<GradientStopSnap> Stops { get; set; }
        }

        private sealed class LineSnap
        {
            public bool? Visible { get; set; }

            public int? Rgb { get; set; }

            public double? Weight { get; set; }

            public List<GradientStopSnap> Stops { get; set; }
        }

        private sealed class GradientStopSnap
        {
            public double Position { get; set; }

            public int Rgb { get; set; }

            public double Transparency { get; set; }
        }

        public static bool TryParseType(string raw, out int xlType, out string canonical, out string error)
        {
            xlType = XlColumnClustered;
            canonical = "column";
            error = null;
            if (string.IsNullOrWhiteSpace(raw))
            {
                return true;
            }

            switch (raw.Trim().ToLowerInvariant())
            {
                case "column":
                case "column_clustered":
                    xlType = XlColumnClustered;
                    canonical = "column";
                    return true;
                case "bar":
                    xlType = XlBarClustered;
                    canonical = "bar";
                    return true;
                case "line":
                    xlType = XlLine;
                    canonical = "line";
                    return true;
                case "pie":
                    xlType = XlPie;
                    canonical = "pie";
                    return true;
                case "pie3d":
                    xlType = Xl3DPie;
                    canonical = "pie";
                    return true;
                default:
                    error = "data-chart-type 仅支持 column/bar/line/pie";
                    return false;
            }
        }

        public static string CanonicalTypeFromXl(int xlType)
        {
            if (xlType == XlBarClustered)
            {
                return "bar";
            }

            if (xlType == XlLine)
            {
                return "line";
            }

            if (xlType == XlPie || xlType == Xl3DPie)
            {
                return "pie";
            }

            return "column";
        }

        public static IEnumerable<XElement> EnumerateTableRows(XElement table)
        {
            if (table == null)
            {
                yield break;
            }

            foreach (XElement child in table.Elements())
            {
                string name = child.Name.LocalName;
                if (string.Equals(name, "tr", StringComparison.OrdinalIgnoreCase))
                {
                    yield return child;
                    continue;
                }

                if (string.Equals(name, "thead", StringComparison.OrdinalIgnoreCase)
                    || string.Equals(name, "tbody", StringComparison.OrdinalIgnoreCase)
                    || string.Equals(name, "tfoot", StringComparison.OrdinalIgnoreCase))
                {
                    foreach (XElement tr in child.Elements().Where(e =>
                        string.Equals(e.Name.LocalName, "tr", StringComparison.OrdinalIgnoreCase)))
                    {
                        yield return tr;
                    }
                }
            }
        }

        public static bool TryParseGrid(XElement table, bool enforceLimitFail, out PptHtmlChartGrid grid, out string error)
        {
            grid = null;
            error = null;
            if (table == null)
            {
                error = "新建 chart 必须在节点内嵌 <table> 灌数，不能建空图";
                return false;
            }

            var columns = new List<PptHtmlChartColumn>();
            var dataRows = new List<List<string>>();
            bool first = true;
            foreach (XElement tr in EnumerateTableRows(table))
            {
                var cells = tr.Elements().Where(e =>
                    string.Equals(e.Name.LocalName, "td", StringComparison.OrdinalIgnoreCase)
                    || string.Equals(e.Name.LocalName, "th", StringComparison.OrdinalIgnoreCase)).ToList();
                if (cells.Count == 0)
                {
                    continue;
                }

                if (first)
                {
                    for (int i = 0; i < cells.Count; i++)
                    {
                        string role = GetAttr(cells[i], "data-col");
                        if (string.IsNullOrWhiteSpace(role))
                        {
                            role = i == 0 ? "category" : "value";
                        }

                        role = role.Trim().ToLowerInvariant();
                        if (role != "category" && role != "value")
                        {
                            error = "th data-col 只能是 category 或 value";
                            return false;
                        }

                        columns.Add(new PptHtmlChartColumn
                        {
                            Role = role,
                            Name = InnerText(cells[i]),
                            Color = NormalizeHexOrNull(GetAttr(cells[i], "data-color")),
                            SeriesType = GetAttr(cells[i], "data-series-type"),
                            AxisY = GetAttr(cells[i], "data-axis-y"),
                            ShowDataLabels = GetAttr(cells[i], "data-show-data-labels")
                        });
                    }

                    first = false;
                    continue;
                }

                var row = new List<string>();
                for (int i = 0; i < cells.Count; i++)
                {
                    row.Add(InnerText(cells[i]));
                }

                dataRows.Add(row);
            }

            if (columns.Count < 2 || dataRows.Count < 1)
            {
                error = "chart 内嵌表至少要有表头行 + 一行数字。"
                    + "写成 <tr><th>类别</th><th>系列</th></tr><tr><td>Q1</td><td>120</td></tr>；"
                    + "thead/tbody 也可以";
                return false;
            }

            if (!columns.Any(c => c.Role == "value"))
            {
                error = "chart 内嵌表至少要有一列 value";
                return false;
            }

            bool over = columns.Count > MaxCols || (dataRows.Count + 1) > MaxRows;
            if (over && enforceLimitFail)
            {
                error = "chart 内嵌表最多 " + MaxRows + " 行 × " + MaxCols + " 列";
                return false;
            }

            if (over)
            {
                if (columns.Count > MaxCols)
                {
                    columns = columns.GetRange(0, MaxCols);
                }

                int maxData = MaxRows - 1;
                if (dataRows.Count > maxData)
                {
                    dataRows = dataRows.GetRange(0, maxData);
                }
            }

            int colCount = columns.Count;
            for (int r = 0; r < dataRows.Count; r++)
            {
                while (dataRows[r].Count < colCount)
                {
                    dataRows[r].Add("");
                }

                if (dataRows[r].Count > colCount)
                {
                    dataRows[r] = dataRows[r].GetRange(0, colCount);
                }

                for (int c = 0; c < colCount; c++)
                {
                    if (columns[c].Role != "value")
                    {
                        continue;
                    }

                    string raw = dataRows[r][c];
                    if (string.IsNullOrWhiteSpace(raw)
                        || !double.TryParse(raw.Trim(), NumberStyles.Float, CultureInfo.InvariantCulture, out _))
                    {
                        error = "第 " + (r + 2) + " 行第 " + (c + 1) + " 列不是数字";
                        return false;
                    }
                }
            }

            grid = new PptHtmlChartGrid
            {
                Columns = columns,
                Rows = dataRows,
                Truncated = over
            };
            return true;
        }

        public static PptHtmlChartFormat ParseFormat(XElement el)
        {
            if (el == null)
            {
                return new PptHtmlChartFormat();
            }

            return new PptHtmlChartFormat
            {
                ChartType = GetAttr(el, "data-chart-type"),
                Title = GetAttr(el, "data-title"),
                Theme = GetAttr(el, "data-theme"),
                Legend = GetAttr(el, "data-legend"),
                ShowDataLabels = GetAttr(el, "data-show-data-labels"),
                PlotColor = GetAttr(el, "data-plot-color"),
                TitleFontSize = GetAttr(el, "data-title-font-size"),
                TitleFontBold = GetAttr(el, "data-title-font-bold"),
                TitleFontColor = GetAttr(el, "data-title-font-color"),
                Gridlines = GetAttr(el, "data-gridlines"),
                GapWidth = GetAttr(el, "data-gap-width"),
                DataMarkers = GetAttr(el, "data-data-markers"),
                MarkerSize = GetAttr(el, "data-marker-size"),
                ChartLineWeight = GetAttr(el, "data-chart-line-weight"),
                AxisYMin = GetAttr(el, "data-axis-y-min"),
                AxisYMax = GetAttr(el, "data-axis-y-max"),
                AxisYMajorUnit = GetAttr(el, "data-axis-y-major-unit"),
                DataLabelType = GetAttr(el, "data-data-label-type"),
                Explosion = GetAttr(el, "data-explosion"),
                FillMissing = GetAttr(el, "data-fill-missing"),
                AxisX = GetAttr(el, "data-axis-x"),
                AxisXType = GetAttr(el, "data-axis-x-type"),
                AxisXFormat = GetAttr(el, "data-axis-x-format"),
                AxisXTickCount = GetAttr(el, "data-axis-x-tick-count"),
                AxisXTickSpacing = GetAttr(el, "data-axis-x-tick-spacing"),
                AxisXBetween = GetAttr(el, "data-axis-x-between"),
                AxisY = GetAttr(el, "data-axis-y"),
                AxisYSecondary = GetAttr(el, "data-axis-y-secondary")
            };
        }

        public static string BuildInnerHtml(PptHtmlChartGrid grid)
        {
            if (grid == null || grid.Columns == null)
            {
                return "";
            }

            var sb = new StringBuilder();
            sb.AppendLine("    <tr>");
            foreach (PptHtmlChartColumn col in grid.Columns)
            {
                sb.Append("      <th data-col=\"").Append(EscapeAttr(col.Role ?? "value")).Append("\"");
                if (!string.IsNullOrEmpty(col.Color))
                {
                    sb.Append(" data-color=\"").Append(EscapeAttr(col.Color)).Append("\"");
                }

                if (!string.IsNullOrEmpty(col.SeriesType))
                {
                    sb.Append(" data-series-type=\"").Append(EscapeAttr(col.SeriesType)).Append("\"");
                }

                if (!string.IsNullOrEmpty(col.AxisY))
                {
                    sb.Append(" data-axis-y=\"").Append(EscapeAttr(col.AxisY)).Append("\"");
                }

                if (!string.IsNullOrEmpty(col.ShowDataLabels))
                {
                    sb.Append(" data-show-data-labels=\"").Append(EscapeAttr(col.ShowDataLabels)).Append("\"");
                }

                sb.Append(">").Append(EscapeText(col.Name)).AppendLine("</th>");
            }

            sb.AppendLine("    </tr>");
            if (grid.Rows != null)
            {
                foreach (List<string> row in grid.Rows)
                {
                    sb.Append("    <tr>");
                    int n = grid.Columns.Count;
                    for (int i = 0; i < n; i++)
                    {
                        string cell = row != null && i < row.Count ? row[i] : "";
                        sb.Append("<td>").Append(EscapeText(cell)).Append("</td>");
                    }

                    sb.AppendLine("</tr>");
                }
            }

            return sb.ToString();
        }

        public static void AppendFormatAttrs(StringBuilder sb, PptHtmlChartFormat fmt)
        {
            if (sb == null || fmt == null)
            {
                return;
            }

            WriteAttr(sb, "data-chart-type", fmt.ChartType);
            WriteAttr(sb, "data-title", fmt.Title);
            WriteAttr(sb, "data-theme", fmt.Theme);
            WriteAttr(sb, "data-legend", fmt.Legend);
            WriteAttr(sb, "data-show-data-labels", fmt.ShowDataLabels);
            WriteAttr(sb, "data-plot-color", fmt.PlotColor);
            WriteAttr(sb, "data-title-font-size", fmt.TitleFontSize);
            WriteAttr(sb, "data-title-font-bold", fmt.TitleFontBold);
            WriteAttr(sb, "data-title-font-color", fmt.TitleFontColor);
            WriteAttr(sb, "data-gridlines", fmt.Gridlines);
            WriteAttr(sb, "data-gap-width", fmt.GapWidth);
            WriteAttr(sb, "data-data-markers", fmt.DataMarkers);
            WriteAttr(sb, "data-marker-size", fmt.MarkerSize);
            WriteAttr(sb, "data-chart-line-weight", fmt.ChartLineWeight);
            WriteAttr(sb, "data-axis-y-min", fmt.AxisYMin);
            WriteAttr(sb, "data-axis-y-max", fmt.AxisYMax);
            WriteAttr(sb, "data-axis-y-major-unit", fmt.AxisYMajorUnit);
            WriteAttr(sb, "data-data-label-type", fmt.DataLabelType);
            WriteAttr(sb, "data-explosion", fmt.Explosion);
            WriteAttr(sb, "data-fill-missing", fmt.FillMissing);
            WriteAttr(sb, "data-axis-x", fmt.AxisX);
            WriteAttr(sb, "data-axis-x-type", fmt.AxisXType);
            WriteAttr(sb, "data-axis-x-format", fmt.AxisXFormat);
            WriteAttr(sb, "data-axis-x-tick-count", fmt.AxisXTickCount);
            WriteAttr(sb, "data-axis-x-tick-spacing", fmt.AxisXTickSpacing);
            WriteAttr(sb, "data-axis-x-between", fmt.AxisXBetween);
            WriteAttr(sb, "data-axis-y", fmt.AxisY);
            WriteAttr(sb, "data-axis-y-secondary", fmt.AxisYSecondary);
        }

        public static bool LooksLikeChart(object shape)
        {
            if (shape == null)
            {
                return false;
            }

            try
            {
                return IsTruthy(WppCom.GetProperty(shape, "HasChart"));
            }
            catch (Exception)
            {
                return false;
            }
        }

        public static bool TryRead(object shape, out PptHtmlChartReadModel model, out string error)
        {
            model = null;
            error = null;
            object chart = TryGetChart(shape);
            if (chart == null)
            {
                error = "无法读取图表 Chart 对象";
                return false;
            }

            var format = new PptHtmlChartFormat();
            try
            {
                object t = WppCom.GetProperty(chart, "ChartType");
                if (t != null)
                {
                    format.ChartType = CanonicalTypeFromXl(Convert.ToInt32(t));
                }
            }
            catch (Exception)
            {
            }

            TryReadTitle(chart, format);
            TryReadLegend(chart, format);
            TryReadPlotAndLabels(chart, format);
            TryReadAxes(chart, format);

            if (!TryReadGrid(chart, out PptHtmlChartGrid grid, out error))
            {
                return false;
            }

            model = new PptHtmlChartReadModel
            {
                Format = format,
                Grid = grid
            };
            return true;
        }

        public static bool TryPourGrid(object chart, PptHtmlChartGrid grid, out string error)
        {
            error = null;
            if (chart == null || grid == null || !grid.IsPourable)
            {
                error = "新建 chart 必须在节点内嵌 <table> 灌数，不能建空图";
                return false;
            }

            object excelApp = null;
            try
            {
                int wantSeries = 0;
                for (int c = 0; c < grid.Columns.Count; c++)
                {
                    if (grid.Columns[c].Role != "category")
                    {
                        wantSeries++;
                    }
                }

                // 与 Word FillChartData 相同：直接写 Series，不 Activate ChartData（避免拉起内嵌 Excel）
                if (!TryEnsureSeriesCount(chart, wantSeries, out error))
                {
                    if (!TryExpandViaChartData(chart, grid, out excelApp, out error))
                    {
                        return false;
                    }

                    if (!TryEnsureSeriesCount(chart, wantSeries, out error))
                    {
                        return false;
                    }
                }

                if (!TryPourSeriesLikeWord(chart, grid, out error))
                {
                    return false;
                }

                TrimExtraSeries(chart, wantSeries, out _);
                TryInvoke(chart, "Refresh");
                ApplySeriesExtras(chart, grid);
                return true;
            }
            catch (Exception ex)
            {
                error = "灌入图表数据失败: " + ex.Message;
                return false;
            }
            finally
            {
                HideEmbeddedExcel(excelApp);
            }
        }

        public static bool TryApplyFormat(object chart, PptHtmlChartFormat fmt, List<string> warnings, out string error)
        {
            error = null;
            if (chart == null || fmt == null)
            {
                return true;
            }

            if (!string.IsNullOrWhiteSpace(fmt.ChartType))
            {
                if (!TryParseType(fmt.ChartType, out int xl, out string canon, out error))
                {
                    return false;
                }

                if (!IsSameChartType(chart, canon))
                {
                    TryCaptureTickLabelFont(chart, out object tickColor, out object tickSize);
                    try
                    {
                        WppCom.TrySetProperty(chart, "ChartType", xl);
                    }
                    catch (Exception ex)
                    {
                        error = "无法改 data-chart-type: " + ex.Message;
                        return false;
                    }

                    TryRestoreTickLabelFont(chart, tickColor, tickSize);
                }
            }

            if (fmt.Title != null)
            {
                TrySetTitle(chart, fmt.Title, warnings);
            }

            TrySetOptional(chart, fmt, warnings);
            return true;
        }

        /// <summary>
        /// 改已有图：先 AddChart 再建、再删旧图。新图自带包内 embeddings，不继承外链。
        /// </summary>
        public static bool TryReplaceOnSlide(
            object shapes,
            object oldShape,
            PptHtmlChartGrid grid,
            PptHtmlChartFormat format,
            float? left,
            float? top,
            float? width,
            float? height,
            List<string> warnings,
            out object newShape,
            out string error)
        {
            newShape = null;
            error = null;
            if (shapes == null || oldShape == null)
            {
                error = "无法定位要替换的图表";
                return false;
            }

            if (!TryReadBox(oldShape, out float oldLeft, out float oldTop, out float oldWidth, out float oldHeight))
            {
                error = "无法读取原图位置";
                return false;
            }

            float useLeft = left ?? oldLeft;
            float useTop = top ?? oldTop;
            float useWidth = width ?? oldWidth;
            float useHeight = height ?? oldHeight;

            PptHtmlChartFormat useFormat = format ?? new PptHtmlChartFormat();
            PptHtmlChartGrid useGrid = grid != null && grid.IsPourable ? grid : null;
            if (useGrid == null)
            {
                if (TryRead(oldShape, out PptHtmlChartReadModel model, out _)
                    && model != null
                    && model.Grid != null
                    && model.Grid.IsPourable)
                {
                    useGrid = model.Grid;
                    if (string.IsNullOrWhiteSpace(useFormat.ChartType)
                        && model.Format != null
                        && !string.IsNullOrWhiteSpace(model.Format.ChartType))
                    {
                        useFormat.ChartType = model.Format.ChartType;
                    }
                }
            }

            if (useGrid == null || !useGrid.IsPourable)
            {
                error = "改已有 chart 须带内嵌 <table>（将删旧图重建，避免继承外链）";
                return false;
            }

            object oldChart = TryGetChart(oldShape);
            if (string.IsNullOrWhiteSpace(useFormat.ChartType))
            {
                try
                {
                    object t = oldChart == null ? null : WppCom.GetProperty(oldChart, "ChartType");
                    if (t != null)
                    {
                        useFormat.ChartType = CanonicalTypeFromXl(Convert.ToInt32(t));
                    }
                }
                catch (Exception)
                {
                }
            }

            if (!TryParseType(useFormat.ChartType, out int xlType, out _, out error))
            {
                return false;
            }

            ChartStyleSnap snap = null;
            try
            {
                StyleLog(warnings, "重建开始 type=" + (useFormat.ChartType ?? "")
                    + " xl=" + xlType
                    + " box=" + useLeft.ToString("0.#", CultureInfo.InvariantCulture)
                    + "," + useTop.ToString("0.#", CultureInfo.InvariantCulture)
                    + " " + useWidth.ToString("0.#", CultureInfo.InvariantCulture)
                    + "x" + useHeight.ToString("0.#", CultureInfo.InvariantCulture));
                snap = TryCaptureStyle(oldChart, warnings, "旧图");
            }
            catch (Exception ex)
            {
                StyleLog(warnings, "拍旧图样式异常: " + ex.Message);
            }

            if (snap != null && snap.Series != null && snap.Series.Count == 1
                && snap.Series[0].ChartType.HasValue
                && snap.Series[0].ChartType.Value != xlType)
            {
                StyleLog(warnings, "建图类型改用快照 " + xlType + " → " + snap.Series[0].ChartType.Value);
                xlType = snap.Series[0].ChartType.Value;
            }

            if (!TryCreateOnSlide(
                shapes,
                useLeft,
                useTop,
                useWidth,
                useHeight,
                xlType,
                useGrid,
                useFormat,
                warnings,
                out newShape,
                out error,
                snap == null || !snap.ChartStyle.HasValue ? -1 : snap.ChartStyle.Value,
                newLayout: false))
            {
                return false;
            }

            object newChart = TryGetChart(newShape);
            TryApplyStyleSnap(newChart, snap, warnings);
            try
            {
                TryCaptureStyle(newChart, warnings, "新图套回后");
            }
            catch (Exception ex)
            {
                StyleLog(warnings, "回读新图异常: " + ex.Message);
            }

            TryDelete(oldShape);
            warnings?.Add("已删除原图并新建（不继承外链；已套回原图样式）");
            if (!string.IsNullOrEmpty(EasyWriteLog.CurrentLogPath))
            {
                warnings?.Add("chart 样式日志: " + EasyWriteLog.CurrentLogPath);
            }

            return true;
        }

        private static ChartStyleSnap TryCaptureStyle(object chart, List<string> warnings = null, string tag = null)
        {
            if (chart == null)
            {
                return null;
            }

            var snap = new ChartStyleSnap { Series = new List<SeriesStyleSnap>() };
            try
            {
                object style = WppCom.GetProperty(chart, "ChartStyle");
                if (style != null)
                {
                    snap.ChartStyle = Convert.ToInt32(style);
                }
            }
            catch (Exception)
            {
            }

            try
            {
                snap.HasTitle = IsTruthy(WppCom.GetProperty(chart, "HasTitle"));
                if (snap.HasTitle == true)
                {
                    object title = WppCom.GetProperty(chart, "ChartTitle");
                    object font = title == null ? null : WppCom.GetProperty(title, "Font");
                    snap.TitleFontColor = TryReadFontColorHex(font);
                    object sz = font == null ? null : WppCom.GetProperty(font, "Size");
                    if (sz != null)
                    {
                        snap.TitleFontSize = Convert.ToDouble(sz).ToString("0.##", CultureInfo.InvariantCulture);
                    }

                    object bold = font == null ? null : WppCom.GetProperty(font, "Bold");
                    if (bold != null)
                    {
                        snap.TitleFontBold = IsTruthy(bold);
                    }
                }
            }
            catch (Exception)
            {
            }

            try
            {
                snap.HasLegend = IsTruthy(WppCom.GetProperty(chart, "HasLegend"));
                if (snap.HasLegend == true)
                {
                    object legend = WppCom.GetProperty(chart, "Legend");
                    object pos = legend == null ? null : WppCom.GetProperty(legend, "Position");
                    if (pos != null)
                    {
                        snap.LegendPosition = Convert.ToInt32(pos);
                    }

                    object legendFont = legend == null ? null : WppCom.GetProperty(legend, "Font");
                    snap.LegendFontColor = TryReadFontColorHex(legendFont);
                }
            }
            catch (Exception)
            {
            }

            TryReadAreaFill(chart, "ChartArea", out bool? areaVis, out int? areaRgb);
            snap.ChartAreaFillVisible = areaVis;
            snap.ChartAreaFillRgb = areaRgb;
            TryReadAreaFill(chart, "PlotArea", out bool? plotVis, out int? plotRgb);
            snap.PlotFillVisible = plotVis;
            snap.PlotFillRgb = plotRgb;
            TryCapturePlotLayout(chart, snap);
            TryCaptureChartGroup(chart, snap);

            snap.Category = TryCaptureAxis(chart, XlCategory, XlPrimary);
            snap.Value = TryCaptureAxis(chart, XlValue, XlPrimary);
            snap.ValueSecondary = TryCaptureAxis(chart, XlValue, XlSecondary);

            try
            {
                object sc = TryInvoke(chart, "SeriesCollection");
                if (sc == null)
                {
                    sc = WppCom.GetProperty(chart, "SeriesCollection");
                }

                int n = Convert.ToInt32(WppCom.GetProperty(sc, "Count"));
                for (int i = 1; i <= n; i++)
                {
                    object series = GetSeries(chart, i);
                    if (series == null)
                    {
                        continue;
                    }

                    var one = new SeriesStyleSnap();
                    try
                    {
                        object t = WppCom.GetProperty(series, "ChartType");
                        if (t != null)
                        {
                            one.ChartType = Convert.ToInt32(t);
                        }
                    }
                    catch (Exception)
                    {
                    }

                    try
                    {
                        object g = WppCom.GetProperty(series, "AxisGroup");
                        if (g != null)
                        {
                            one.AxisGroup = Convert.ToInt32(g);
                        }
                    }
                    catch (Exception)
                    {
                    }

                    one.Fill = TryCaptureFill(series, warnings, (tag ?? "图") + " S" + i);
                    one.Line = TryCaptureLine(series, warnings, (tag ?? "图") + " S" + i);
                    TryCaptureMarker(series, one);
                    TryCaptureDataLabels(series, one);
                    snap.Series.Add(one);
                }
            }
            catch (Exception ex)
            {
                StyleLog(warnings, (tag ?? "图") + " 读系列失败: " + ex.Message);
            }

            StyleLog(warnings, DescribeSnap(tag ?? "图", snap));
            return snap;
        }

        private static void TryApplyStyleSnap(object chart, ChartStyleSnap snap, List<string> warnings)
        {
            if (chart == null || snap == null)
            {
                StyleLog(warnings, "套回跳过 chart=" + (chart == null ? "null" : "ok")
                    + " snap=" + (snap == null ? "null" : "ok"));
                return;
            }

            StyleLog(warnings, "开始套回 " + DescribeSnap("快照", snap));
            try
            {
                if (snap.Series != null)
                {
                    for (int i = 0; i < snap.Series.Count; i++)
                    {
                        object series = GetSeries(chart, i + 1);
                        if (series == null)
                        {
                            continue;
                        }

                        SeriesStyleSnap one = snap.Series[i];
                        if (one.ChartType.HasValue)
                        {
                            WppCom.TrySetProperty(series, "ChartType", one.ChartType.Value);
                        }

                        if (one.AxisGroup.HasValue)
                        {
                            WppCom.TrySetProperty(series, "AxisGroup", one.AxisGroup.Value);
                        }
                    }
                }

                if (snap.ChartStyle.HasValue)
                {
                    WppCom.TrySetProperty(chart, "ChartStyle", snap.ChartStyle.Value);
                }

                if (snap.HasTitle.HasValue)
                {
                    WppCom.TrySetProperty(chart, "HasTitle", snap.HasTitle.Value);
                    if (snap.HasTitle.Value)
                    {
                        object title = WppCom.GetProperty(chart, "ChartTitle");
                        object font = title == null ? null : WppCom.GetProperty(title, "Font");
                        if (font != null)
                        {
                            if (!string.IsNullOrEmpty(snap.TitleFontColor)
                                && TryParseHexToOffice(snap.TitleFontColor, out int tRgb))
                            {
                                WppCom.TrySetProperty(font, "Color", tRgb);
                            }

                            if (!string.IsNullOrEmpty(snap.TitleFontSize)
                                && double.TryParse(snap.TitleFontSize, NumberStyles.Float, CultureInfo.InvariantCulture, out double sz))
                            {
                                WppCom.TrySetProperty(font, "Size", sz);
                            }

                            if (snap.TitleFontBold.HasValue)
                            {
                                WppCom.TrySetProperty(font, "Bold", snap.TitleFontBold.Value);
                            }
                        }
                    }
                }

                if (snap.HasLegend.HasValue)
                {
                    WppCom.TrySetProperty(chart, "HasLegend", snap.HasLegend.Value);
                    if (snap.HasLegend.Value)
                    {
                        object legend = WppCom.GetProperty(chart, "Legend");
                        if (snap.LegendPosition.HasValue)
                        {
                            WppCom.TrySetProperty(legend, "Position", snap.LegendPosition.Value);
                        }

                        if (!string.IsNullOrEmpty(snap.LegendFontColor)
                            && TryParseHexToOffice(snap.LegendFontColor, out int lRgb))
                        {
                            object legendFont = legend == null ? null : WppCom.GetProperty(legend, "Font");
                            WppCom.TrySetProperty(legendFont, "Color", lRgb);
                        }
                    }
                }

                TryWriteAreaFill(chart, "ChartArea", snap.ChartAreaFillVisible, snap.ChartAreaFillRgb);
                TryWriteAreaFill(chart, "PlotArea", snap.PlotFillVisible, snap.PlotFillRgb);
                TryApplyChartGroup(chart, snap);
                TryApplyAxis(chart, XlCategory, XlPrimary, snap.Category);
                TryApplyAxis(chart, XlValue, XlPrimary, snap.Value);
                TryApplyAxis(chart, XlValue, XlSecondary, snap.ValueSecondary);

                if (snap.Series != null)
                {
                    for (int i = 0; i < snap.Series.Count; i++)
                    {
                        object series = GetSeries(chart, i + 1);
                        if (series == null)
                        {
                            continue;
                        }

                        SeriesStyleSnap one = snap.Series[i];
                        TryApplyFill(series, one.Fill, warnings, "S" + (i + 1));
                        TryApplyLine(series, one.Line, warnings, "S" + (i + 1));
                        TryApplyMarker(series, one);
                    }
                }

                TryApplyPlotLayout(chart, snap);

                if (snap.Series != null)
                {
                    for (int i = 0; i < snap.Series.Count; i++)
                    {
                        object series = GetSeries(chart, i + 1);
                        if (series == null)
                        {
                            continue;
                        }

                        TryApplyDataLabels(series, snap.Series[i], warnings, "S" + (i + 1));
                    }
                }
            }
            catch (Exception ex)
            {
                StyleLog(warnings, "套回原图样式部分失败: " + ex.Message);
            }
        }

        private static AxisStyleSnap TryCaptureAxis(object chart, int axisType, int group)
        {
            bool? hasAxis = TryReadHasAxis(chart, axisType, group);
            if (hasAxis == false)
            {
                return new AxisStyleSnap { Deleted = true };
            }

            try
            {
                object axis = TryInvoke(chart, "Axes", axisType, group);
                if (axis == null)
                {
                    return hasAxis == true ? null : new AxisStyleSnap { Deleted = true };
                }

                var snap = new AxisStyleSnap { Deleted = false };
                try
                {
                    snap.HasTitle = IsTruthy(WppCom.GetProperty(axis, "HasTitle"));
                    if (snap.HasTitle == true)
                    {
                        object t = WppCom.GetProperty(axis, "AxisTitle");
                        snap.Title = Convert.ToString(WppCom.GetProperty(t, "Text") ?? "");
                    }
                }
                catch (Exception)
                {
                }

                try
                {
                    object ticks = WppCom.GetProperty(axis, "TickLabels");
                    object font = ticks == null ? null : WppCom.GetProperty(ticks, "Font");
                    string hex = TryReadFontColorHex(font);
                    if (!string.IsNullOrEmpty(hex) && TryParseHexToOffice(hex, out int rgb))
                    {
                        snap.TickFontColor = rgb;
                    }

                    object sz = font == null ? null : WppCom.GetProperty(font, "Size");
                    if (sz != null)
                    {
                        snap.TickFontSize = Convert.ToDouble(sz);
                    }

                    object name = font == null ? null : WppCom.GetProperty(font, "Name");
                    if (name != null)
                    {
                        snap.TickFontName = Convert.ToString(name);
                    }

                    object fmt = ticks == null ? null : WppCom.GetProperty(ticks, "NumberFormat");
                    if (fmt != null)
                    {
                        snap.NumberFormat = Convert.ToString(fmt);
                    }
                }
                catch (Exception)
                {
                }

                try
                {
                    object pos = WppCom.GetProperty(axis, "TickLabelPosition");
                    if (pos != null)
                    {
                        snap.TickLabelPosition = Convert.ToInt32(pos);
                    }

                    object major = WppCom.GetProperty(axis, "MajorTickMark");
                    if (major != null)
                    {
                        snap.MajorTickMark = Convert.ToInt32(major);
                    }

                    object minor = WppCom.GetProperty(axis, "MinorTickMark");
                    if (minor != null)
                    {
                        snap.MinorTickMark = Convert.ToInt32(minor);
                    }
                }
                catch (Exception)
                {
                }

                try
                {
                    snap.HasMajorGridlines = IsTruthy(WppCom.GetProperty(axis, "HasMajorGridlines"));
                    if (snap.HasMajorGridlines == true)
                    {
                        object gl = WppCom.GetProperty(axis, "MajorGridlines");
                        object glLine = gl == null ? null : WppCom.GetProperty(WppCom.GetProperty(gl, "Format"), "Line");
                        object glColor = glLine == null ? null : WppCom.GetProperty(glLine, "ForeColor");
                        snap.MajorGridlineRgb = TryReadExplicitRgb(glColor);
                    }
                }
                catch (Exception)
                {
                }

                try
                {
                    object axisLine = WppCom.GetProperty(WppCom.GetProperty(axis, "Format"), "Line");
                    object vis = axisLine == null ? null : WppCom.GetProperty(axisLine, "Visible");
                    if (vis != null)
                    {
                        snap.LineVisible = Convert.ToInt32(vis) != 0;
                    }

                    object axisColor = axisLine == null ? null : WppCom.GetProperty(axisLine, "ForeColor");
                    snap.LineRgb = TryReadExplicitRgb(axisColor);
                    object weight = axisLine == null ? null : WppCom.GetProperty(axisLine, "Weight");
                    if (weight != null)
                    {
                        snap.LineWeight = Convert.ToDouble(weight);
                    }
                }
                catch (Exception)
                {
                }

                if (hasAxis == null
                    && snap.TickLabelPosition == XlTickLabelPositionNone
                    && snap.LineVisible == false
                    && snap.HasTitle != true)
                {
                    snap.Deleted = true;
                }

                return snap;
            }
            catch (Exception)
            {
                return hasAxis == false ? new AxisStyleSnap { Deleted = true } : null;
            }
        }

        private static void TryApplyAxis(object chart, int axisType, int group, AxisStyleSnap snap)
        {
            if (snap == null)
            {
                return;
            }

            if (snap.Deleted == true)
            {
                TryHideAxis(chart, axisType, group);
                return;
            }

            try
            {
                object axis = TryInvoke(chart, "Axes", axisType, group);
                if (axis == null)
                {
                    return;
                }

                if (snap.HasTitle.HasValue)
                {
                    WppCom.TrySetProperty(axis, "HasTitle", snap.HasTitle.Value);
                    if (snap.HasTitle.Value && !string.IsNullOrEmpty(snap.Title))
                    {
                        object at = WppCom.GetProperty(axis, "AxisTitle");
                        WppCom.TrySetProperty(at, "Text", snap.Title);
                    }
                }

                if (snap.TickLabelPosition.HasValue)
                {
                    WppCom.TrySetProperty(axis, "TickLabelPosition", snap.TickLabelPosition.Value);
                }

                if (snap.MajorTickMark.HasValue)
                {
                    WppCom.TrySetProperty(axis, "MajorTickMark", snap.MajorTickMark.Value);
                }

                if (snap.MinorTickMark.HasValue)
                {
                    WppCom.TrySetProperty(axis, "MinorTickMark", snap.MinorTickMark.Value);
                }

                object ticks = WppCom.GetProperty(axis, "TickLabels");
                object font = ticks == null ? null : WppCom.GetProperty(ticks, "Font");
                if (font != null)
                {
                    if (!string.IsNullOrEmpty(snap.TickFontName))
                    {
                        WppCom.TrySetProperty(font, "Name", snap.TickFontName);
                    }

                    if (snap.TickFontColor.HasValue)
                    {
                        WppCom.TrySetProperty(font, "Color", snap.TickFontColor.Value);
                    }

                    if (snap.TickFontSize.HasValue)
                    {
                        WppCom.TrySetProperty(font, "Size", snap.TickFontSize.Value);
                    }
                }

                if (!string.IsNullOrEmpty(snap.NumberFormat) && ticks != null)
                {
                    WppCom.TrySetProperty(ticks, "NumberFormat", snap.NumberFormat);
                }

                if (snap.HasMajorGridlines.HasValue)
                {
                    WppCom.TrySetProperty(axis, "HasMajorGridlines", snap.HasMajorGridlines.Value);
                    if (snap.HasMajorGridlines.Value && snap.MajorGridlineRgb.HasValue)
                    {
                        object gl = WppCom.GetProperty(axis, "MajorGridlines");
                        object glLine = gl == null ? null : WppCom.GetProperty(WppCom.GetProperty(gl, "Format"), "Line");
                        object glColor = glLine == null ? null : WppCom.GetProperty(glLine, "ForeColor");
                        WppCom.TrySetProperty(glColor, "RGB", snap.MajorGridlineRgb.Value);
                    }
                }

                object axisLine = WppCom.GetProperty(WppCom.GetProperty(axis, "Format"), "Line");
                if (snap.LineVisible == false)
                {
                    WppCom.TrySetProperty(axisLine, "Visible", 0);
                }
                else
                {
                    if (snap.LineRgb.HasValue)
                    {
                        object axisColor = axisLine == null ? null : WppCom.GetProperty(axisLine, "ForeColor");
                        WppCom.TrySetProperty(axisColor, "RGB", snap.LineRgb.Value);
                    }

                    if (snap.LineWeight.HasValue)
                    {
                        WppCom.TrySetProperty(axisLine, "Weight", snap.LineWeight.Value);
                    }
                }
            }
            catch (Exception)
            {
            }
        }

        private static void TryReadAreaFill(object chart, string areaName, out bool? visible, out int? rgb)
        {
            visible = null;
            rgb = null;
            try
            {
                object area = WppCom.GetProperty(chart, areaName);
                object fill = WppCom.GetProperty(WppCom.GetProperty(area, "Format"), "Fill");
                object vis = WppCom.GetProperty(fill, "Visible");
                if (vis != null)
                {
                    visible = Convert.ToInt32(vis) != 0;
                }

                if (visible == true)
                {
                    object fc = WppCom.GetProperty(fill, "ForeColor");
                    rgb = TryReadExplicitRgb(fc);
                }
            }
            catch (Exception)
            {
            }
        }

        private static void TryWriteAreaFill(object chart, string areaName, bool? visible, int? rgb)
        {
            if (!visible.HasValue)
            {
                return;
            }

            try
            {
                object area = WppCom.GetProperty(chart, areaName);
                object fill = WppCom.GetProperty(WppCom.GetProperty(area, "Format"), "Fill");
                if (!visible.Value)
                {
                    WppCom.TrySetProperty(fill, "Visible", 0);
                    return;
                }

                WppCom.TrySetProperty(fill, "Visible", -1);
                if (rgb.HasValue)
                {
                    TryInvoke(fill, "Solid");
                    object fc = WppCom.GetProperty(fill, "ForeColor");
                    WppCom.TrySetProperty(fc, "RGB", rgb.Value);
                }
            }
            catch (Exception)
            {
            }
        }

        private static void TryCapturePlotLayout(object chart, ChartStyleSnap snap)
        {
            try
            {
                object plot = WppCom.GetProperty(chart, "PlotArea");
                if (plot == null)
                {
                    return;
                }

                snap.PlotLeft = TryReadFloat(plot, "Left");
                snap.PlotTop = TryReadFloat(plot, "Top");
                snap.PlotWidth = TryReadFloat(plot, "Width");
                snap.PlotHeight = TryReadFloat(plot, "Height");
                snap.PlotInsideLeft = TryReadFloat(plot, "InsideLeft");
                snap.PlotInsideTop = TryReadFloat(plot, "InsideTop");
                snap.PlotInsideWidth = TryReadFloat(plot, "InsideWidth");
                snap.PlotInsideHeight = TryReadFloat(plot, "InsideHeight");
            }
            catch (Exception)
            {
            }
        }

        private static void TryApplyPlotLayout(object chart, ChartStyleSnap snap)
        {
            try
            {
                object plot = WppCom.GetProperty(chart, "PlotArea");
                if (plot == null)
                {
                    return;
                }

                if (snap.PlotLeft.HasValue)
                {
                    WppCom.TrySetProperty(plot, "Left", snap.PlotLeft.Value);
                }

                if (snap.PlotTop.HasValue)
                {
                    WppCom.TrySetProperty(plot, "Top", snap.PlotTop.Value);
                }

                if (snap.PlotWidth.HasValue && snap.PlotWidth.Value > 0)
                {
                    WppCom.TrySetProperty(plot, "Width", snap.PlotWidth.Value);
                }

                if (snap.PlotHeight.HasValue && snap.PlotHeight.Value > 0)
                {
                    WppCom.TrySetProperty(plot, "Height", snap.PlotHeight.Value);
                }

                if (snap.PlotInsideLeft.HasValue)
                {
                    WppCom.TrySetProperty(plot, "InsideLeft", snap.PlotInsideLeft.Value);
                }

                if (snap.PlotInsideTop.HasValue)
                {
                    WppCom.TrySetProperty(plot, "InsideTop", snap.PlotInsideTop.Value);
                }

                if (snap.PlotInsideWidth.HasValue && snap.PlotInsideWidth.Value > 0)
                {
                    WppCom.TrySetProperty(plot, "InsideWidth", snap.PlotInsideWidth.Value);
                }

                if (snap.PlotInsideHeight.HasValue && snap.PlotInsideHeight.Value > 0)
                {
                    WppCom.TrySetProperty(plot, "InsideHeight", snap.PlotInsideHeight.Value);
                }
            }
            catch (Exception)
            {
            }
        }

        private static void TryCaptureChartGroup(object chart, ChartStyleSnap snap)
        {
            try
            {
                object g = TryInvoke(chart, "ChartGroups", 1);
                if (g == null)
                {
                    return;
                }

                object gap = WppCom.GetProperty(g, "GapWidth");
                if (gap != null)
                {
                    snap.GapWidth = Convert.ToInt32(gap);
                }

                object overlap = WppCom.GetProperty(g, "Overlap");
                if (overlap != null)
                {
                    snap.Overlap = Convert.ToInt32(overlap);
                }
            }
            catch (Exception)
            {
            }
        }

        private static void TryApplyChartGroup(object chart, ChartStyleSnap snap)
        {
            try
            {
                object g = TryInvoke(chart, "ChartGroups", 1);
                if (g == null)
                {
                    return;
                }

                if (snap.GapWidth.HasValue)
                {
                    WppCom.TrySetProperty(g, "GapWidth", snap.GapWidth.Value);
                }

                if (snap.Overlap.HasValue)
                {
                    WppCom.TrySetProperty(g, "Overlap", snap.Overlap.Value);
                }
            }
            catch (Exception)
            {
            }
        }

        private static FillSnap TryCaptureFill(object series, List<string> warnings = null, string tag = null)
        {
            string prefix = (tag ?? "fill");
            try
            {
                object fill = WppCom.GetProperty(WppCom.GetProperty(series, "Format"), "Fill");
                if (fill == null)
                {
                    StyleLog(warnings, prefix + " Format.Fill=null");
                    return null;
                }

                var snap = new FillSnap();
                object vis = WppCom.GetProperty(fill, "Visible");
                if (vis != null)
                {
                    snap.Visible = Convert.ToInt32(vis) != 0;
                }

                object fillType = WppCom.GetProperty(fill, "Type");
                if (fillType != null)
                {
                    snap.FillType = Convert.ToInt32(fillType);
                }

                object angle = WppCom.GetProperty(fill, "GradientAngle");
                if (angle != null)
                {
                    snap.Angle = Convert.ToDouble(angle);
                }

                snap.Stops = TryReadGradientStops(fill, warnings, prefix);
                if (snap.Stops == null || snap.Stops.Count < 2)
                {
                    snap.SolidRgb = TryReadResolvedRgb(WppCom.GetProperty(fill, "ForeColor"));
                }
                else
                {
                    snap.SolidRgb = PickSolidFromStops(snap.Stops);
                }

                StyleLog(warnings, prefix + " Fill.Visible=" + vis
                    + " Type=" + fillType
                    + " Angle=" + angle
                    + " stops=" + DescribeStops(snap.Stops)
                    + " solid=" + HexOf(snap.SolidRgb));
                return snap;
            }
            catch (Exception ex)
            {
                StyleLog(warnings, prefix + " 拍填充失败: " + ex.Message);
                return null;
            }
        }

        private static void TryApplyFill(object series, FillSnap snap, List<string> warnings = null, string tag = null)
        {
            string prefix = "套填充 " + (tag ?? "");
            if (snap == null)
            {
                return;
            }

            try
            {
                object fill = WppCom.GetProperty(WppCom.GetProperty(series, "Format"), "Fill");
                if (fill == null)
                {
                    StyleLog(warnings, prefix + " Format.Fill=null");
                    return;
                }

                if (snap.Visible == false && (snap.Stops == null || snap.Stops.Count < 2))
                {
                    WppCom.TrySetProperty(fill, "Visible", 0);
                    StyleLog(warnings, prefix + " 隐藏填充");
                    return;
                }

                WppCom.TrySetProperty(fill, "Visible", -1);
                if (snap.Stops != null && snap.Stops.Count >= 2)
                {
                    bool ok = TryWriteGradientStops(fill, snap.Stops, snap.Angle);
                    StyleLog(warnings, prefix + " 渐变 " + DescribeStops(snap.Stops) + " ok=" + ok);
                    if (ok)
                    {
                        return;
                    }
                }

                if (snap.SolidRgb.HasValue)
                {
                    TryInvoke(fill, "Solid");
                    object fc = WppCom.GetProperty(fill, "ForeColor");
                    WppCom.TrySetProperty(fc, "RGB", snap.SolidRgb.Value);
                    StyleLog(warnings, prefix + " 实色 " + HexOf(snap.SolidRgb));
                }
            }
            catch (Exception ex)
            {
                StyleLog(warnings, prefix + " 异常: " + ex.Message);
            }
        }

        private static LineSnap TryCaptureLine(object series, List<string> warnings = null, string tag = null)
        {
            string prefix = tag ?? "line";
            try
            {
                object line = WppCom.GetProperty(WppCom.GetProperty(series, "Format"), "Line");
                if (line == null)
                {
                    StyleLog(warnings, prefix + " Format.Line=null");
                    return null;
                }

                var snap = new LineSnap();
                object vis = WppCom.GetProperty(line, "Visible");
                if (vis != null)
                {
                    snap.Visible = Convert.ToInt32(vis) != 0;
                }

                object weight = WppCom.GetProperty(line, "Weight");
                if (weight != null)
                {
                    snap.Weight = Convert.ToDouble(weight);
                }

                object fc = WppCom.GetProperty(line, "ForeColor");
                object colorType = fc == null ? null : WppCom.GetProperty(fc, "Type");
                object rawRgb = fc == null ? null : WppCom.GetProperty(fc, "RGB");
                StyleLog(warnings, prefix + " Line.Visible=" + vis
                    + " Weight=" + weight
                    + " ForeColor.Type=" + colorType
                    + " ForeColor.RGB=" + rawRgb);

                if (IsPhantomStroke(snap.Weight, colorType, rawRgb))
                {
                    snap.Visible = false;
                    snap.Rgb = null;
                    snap.Stops = null;
                    StyleLog(warnings, prefix + " 原图无线条（noFill/自动线），不描边");
                    return snap;
                }

                snap.Stops = TryReadGradientStops(line, warnings, prefix + " Line");
                if (snap.Stops == null || snap.Stops.Count < 2)
                {
                    object lineFill = null;
                    try
                    {
                        lineFill = WppCom.GetProperty(line, "Fill");
                    }
                    catch (Exception ex)
                    {
                        StyleLog(warnings, prefix + " Line.Fill 不可用: " + ex.Message);
                    }

                    if (lineFill != null)
                    {
                        snap.Stops = TryReadGradientStops(lineFill, warnings, prefix + " Line.Fill");
                    }
                }

                if (snap.Stops != null && snap.Stops.Count >= 2)
                {
                    snap.Rgb = PickSolidFromStops(snap.Stops);
                }
                else
                {
                    snap.Rgb = TryReadResolvedRgb(fc);
                }

                return snap;
            }
            catch (Exception ex)
            {
                StyleLog(warnings, prefix + " 拍线失败: " + ex.Message);
                return null;
            }
        }

        private static void TryApplyLine(object series, LineSnap snap, List<string> warnings = null, string tag = null)
        {
            string prefix = "套线 " + (tag ?? "");
            if (snap == null)
            {
                StyleLog(warnings, prefix + " snap=null 跳过");
                return;
            }

            try
            {
                object line = WppCom.GetProperty(WppCom.GetProperty(series, "Format"), "Line");
                if (line == null)
                {
                    StyleLog(warnings, prefix + " Format.Line=null");
                    return;
                }

                if (snap.Visible == false || IsPhantomStroke(snap.Weight, null, snap.Rgb))
                {
                    WppCom.TrySetProperty(line, "Visible", 0);
                    try
                    {
                        object border = WppCom.GetProperty(series, "Border");
                        WppCom.TrySetProperty(border, "LineStyle", -4142);
                    }
                    catch (Exception)
                    {
                    }

                    StyleLog(warnings, prefix + " 按原图不描边");
                    return;
                }

                WppCom.TrySetProperty(line, "Visible", -1);
                if (snap.Weight.HasValue && snap.Weight.Value > 0 && snap.Weight.Value < 50)
                {
                    WppCom.TrySetProperty(line, "Weight", snap.Weight.Value);
                }

                object lineFill = null;
                try
                {
                    lineFill = WppCom.GetProperty(line, "Fill");
                }
                catch (Exception)
                {
                }

                bool wroteGrad = false;
                if (snap.Stops != null && snap.Stops.Count >= 2)
                {
                    wroteGrad = TryWriteGradientStops(line, snap.Stops);
                    if (!wroteGrad && lineFill != null)
                    {
                        wroteGrad = TryWriteGradientStops(lineFill, snap.Stops);
                    }
                }

                if (snap.Rgb.HasValue)
                {
                    TryWriteLineRgb(line, snap.Rgb.Value);
                }

                StyleLog(warnings, prefix + " vis=" + snap.Visible
                    + " rgb=" + HexOf(snap.Rgb)
                    + " stops=" + (snap.Stops == null ? 0 : snap.Stops.Count)
                    + " grad=" + wroteGrad);
            }
            catch (Exception ex)
            {
                StyleLog(warnings, prefix + " 异常: " + ex.Message);
            }
        }

        private static void TryCaptureMarker(object series, SeriesStyleSnap one)
        {
            try
            {
                object style = WppCom.GetProperty(series, "MarkerStyle");
                if (style != null)
                {
                    one.MarkerStyle = Convert.ToInt32(style);
                }

                object size = WppCom.GetProperty(series, "MarkerSize");
                if (size != null)
                {
                    one.MarkerSize = Convert.ToInt32(size);
                }

                one.MarkerForeRgb = TryReadMarkerColor(WppCom.GetProperty(series, "MarkerForegroundColor"));
                one.MarkerBackRgb = TryReadMarkerColor(WppCom.GetProperty(series, "MarkerBackgroundColor"));
                if (!one.MarkerForeRgb.HasValue || !one.MarkerBackRgb.HasValue)
                {
                    int? lineRgb = one.Line == null ? null : one.Line.Rgb;
                    if (lineRgb.HasValue)
                    {
                        one.MarkerForeRgb = one.MarkerForeRgb ?? lineRgb;
                        one.MarkerBackRgb = one.MarkerBackRgb ?? lineRgb;
                    }
                }
            }
            catch (Exception)
            {
            }
        }

        private static void TryApplyMarker(object series, SeriesStyleSnap one)
        {
            if (one == null)
            {
                return;
            }

            try
            {
                if (one.MarkerStyle.HasValue)
                {
                    WppCom.TrySetProperty(series, "MarkerStyle", one.MarkerStyle.Value);
                }

                if (one.MarkerSize.HasValue)
                {
                    WppCom.TrySetProperty(series, "MarkerSize", one.MarkerSize.Value);
                }

                if (one.MarkerForeRgb.HasValue)
                {
                    WppCom.TrySetProperty(series, "MarkerForegroundColor", one.MarkerForeRgb.Value);
                }

                if (one.MarkerBackRgb.HasValue)
                {
                    WppCom.TrySetProperty(series, "MarkerBackgroundColor", one.MarkerBackRgb.Value);
                }
            }
            catch (Exception)
            {
            }
        }

        private static void TryCaptureDataLabels(object series, SeriesStyleSnap one)
        {
            try
            {
                object has = WppCom.GetProperty(series, "HasDataLabels");
                StyleLog(null, "HasDataLabels raw=" + (has == null ? "null" : has.ToString() + "/" + has.GetType().Name));
                if (has != null)
                {
                    one.HasDataLabels = IsTruthy(has);
                }
            }
            catch (Exception)
            {
            }

            if (one.HasDataLabels != true && IsLineLike(one.ChartType))
            {
                bool? onPoint = TryAnyPointHasLabel(series);
                if (onPoint == true)
                {
                    one.HasDataLabels = true;
                }
            }

            object dls = one.HasDataLabels == true ? TryGetDataLabels(series) : null;
            if (dls != null)
            {
                try
                {
                    object pos = WppCom.GetProperty(dls, "Position");
                    if (pos != null)
                    {
                        one.DataLabelPosition = Convert.ToInt32(pos);
                    }
                }
                catch (Exception)
                {
                }

                try
                {
                    object font = WppCom.GetProperty(dls, "Font");
                    if (font != null)
                    {
                        object name = WppCom.GetProperty(font, "Name");
                        if (name != null)
                        {
                            one.DataLabelFontName = Convert.ToString(name);
                        }

                        object sz = WppCom.GetProperty(font, "Size");
                        if (sz != null)
                        {
                            one.DataLabelFontSize = Convert.ToDouble(sz);
                        }

                        string hex = TryReadFontColorHex(font);
                        if (!string.IsNullOrEmpty(hex) && TryParseHexToOffice(hex, out int rgb))
                        {
                            one.DataLabelFontColor = rgb;
                        }
                    }
                }
                catch (Exception)
                {
                }

                try
                {
                    object fmt = WppCom.GetProperty(dls, "NumberFormat");
                    if (fmt != null)
                    {
                        one.DataLabelNumberFormat = Convert.ToString(fmt);
                    }
                }
                catch (Exception)
                {
                }
            }

        }

        private static void TryApplyDataLabels(object series, SeriesStyleSnap one, List<string> warnings = null, string tag = null)
        {
            string prefix = "套标签 " + (tag ?? "");
            if (one == null)
            {
                return;
            }

            bool want = one.HasDataLabels == true
                || one.DataLabelPosition.HasValue
                || one.DataLabelFontColor.HasValue
                || !string.IsNullOrEmpty(one.DataLabelNumberFormat);
            StyleLog(warnings, prefix + " HasDataLabels=" + one.HasDataLabels
                + " want=" + want
                + " pos=" + one.DataLabelPosition
                + " fontColor=" + HexOf(one.DataLabelFontColor)
                + " fmt=" + (one.DataLabelNumberFormat ?? ""));
            if (!want)
            {
                if (one.HasDataLabels == false)
                {
                    WppCom.TrySetProperty(series, "HasDataLabels", false);
                    StyleLog(warnings, prefix + " 按快照关闭标签");
                }

                return;
            }

            try
            {
                if (TryInvoke(series, "ApplyDataLabels", XlDataLabelsShowValue) == null)
                {
                    WppCom.TrySetProperty(series, "HasDataLabels", true);
                    StyleLog(warnings, prefix + " ApplyDataLabels 返回 null，改设 HasDataLabels");
                }
                else
                {
                    StyleLog(warnings, prefix + " ApplyDataLabels(2) 已调用");
                }
            }
            catch (Exception ex)
            {
                StyleLog(warnings, prefix + " ApplyDataLabels 异常: " + ex.Message);
                WppCom.TrySetProperty(series, "HasDataLabels", true);
            }

            object dls = TryGetDataLabels(series);
            if (dls == null)
            {
                return;
            }

            WppCom.TrySetProperty(dls, "ShowValue", true);
            WppCom.TrySetProperty(dls, "ShowCategoryName", false);
            WppCom.TrySetProperty(dls, "ShowSeriesName", false);
            WppCom.TrySetProperty(dls, "ShowPercentage", false);
            WppCom.TrySetProperty(dls, "ShowLegendKey", false);

            if (one.DataLabelPosition.HasValue)
            {
                WppCom.TrySetProperty(dls, "Position", one.DataLabelPosition.Value);
            }
            else if (IsLineLike(one.ChartType))
            {
                WppCom.TrySetProperty(dls, "Position", 0);
            }

            try
            {
                object font = WppCom.GetProperty(dls, "Font");
                if (font != null)
                {
                    if (!string.IsNullOrEmpty(one.DataLabelFontName))
                    {
                        WppCom.TrySetProperty(font, "Name", one.DataLabelFontName);
                    }

                    if (one.DataLabelFontSize.HasValue)
                    {
                        WppCom.TrySetProperty(font, "Size", one.DataLabelFontSize.Value);
                    }

                    if (one.DataLabelFontColor.HasValue)
                    {
                        WppCom.TrySetProperty(font, "Color", one.DataLabelFontColor.Value);
                    }
                    else if (IsLineLike(one.ChartType))
                    {
                        WppCom.TrySetProperty(font, "Color", 0x00FFFFFF);
                    }
                }
            }
            catch (Exception)
            {
            }

            if (!string.IsNullOrEmpty(one.DataLabelNumberFormat)
                && one.DataLabelNumberFormat.IndexOf(';') < 0)
            {
                WppCom.TrySetProperty(dls, "NumberFormatLinked", false);
                WppCom.TrySetProperty(dls, "NumberFormat", one.DataLabelNumberFormat);
            }
        }

        private static object TryGetDataLabels(object series)
        {
            object dls = TryInvoke(series, "DataLabels");
            if (dls != null)
            {
                return dls;
            }

            try
            {
                return WppCom.GetProperty(series, "DataLabels");
            }
            catch (Exception)
            {
                return null;
            }
        }

        private static bool? TryAnyPointHasLabel(object series)
        {
            try
            {
                object pts = TryInvoke(series, "Points");
                if (pts == null)
                {
                    pts = WppCom.GetProperty(series, "Points");
                }

                int n = Convert.ToInt32(WppCom.GetProperty(pts, "Count"));
                int take = Math.Min(n, 8);
                for (int i = 1; i <= take; i++)
                {
                    object pt = WppCom.GetIndexed(pts, i);
                    object has = pt == null ? null : WppCom.GetProperty(pt, "HasDataLabel");
                    if (IsTruthy(has))
                    {
                        return true;
                    }
                }
            }
            catch (Exception)
            {
            }

            return null;
        }

        private static bool IsLineLike(int? chartType)
        {
            if (!chartType.HasValue)
            {
                return false;
            }

            int t = chartType.Value;
            return t == XlLine || t == XlLineMarkers || t == 63 || t == 64 || t == 66;
        }

        private static List<GradientStopSnap> TryReadGradientStops(object fill, List<string> warnings = null, string tag = null)
        {
            if (fill == null)
            {
                return null;
            }

            try
            {
                object type = WppCom.GetProperty(fill, "Type");
                object gs = WppCom.GetProperty(fill, "GradientStops");
                if (gs == null)
                {
                    StyleLog(warnings, (tag ?? "grad") + " Type=" + type + " GradientStops=null");
                    return null;
                }

                int n = Convert.ToInt32(WppCom.GetProperty(gs, "Count"));
                var list = new List<GradientStopSnap>();
                for (int i = 1; i <= n; i++)
                {
                    object stop = WppCom.GetIndexed(gs, i);
                    if (stop == null)
                    {
                        continue;
                    }

                    var one = new GradientStopSnap();
                    object pos = WppCom.GetProperty(stop, "Position");
                    if (pos != null)
                    {
                        one.Position = Convert.ToDouble(pos);
                    }

                    object trans = WppCom.GetProperty(stop, "Transparency");
                    if (trans != null)
                    {
                        one.Transparency = Convert.ToDouble(trans);
                    }

                    object color = WppCom.GetProperty(stop, "Color");
                    int? rgb = TryReadResolvedRgb(color);
                    if (!rgb.HasValue)
                    {
                        rgb = TryReadOleColor(color == null ? null : WppCom.GetProperty(color, "RGB"));
                    }

                    if (!rgb.HasValue)
                    {
                        continue;
                    }

                    one.Rgb = rgb.Value;
                    list.Add(one);
                }

                StyleLog(warnings, (tag ?? "grad") + " Type=" + type + " n=" + n + " " + DescribeStops(list));
                return list.Count >= 2 ? list : null;
            }
            catch (Exception ex)
            {
                StyleLog(warnings, (tag ?? "grad") + " 读停靠点失败: " + ex.Message);
                return null;
            }
        }

        private static bool TryWriteGradientStops(object fill, List<GradientStopSnap> stops, double? angle = null)
        {
            if (fill == null || stops == null || stops.Count < 2)
            {
                return false;
            }

            try
            {
                TryInvoke(fill, "TwoColorGradient", 1, 1);
                if (angle.HasValue)
                {
                    WppCom.TrySetProperty(fill, "GradientAngle", angle.Value);
                }

                object gs = WppCom.GetProperty(fill, "GradientStops");
                if (gs == null)
                {
                    return false;
                }

                int count = Convert.ToInt32(WppCom.GetProperty(gs, "Count"));
                int use = Math.Min(count, stops.Count);
                for (int i = 0; i < use; i++)
                {
                    object stop = WppCom.GetIndexed(gs, i + 1);
                    WppCom.TrySetProperty(stop, "Position", stops[i].Position);
                    object color = WppCom.GetProperty(stop, "Color");
                    WppCom.TrySetProperty(color, "RGB", stops[i].Rgb);
                    WppCom.TrySetProperty(stop, "Transparency", stops[i].Transparency);
                }

                for (int i = count; i < stops.Count; i++)
                {
                    TryInvoke(gs, "Insert", stops[i].Rgb, stops[i].Position, stops[i].Transparency);
                }

                int after = Convert.ToInt32(WppCom.GetProperty(gs, "Count"));
                for (int i = after; i > stops.Count; i--)
                {
                    object extra = WppCom.GetIndexed(gs, i);
                    TryInvoke(extra, "Delete");
                }

                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        private static bool? TryReadHasAxis(object chart, int axisType, int group)
        {
            try
            {
                object v = chart.GetType().InvokeMember(
                    "HasAxis",
                    BindingFlags.GetProperty | BindingFlags.Instance | BindingFlags.Public,
                    null,
                    chart,
                    new object[] { axisType, group });
                if (v != null)
                {
                    return IsTruthy(v);
                }
            }
            catch (Exception)
            {
            }

            try
            {
                return TryInvoke(chart, "Axes", axisType, group) != null;
            }
            catch (Exception)
            {
                return false;
            }
        }

        private static void TryHideAxis(object chart, int axisType, int group)
        {
            try
            {
                chart.GetType().InvokeMember(
                    "HasAxis",
                    BindingFlags.SetProperty | BindingFlags.Instance | BindingFlags.Public,
                    null,
                    chart,
                    new object[] { axisType, group, false });
            }
            catch (Exception)
            {
            }

            try
            {
                object axis = TryInvoke(chart, "Axes", axisType, group);
                if (axis == null)
                {
                    return;
                }

                TryInvoke(axis, "Delete");
            }
            catch (Exception)
            {
            }

            try
            {
                object axis = TryInvoke(chart, "Axes", axisType, group);
                if (axis == null)
                {
                    return;
                }

                WppCom.TrySetProperty(axis, "HasTitle", false);
                WppCom.TrySetProperty(axis, "HasMajorGridlines", false);
                WppCom.TrySetProperty(axis, "TickLabelPosition", XlTickLabelPositionNone);
                WppCom.TrySetProperty(axis, "MajorTickMark", XlTickMarkNone);
                WppCom.TrySetProperty(axis, "MinorTickMark", XlTickMarkNone);
                object axisLine = WppCom.GetProperty(WppCom.GetProperty(axis, "Format"), "Line");
                WppCom.TrySetProperty(axisLine, "Visible", 0);
            }
            catch (Exception)
            {
            }
        }

        private static float? TryReadFloat(object target, string name)
        {
            try
            {
                object v = WppCom.GetProperty(target, name);
                if (v == null)
                {
                    return null;
                }

                return Convert.ToSingle(v);
            }
            catch (Exception)
            {
                return null;
            }
        }

        private static int? TryReadOleColor(object raw)
        {
            if (raw == null)
            {
                return null;
            }

            try
            {
                int value = Convert.ToInt32(Convert.ToDouble(raw)) & 0x00FFFFFF;
                return value;
            }
            catch (Exception)
            {
                return null;
            }
        }

        private static void StyleLog(List<string> warnings, string message)
        {
            if (string.IsNullOrEmpty(message))
            {
                return;
            }

            string line = "[PptChartSnap] " + message;
            try
            {
                EasyWriteLog.WriteLine(line);
            }
            catch (Exception)
            {
            }

            warnings?.Add(line);
        }

        private static string DescribeSnap(string tag, ChartStyleSnap snap)
        {
            if (snap == null)
            {
                return tag + " snap=null";
            }

            var sb = new StringBuilder();
            sb.Append(tag)
                .Append(" style=").Append(snap.ChartStyle)
                .Append(" title=").Append(snap.HasTitle)
                .Append(" legend=").Append(snap.HasLegend)
                .Append(" plotFill=").Append(snap.PlotFillVisible)
                .Append(" gap=").Append(snap.GapWidth)
                .Append(" overlap=").Append(snap.Overlap)
                .Append(" catDel=").Append(snap.Category == null ? "null" : Convert.ToString(snap.Category.Deleted))
                .Append(" valDel=").Append(snap.Value == null ? "null" : Convert.ToString(snap.Value.Deleted));
            if (snap.Series != null)
            {
                for (int i = 0; i < snap.Series.Count; i++)
                {
                    sb.Append(" | ").Append(DescribeSeries(i + 1, snap.Series[i]));
                }
            }

            return sb.ToString();
        }

        private static string DescribeSeries(int index, SeriesStyleSnap one)
        {
            if (one == null)
            {
                return "S" + index + "=null";
            }

            return "S" + index
                + " type=" + one.ChartType
                + " axis=" + one.AxisGroup
                + " line=" + DescribeLine(one.Line)
                + " marker=" + one.MarkerStyle + "/" + one.MarkerSize
                + " mkRgb=" + HexOf(one.MarkerForeRgb) + "/" + HexOf(one.MarkerBackRgb)
                + " labels=" + one.HasDataLabels
                + " lblPos=" + one.DataLabelPosition
                + " lblColor=" + HexOf(one.DataLabelFontColor)
                + " fillVis=" + (one.Fill == null ? "null" : Convert.ToString(one.Fill.Visible))
                + " fillType=" + (one.Fill == null ? "-" : Convert.ToString(one.Fill.FillType))
                + " fillRgb=" + (one.Fill == null ? "-" : HexOf(one.Fill.SolidRgb))
                + " fillStops=" + DescribeStops(one.Fill == null ? null : one.Fill.Stops);
        }

        private static bool IsPhantomStroke(double? weight, object colorType, object rawRgb)
        {
            if (weight.HasValue && (weight.Value <= 0 || weight.Value > 40))
            {
                return true;
            }

            if (colorType == null)
            {
                return false;
            }

            try
            {
                int type = Convert.ToInt32(colorType);
                if (type == 1)
                {
                    return false;
                }

                if (rawRgb == null)
                {
                    return true;
                }

                int rgb = Convert.ToInt32(Convert.ToDouble(rawRgb)) & 0x00FFFFFF;
                return rgb == 0;
            }
            catch (Exception)
            {
                return false;
            }
        }

        private static string DescribeLine(LineSnap line)
        {
            if (line == null)
            {
                return "null";
            }

            return "vis=" + line.Visible
                + " rgb=" + HexOf(line.Rgb)
                + " w=" + (line.Weight.HasValue ? line.Weight.Value.ToString("0.##", CultureInfo.InvariantCulture) : "-")
                + " stops=" + (line.Stops == null ? 0 : line.Stops.Count);
        }

        private static int? PickSolidFromStops(List<GradientStopSnap> stops)
        {
            if (stops == null || stops.Count == 0)
            {
                return null;
            }

            GradientStopSnap best = stops[0];
            for (int i = 1; i < stops.Count; i++)
            {
                if (stops[i].Transparency <= best.Transparency)
                {
                    best = stops[i];
                }
            }

            return best.Rgb;
        }

        private static string DescribeStops(List<GradientStopSnap> stops)
        {
            if (stops == null || stops.Count == 0)
            {
                return "0";
            }

            var sb = new StringBuilder();
            sb.Append(stops.Count);
            foreach (GradientStopSnap s in stops)
            {
                sb.Append('[')
                    .Append(s.Position.ToString("0.##", CultureInfo.InvariantCulture))
                    .Append(':')
                    .Append(HexOf(s.Rgb))
                    .Append('@')
                    .Append(s.Transparency.ToString("0.##", CultureInfo.InvariantCulture))
                    .Append(']');
            }

            return sb.ToString();
        }

        private static string HexOf(int? rgb)
        {
            if (!rgb.HasValue)
            {
                return "-";
            }

            return OfficeRgbToHex(rgb.Value);
        }

        private static int? TryReadExplicitRgb(object colorFormat)
        {
            if (colorFormat == null)
            {
                return null;
            }

            try
            {
                object type = WppCom.GetProperty(colorFormat, "Type");
                // msoColorTypeRGB = 1；主题/自动色不要写成 #000000 / #FFFFFF
                if (type != null && Convert.ToInt32(type) != 1)
                {
                    return null;
                }

                object rgb = WppCom.GetProperty(colorFormat, "RGB");
                if (rgb == null)
                {
                    return null;
                }

                int value = Convert.ToInt32(Convert.ToDouble(rgb)) & 0x00FFFFFF;
                if (value == 0 && type == null)
                {
                    return null;
                }

                return value;
            }
            catch (Exception)
            {
                return null;
            }
        }

        private static int? TryReadResolvedRgb(object colorFormat)
        {
            if (colorFormat == null)
            {
                return null;
            }

            try
            {
                object rgb = WppCom.GetProperty(colorFormat, "RGB");
                if (rgb == null)
                {
                    return TryReadOleColor(colorFormat);
                }

                int value = Convert.ToInt32(Convert.ToDouble(rgb));
                if (value < 0)
                {
                    return null;
                }

                return value & 0x00FFFFFF;
            }
            catch (Exception)
            {
                return TryReadOleColor(colorFormat);
            }
        }

        private static void TryWriteLineRgb(object line, int rgb)
        {
            try
            {
                object fc = WppCom.GetProperty(line, "ForeColor");
                WppCom.TrySetProperty(fc, "Type", 1);
                WppCom.TrySetProperty(fc, "RGB", rgb);
                WppCom.TrySetProperty(line, "ForeColor", rgb);
            }
            catch (Exception)
            {
            }
        }

        private static int? TryReadMarkerColor(object raw)
        {
            if (raw == null)
            {
                return null;
            }

            try
            {
                int value = Convert.ToInt32(Convert.ToDouble(raw));
                if (value < 0)
                {
                    return null;
                }

                return value & 0x00FFFFFF;
            }
            catch (Exception)
            {
                return null;
            }
        }

        private static string TryReadFontColorHex(object font)
        {
            if (font == null)
            {
                return null;
            }

            try
            {
                object color = WppCom.GetProperty(font, "Color");
                if (color == null)
                {
                    return null;
                }

                return OfficeRgbToHex(Convert.ToInt32(Convert.ToDouble(color)));
            }
            catch (Exception)
            {
                return null;
            }
        }

        private static bool TryReadBox(
            object shape,
            out float left,
            out float top,
            out float width,
            out float height)
        {
            left = top = width = height = 0;
            try
            {
                left = Convert.ToSingle(WppCom.GetProperty(shape, "Left"));
                top = Convert.ToSingle(WppCom.GetProperty(shape, "Top"));
                width = Convert.ToSingle(WppCom.GetProperty(shape, "Width"));
                height = Convert.ToSingle(WppCom.GetProperty(shape, "Height"));
                return width > 0 && height > 0;
            }
            catch (Exception)
            {
                return false;
            }
        }

        public static bool TryCreateOnSlide(
            object shapes,
            float left,
            float top,
            float width,
            float height,
            int xlType,
            PptHtmlChartGrid grid,
            PptHtmlChartFormat format,
            List<string> warnings,
            out object shape,
            out string error,
            int chartStyle = -1,
            bool newLayout = true)
        {
            shape = null;
            error = null;
            if (grid == null || !grid.IsPourable)
            {
                error = "新建 chart 必须在节点内嵌 <table> 灌数，不能建空图";
                return false;
            }

            try
            {
                try
                {
                    shape = WppCom.Invoke(
                        shapes,
                        "AddChart2",
                        chartStyle,
                        xlType,
                        left,
                        top,
                        width,
                        height,
                        newLayout);
                }
                catch (Exception)
                {
                    try
                    {
                        shape = WppCom.Invoke(shapes, "AddChart2", chartStyle, xlType, left, top, width, height);
                    }
                    catch (Exception)
                    {
                        shape = WppCom.Invoke(shapes, "AddChart", xlType, left, top, width, height);
                    }
                }
            }
            catch (Exception ex)
            {
                error = "创建图表失败: " + ex.Message;
                return false;
            }

            if (shape == null)
            {
                error = "创建图表失败";
                return false;
            }

            object chart = TryGetChart(shape);
            if (chart == null)
            {
                TryDelete(shape);
                error = "创建图表后无法访问 Chart";
                return false;
            }

            if (!TryPourGrid(chart, grid, out error))
            {
                TryDelete(shape);
                shape = null;
                return false;
            }

            if (!TryApplyFormat(chart, format, warnings, out error))
            {
                TryDelete(shape);
                shape = null;
                return false;
            }

            EnsureCategoryAxisLabels(chart, grid);
            return true;
        }

        private static object TryGetChart(object shape)
        {
            if (shape == null)
            {
                return null;
            }

            try
            {
                object has = WppCom.GetProperty(shape, "HasChart");
                if (has != null && !IsTruthy(has))
                {
                    return null;
                }
            }
            catch (Exception)
            {
            }

            try
            {
                return WppCom.GetProperty(shape, "Chart");
            }
            catch (Exception)
            {
                return null;
            }
        }

        private static bool TryReadGrid(object chart, out PptHtmlChartGrid grid, out string error)
        {
            grid = null;
            error = null;
            // 只读 SeriesCollection，不 Activate ChartData。
            // 访问内嵌簿会拉起 Excel，读一页带图 HTML 会像卡住。
            if (TryReadGridFromSeries(chart, out grid, out error) && grid != null && grid.IsPourable)
            {
                return true;
            }

            if (string.IsNullOrEmpty(error))
            {
                error = "无法读取图表数据";
            }

            return false;
        }

        private static bool TryReadGridFromSeries(object chart, out PptHtmlChartGrid grid, out string error)
        {
            grid = null;
            error = null;
            try
            {
                object sc = TryInvoke(chart, "SeriesCollection");
                if (sc == null)
                {
                    sc = WppCom.GetProperty(chart, "SeriesCollection");
                }

                int count = Convert.ToInt32(WppCom.GetProperty(sc, "Count"));
                if (count < 1)
                {
                    error = "图表没有系列";
                    return false;
                }

                object s1 = GetSeries(chart, 1);
                object xvals = WppCom.GetProperty(s1, "XValues");
                List<string> cats = ToStringList(xvals);
                if (cats.Count == 0)
                {
                    error = "图表没有类别";
                    return false;
                }

                var columns = new List<PptHtmlChartColumn>
                {
                    new PptHtmlChartColumn { Role = "category", Name = "类别" }
                };
                var seriesValues = new List<List<string>>();
                int useSeries = Math.Min(count, MaxCols - 1);
                for (int i = 1; i <= useSeries; i++)
                {
                    object series = GetSeries(chart, i);
                    string name = Convert.ToString(WppCom.GetProperty(series, "Name") ?? ("系列" + i));
                    var col = new PptHtmlChartColumn
                    {
                        Role = "value",
                        Name = name,
                        Color = TryReadSeriesColor(series)
                    };
                    columns.Add(col);
                    seriesValues.Add(ToStringList(WppCom.GetProperty(series, "Values")));
                }

                int useRows = Math.Min(cats.Count, MaxRows - 1);
                var rows = new List<List<string>>();
                for (int r = 0; r < useRows; r++)
                {
                    var row = new List<string> { cats[r] };
                    for (int s = 0; s < seriesValues.Count; s++)
                    {
                        row.Add(r < seriesValues[s].Count ? seriesValues[s][r] : "");
                    }

                    rows.Add(row);
                }

                grid = new PptHtmlChartGrid
                {
                    Columns = columns,
                    Rows = rows,
                    Truncated = cats.Count + 1 > MaxRows || count + 1 > MaxCols
                };
                return grid.IsPourable;
            }
            catch (Exception ex)
            {
                error = "读 SeriesCollection 失败: " + ex.Message;
                return false;
            }
        }

        private static void ApplySeriesExtras(object chart, PptHtmlChartGrid grid)
        {
            try
            {
                object sc = TryInvoke(chart, "SeriesCollection");
                int si = 0;
                for (int i = 0; i < grid.Columns.Count; i++)
                {
                    PptHtmlChartColumn col = grid.Columns[i];
                    if (col.Role == "category")
                    {
                        continue;
                    }

                    si++;
                    object series = WppCom.GetIndexed(sc, si);
                    if (series == null)
                    {
                        continue;
                    }

                    if (!string.IsNullOrEmpty(col.Color) && TryParseHexToOffice(col.Color, out int rgb))
                    {
                        try
                        {
                            object fmt = WppCom.GetProperty(series, "Format");
                            object fill = WppCom.GetProperty(fmt, "Fill");
                            object fc = WppCom.GetProperty(fill, "ForeColor");
                            WppCom.TrySetProperty(fc, "RGB", rgb);
                        }
                        catch (Exception)
                        {
                        }
                    }

                    if (!string.IsNullOrWhiteSpace(col.SeriesType)
                        && TryParseType(col.SeriesType, out int xl, out _, out _))
                    {
                        WppCom.TrySetProperty(series, "ChartType", xl);
                    }

                    if (string.Equals(col.AxisY, "secondary", StringComparison.OrdinalIgnoreCase))
                    {
                        WppCom.TrySetProperty(series, "AxisGroup", XlSecondary);
                    }
                }
            }
            catch (Exception)
            {
            }
        }

        private static void TryReadTitle(object chart, PptHtmlChartFormat format)
        {
            try
            {
                object has = WppCom.GetProperty(chart, "HasTitle");
                if (!IsTruthy(has))
                {
                    return;
                }

                object title = WppCom.GetProperty(chart, "ChartTitle");
                format.Title = Convert.ToString(WppCom.GetProperty(title, "Text") ?? "");
                object font = WppCom.GetProperty(title, "Font");
                if (font != null)
                {
                    object sz = WppCom.GetProperty(font, "Size");
                    if (sz != null)
                    {
                        format.TitleFontSize = Convert.ToDouble(sz).ToString("0.##", CultureInfo.InvariantCulture);
                    }

                    object bold = WppCom.GetProperty(font, "Bold");
                    if (bold != null)
                    {
                        format.TitleFontBold = IsTruthy(bold) ? "true" : "false";
                    }

                    object color = WppCom.GetProperty(font, "Color");
                    if (color != null)
                    {
                        format.TitleFontColor = OfficeRgbToHex(Convert.ToInt32(Convert.ToDouble(color)));
                    }
                }
            }
            catch (Exception)
            {
            }
        }

        private static void TryReadLegend(object chart, PptHtmlChartFormat format)
        {
            try
            {
                object has = WppCom.GetProperty(chart, "HasLegend");
                if (!IsTruthy(has))
                {
                    format.Legend = "none";
                    return;
                }

                object legend = WppCom.GetProperty(chart, "Legend");
                object pos = WppCom.GetProperty(legend, "Position");
                format.Legend = LegendFromXl(Convert.ToInt32(pos));
            }
            catch (Exception)
            {
            }
        }

        private static void TryReadPlotAndLabels(object chart, PptHtmlChartFormat format)
        {
            try
            {
                object plot = WppCom.GetProperty(chart, "PlotArea");
                object fill = WppCom.GetProperty(WppCom.GetProperty(plot, "Format"), "Fill");
                if (!IsTruthy(WppCom.GetProperty(fill, "Visible")))
                {
                    return;
                }

                object fc = WppCom.GetProperty(fill, "ForeColor");
                object type = WppCom.GetProperty(fc, "Type");
                // msoColorTypeRGB = 1；主题/自动色读出来常是 0，不能当成 #000000
                if (type != null && Convert.ToInt32(type) != 1)
                {
                    return;
                }

                object rgb = WppCom.GetProperty(fc, "RGB");
                if (rgb == null)
                {
                    return;
                }

                int rgbVal = Convert.ToInt32(Convert.ToDouble(rgb));
                if (rgbVal == 0 && type == null)
                {
                    return;
                }

                format.PlotColor = OfficeRgbToHex(rgbVal);
            }
            catch (Exception)
            {
            }
        }

        private static void TryReadAxes(object chart, PptHtmlChartFormat format)
        {
            try
            {
                object x = TryInvoke(chart, "Axes", XlCategory, XlPrimary);
                if (x != null && IsTruthy(WppCom.GetProperty(x, "HasTitle")))
                {
                    object t = WppCom.GetProperty(x, "AxisTitle");
                    format.AxisX = Convert.ToString(WppCom.GetProperty(t, "Text") ?? "");
                }
            }
            catch (Exception)
            {
            }

            try
            {
                object y = TryInvoke(chart, "Axes", XlValue, XlPrimary);
                if (y != null && IsTruthy(WppCom.GetProperty(y, "HasTitle")))
                {
                    object t = WppCom.GetProperty(y, "AxisTitle");
                    format.AxisY = Convert.ToString(WppCom.GetProperty(t, "Text") ?? "");
                }
            }
            catch (Exception)
            {
            }
        }

        private static void TrySetTitle(object chart, string title, List<string> warnings)
        {
            try
            {
                if (string.IsNullOrEmpty(title))
                {
                    WppCom.TrySetProperty(chart, "HasTitle", false);
                    return;
                }

                WppCom.TrySetProperty(chart, "HasTitle", true);
                object ct = WppCom.GetProperty(chart, "ChartTitle");
                WppCom.TrySetProperty(ct, "Text", title);
            }
            catch (Exception)
            {
                warnings?.Add("未能套用 data-title");
            }
        }

        private static void TrySetOptional(object chart, PptHtmlChartFormat fmt, List<string> warnings)
        {
            if (!string.IsNullOrWhiteSpace(fmt.Theme))
            {
                try
                {
                    int style = string.Equals(fmt.Theme, "office", StringComparison.OrdinalIgnoreCase) ? 1 : 1;
                    object cur = WppCom.GetProperty(chart, "ChartStyle");
                    if (cur == null || Convert.ToInt32(cur) != style)
                    {
                        WppCom.TrySetProperty(chart, "ChartStyle", style);
                    }
                }
                catch (Exception)
                {
                    Warn(warnings, "data-theme");
                }
            }

            if (!string.IsNullOrWhiteSpace(fmt.Legend))
            {
                try
                {
                    if (string.Equals(fmt.Legend, "none", StringComparison.OrdinalIgnoreCase))
                    {
                        WppCom.TrySetProperty(chart, "HasLegend", false);
                    }
                    else
                    {
                        WppCom.TrySetProperty(chart, "HasLegend", true);
                        object legend = WppCom.GetProperty(chart, "Legend");
                        WppCom.TrySetProperty(legend, "Position", LegendToXl(fmt.Legend));
                    }
                }
                catch (Exception)
                {
                    Warn(warnings, "data-legend");
                }
            }

            if (!string.IsNullOrWhiteSpace(fmt.ShowDataLabels))
            {
                try
                {
                    bool on = IsTrue(fmt.ShowDataLabels);
                    object sc = TryInvoke(chart, "SeriesCollection");
                    int n = Convert.ToInt32(WppCom.GetProperty(sc, "Count"));
                    for (int i = 1; i <= n; i++)
                    {
                        object s = WppCom.GetIndexed(sc, i);
                        WppCom.TrySetProperty(s, "HasDataLabels", on);
                    }
                }
                catch (Exception)
                {
                    Warn(warnings, "data-show-data-labels");
                }
            }

            if (!string.IsNullOrWhiteSpace(fmt.PlotColor) && TryParseHexToOffice(fmt.PlotColor, out int plotRgb))
            {
                try
                {
                    object plot = WppCom.GetProperty(chart, "PlotArea");
                    object fill = WppCom.GetProperty(WppCom.GetProperty(plot, "Format"), "Fill");
                    if (!FillRgbAlreadyMatches(fill, plotRgb))
                    {
                        TryInvoke(fill, "Solid");
                        object fc = WppCom.GetProperty(fill, "ForeColor");
                        WppCom.TrySetProperty(fc, "RGB", plotRgb);
                    }
                }
                catch (Exception)
                {
                    Warn(warnings, "data-plot-color");
                }
            }

            if (!string.IsNullOrWhiteSpace(fmt.TitleFontSize)
                || !string.IsNullOrWhiteSpace(fmt.TitleFontBold)
                || !string.IsNullOrWhiteSpace(fmt.TitleFontColor))
            {
                try
                {
                    object title = WppCom.GetProperty(chart, "ChartTitle");
                    object font = WppCom.GetProperty(title, "Font");
                    if (!string.IsNullOrWhiteSpace(fmt.TitleFontSize)
                        && double.TryParse(fmt.TitleFontSize, NumberStyles.Float, CultureInfo.InvariantCulture, out double sz))
                    {
                        WppCom.TrySetProperty(font, "Size", sz);
                    }

                    if (!string.IsNullOrWhiteSpace(fmt.TitleFontBold))
                    {
                        WppCom.TrySetProperty(font, "Bold", IsTrue(fmt.TitleFontBold));
                    }

                    if (!string.IsNullOrWhiteSpace(fmt.TitleFontColor)
                        && TryParseHexToOffice(fmt.TitleFontColor, out int tRgb))
                    {
                        WppCom.TrySetProperty(font, "Color", tRgb);
                    }
                }
                catch (Exception)
                {
                    Warn(warnings, "data-title-font-*");
                }
            }

            TrySetAxisTitle(chart, XlCategory, XlPrimary, fmt.AxisX, warnings, "data-axis-x");
            TrySetAxisTitle(chart, XlValue, XlPrimary, fmt.AxisY, warnings, "data-axis-y");
            TrySetAxisTitle(chart, XlValue, XlSecondary, fmt.AxisYSecondary, warnings, "data-axis-y-secondary");

            if (!string.IsNullOrWhiteSpace(fmt.AxisYMin)
                || !string.IsNullOrWhiteSpace(fmt.AxisYMax)
                || !string.IsNullOrWhiteSpace(fmt.AxisYMajorUnit))
            {
                try
                {
                    object y = TryInvoke(chart, "Axes", XlValue, XlPrimary);
                    if (!string.IsNullOrWhiteSpace(fmt.AxisYMin)
                        && double.TryParse(fmt.AxisYMin, NumberStyles.Float, CultureInfo.InvariantCulture, out double mn))
                    {
                        WppCom.TrySetProperty(y, "MinimumScale", mn);
                    }

                    if (!string.IsNullOrWhiteSpace(fmt.AxisYMax)
                        && double.TryParse(fmt.AxisYMax, NumberStyles.Float, CultureInfo.InvariantCulture, out double mx))
                    {
                        WppCom.TrySetProperty(y, "MaximumScale", mx);
                    }

                    if (!string.IsNullOrWhiteSpace(fmt.AxisYMajorUnit)
                        && double.TryParse(fmt.AxisYMajorUnit, NumberStyles.Float, CultureInfo.InvariantCulture, out double un))
                    {
                        WppCom.TrySetProperty(y, "MajorUnit", un);
                    }
                }
                catch (Exception)
                {
                    Warn(warnings, "data-axis-y-min/max/major-unit");
                }
            }

            if (!string.IsNullOrWhiteSpace(fmt.GapWidth)
                && int.TryParse(fmt.GapWidth, NumberStyles.Integer, CultureInfo.InvariantCulture, out int gap))
            {
                try
                {
                    object g = TryInvoke(chart, "ChartGroups", 1);
                    WppCom.TrySetProperty(g, "GapWidth", gap);
                }
                catch (Exception)
                {
                    Warn(warnings, "data-gap-width");
                }
            }

            if (!string.IsNullOrWhiteSpace(fmt.Explosion)
                && int.TryParse(fmt.Explosion, NumberStyles.Integer, CultureInfo.InvariantCulture, out int exp))
            {
                try
                {
                    object sc = TryInvoke(chart, "SeriesCollection");
                    object s = WppCom.GetIndexed(sc, 1);
                    WppCom.TrySetProperty(s, "Explosion", exp);
                }
                catch (Exception)
                {
                    Warn(warnings, "data-explosion");
                }
            }
        }

        private static void TrySetAxisTitle(
            object chart,
            int axisType,
            int group,
            string title,
            List<string> warnings,
            string field)
        {
            if (title == null)
            {
                return;
            }

            try
            {
                object axis = TryInvoke(chart, "Axes", axisType, group);
                if (string.IsNullOrEmpty(title))
                {
                    WppCom.TrySetProperty(axis, "HasTitle", false);
                    return;
                }

                WppCom.TrySetProperty(axis, "HasTitle", true);
                object at = WppCom.GetProperty(axis, "AxisTitle");
                WppCom.TrySetProperty(at, "Text", title);
            }
            catch (Exception)
            {
                Warn(warnings, field);
            }
        }

        private static void SuppressExcel(object excelApp)
        {
            if (excelApp == null)
            {
                return;
            }

            try
            {
                WppCom.TrySetProperty(excelApp, "DisplayAlerts", false);
                WppCom.TrySetProperty(excelApp, "ScreenUpdating", false);
                WppCom.TrySetProperty(excelApp, "Visible", false);
            }
            catch (Exception)
            {
            }
        }

        private static void HideEmbeddedExcel(object excelApp)
        {
            SuppressExcel(excelApp);
            if (excelApp == null)
            {
                return;
            }

            try
            {
                object windows = WppCom.GetProperty(excelApp, "Windows");
                int count = Convert.ToInt32(WppCom.GetProperty(windows, "Count"));
                for (int i = count; i >= 1; i--)
                {
                    object win = WppCom.GetIndexed(windows, i);
                    WppCom.TrySetProperty(win, "Visible", false);
                }
            }
            catch (Exception)
            {
            }
        }

        private static bool TryExpandViaChartData(
            object chart,
            PptHtmlChartGrid grid,
            out object excelApp,
            out string error)
        {
            excelApp = null;
            error = null;
            try
            {
                object chartData = WppCom.GetProperty(chart, "ChartData");
                if (chartData == null)
                {
                    error = "无法访问 ChartData";
                    return false;
                }

                TryInvoke(chartData, "Activate");
                object workbook = WppCom.GetProperty(chartData, "Workbook");
                if (workbook == null)
                {
                    error = "无法打开图表内嵌工作簿";
                    return false;
                }

                excelApp = WppCom.GetProperty(workbook, "Application");
                SuppressExcel(excelApp);
                object sheets = WppCom.GetProperty(workbook, "Worksheets");
                object ws = WppCom.GetIndexed(sheets, 1);
                if (ws == null)
                {
                    error = "图表内嵌表不存在";
                    return false;
                }

                TryClearSheet(ws);
                int cols = grid.Columns.Count;
                int rows = grid.Rows.Count;
                for (int c = 0; c < cols; c++)
                {
                    SetCell(ws, 1, c + 1, grid.Columns[c].Name ?? "");
                }

                for (int r = 0; r < rows; r++)
                {
                    for (int c = 0; c < cols; c++)
                    {
                        string raw = r < grid.Rows.Count && c < grid.Rows[r].Count ? grid.Rows[r][c] : "";
                        if (grid.Columns[c].Role == "value"
                            && double.TryParse(raw, NumberStyles.Float, CultureInfo.InvariantCulture, out double n))
                        {
                            SetCell(ws, r + 2, c + 1, n);
                        }
                        else
                        {
                            SetCell(ws, r + 2, c + 1, raw ?? "");
                        }
                    }
                }

                object range = TryGetDataRange(ws, rows + 1, cols);
                if (range != null)
                {
                    TrySetSourceData(chart, range);
                }

                return true;
            }
            catch (Exception ex)
            {
                error = "灌入图表数据失败: " + ex.Message;
                return false;
            }
        }

        private static object TryGetDataRange(object ws, int lastRow, int lastCol)
        {
            string a2 = ColLetter(lastCol) + lastRow.ToString(CultureInfo.InvariantCulture);
            object range = TryGetExcelRange(ws, "A1", a2);
            if (range != null)
            {
                return range;
            }

            range = TryGetExcelRange(ws, "A1:" + a2, null);
            if (range != null)
            {
                return range;
            }

            object c1 = GetCellObject(ws, 1, 1);
            object c2 = GetCellObject(ws, lastRow, lastCol);
            if (c1 != null && c2 != null)
            {
                range = TryGetExcelRange(ws, c1, c2);
                if (range != null)
                {
                    return range;
                }
            }

            try
            {
                return WppCom.GetProperty(ws, "UsedRange");
            }
            catch (Exception)
            {
                return null;
            }
        }

        private static object TryGetExcelRange(object ws, object arg1, object arg2)
        {
            if (ws == null || arg1 == null)
            {
                return null;
            }

            object[] args = arg2 == null ? new object[] { arg1 } : new object[] { arg1, arg2 };
            const BindingFlags flags = BindingFlags.GetProperty
                | BindingFlags.InvokeMethod
                | BindingFlags.Instance
                | BindingFlags.Public;
            try
            {
                return ws.GetType().InvokeMember("Range", flags, null, ws, args);
            }
            catch (Exception)
            {
            }

            try
            {
                return ws.GetType().InvokeMember("get_Range", flags, null, ws, args);
            }
            catch (Exception)
            {
                return null;
            }
        }

        private static bool TrySetSourceData(object chart, object range)
        {
            try
            {
                WppCom.Invoke(chart, "SetSourceData", range, 2);
                return true;
            }
            catch (Exception)
            {
            }

            try
            {
                WppCom.Invoke(chart, "SetSourceData", range);
                return true;
            }
            catch (Exception)
            {
            }

            return false;
        }

        private static object GetSeries(object chart, int index)
        {
            object series = TryInvoke(chart, "SeriesCollection", index);
            if (series != null)
            {
                return series;
            }

            object sc = TryInvoke(chart, "SeriesCollection");
            if (sc == null)
            {
                sc = WppCom.GetProperty(chart, "SeriesCollection");
            }

            return WppCom.GetIndexed(sc, index);
        }

        private static int GetSeriesCount(object chart)
        {
            object sc = TryInvoke(chart, "SeriesCollection");
            if (sc == null)
            {
                sc = WppCom.GetProperty(chart, "SeriesCollection");
            }

            return Convert.ToInt32(WppCom.GetProperty(sc, "Count"));
        }

        private static bool FillRgbAlreadyMatches(object fill, int rgb)
        {
            try
            {
                object vis = WppCom.GetProperty(fill, "Visible");
                if (vis != null && Convert.ToInt32(vis) == 0)
                {
                    return false;
                }

                object fc = WppCom.GetProperty(fill, "ForeColor");
                object cur = fc == null ? null : WppCom.GetProperty(fc, "RGB");
                if (cur == null)
                {
                    return false;
                }

                return (Convert.ToInt32(cur) & 0x00FFFFFF) == (rgb & 0x00FFFFFF);
            }
            catch (Exception)
            {
                return false;
            }
        }

        private static bool IsSameChartType(object chart, string canon)
        {
            try
            {
                object t = WppCom.GetProperty(chart, "ChartType");
                if (t == null)
                {
                    return false;
                }

                return string.Equals(
                    CanonicalTypeFromXl(Convert.ToInt32(t)),
                    canon,
                    StringComparison.OrdinalIgnoreCase);
            }
            catch (Exception)
            {
                return false;
            }
        }

        private static bool GridAlreadyMatches(object chart, PptHtmlChartGrid grid)
        {
            if (chart == null || grid == null || !grid.IsPourable)
            {
                return false;
            }

            if (!TryReadGridFromSeries(chart, out PptHtmlChartGrid cur, out _)
                || cur == null
                || !cur.IsPourable
                || cur.Columns == null
                || cur.Rows == null
                || cur.Columns.Count != grid.Columns.Count
                || cur.Rows.Count != grid.Rows.Count)
            {
                return false;
            }

            for (int c = 0; c < grid.Columns.Count; c++)
            {
                string role = grid.Columns[c].Role ?? "";
                if (!string.Equals(cur.Columns[c].Role ?? "", role, StringComparison.OrdinalIgnoreCase))
                {
                    return false;
                }

                if (role != "category"
                    && !string.Equals(cur.Columns[c].Name ?? "", grid.Columns[c].Name ?? "", StringComparison.Ordinal))
                {
                    return false;
                }
            }

            for (int r = 0; r < grid.Rows.Count; r++)
            {
                List<string> a = grid.Rows[r];
                List<string> b = cur.Rows[r];
                if (a == null || b == null || a.Count < grid.Columns.Count || b.Count < grid.Columns.Count)
                {
                    return false;
                }

                for (int c = 0; c < grid.Columns.Count; c++)
                {
                    if (grid.Columns[c].Role == "value")
                    {
                        if (!double.TryParse(a[c], NumberStyles.Float, CultureInfo.InvariantCulture, out double x)
                            || !double.TryParse(b[c], NumberStyles.Float, CultureInfo.InvariantCulture, out double y)
                            || Math.Abs(x - y) > 0.0001)
                        {
                            return false;
                        }
                    }
                    else if (!string.Equals(a[c] ?? "", b[c] ?? "", StringComparison.Ordinal))
                    {
                        return false;
                    }
                }
            }

            return true;
        }

        private static List<string> CategoryLabels(PptHtmlChartGrid grid)
        {
            var cats = new List<string>();
            if (grid == null || grid.Rows == null)
            {
                return cats;
            }

            int catCol = 0;
            if (grid.Columns != null)
            {
                for (int i = 0; i < grid.Columns.Count; i++)
                {
                    if (grid.Columns[i].Role == "category")
                    {
                        catCol = i;
                        break;
                    }
                }
            }

            foreach (List<string> row in grid.Rows)
            {
                cats.Add(row != null && catCol < row.Count ? (row[catCol] ?? "") : "");
            }

            return cats;
        }

        private static void EnsureCategoryAxisLabels(object chart, PptHtmlChartGrid grid)
        {
            if (chart == null)
            {
                return;
            }

            try
            {
                object axis = TryInvoke(chart, "Axes", XlCategory, XlPrimary);
                if (axis == null)
                {
                    return;
                }

                // 2018/2019 这类标签灌进 XValues 后常被收成时间轴，横坐标字会空白
                WppCom.TrySetProperty(axis, "CategoryType", XlCategoryScale);

                List<string> cats = CategoryLabels(grid);
                if (cats.Count > 0)
                {
                    WppCom.TrySetProperty(axis, "CategoryNames", cats.ToArray());
                }

                object pos = WppCom.GetProperty(axis, "TickLabelPosition");
                if (pos == null || Convert.ToInt32(pos) == XlTickLabelPositionNone)
                {
                    WppCom.TrySetProperty(axis, "TickLabelPosition", XlTickLabelPositionNextToAxis);
                }

                try
                {
                    object ticks = WppCom.GetProperty(axis, "TickLabels");
                    WppCom.TrySetProperty(ticks, "NumberFormat", "@");
                }
                catch (Exception)
                {
                }
            }
            catch (Exception)
            {
            }
        }

        private static void TryCaptureTickLabelFont(object chart, out object color, out object size)
        {
            color = null;
            size = null;
            try
            {
                object axis = TryInvoke(chart, "Axes", XlCategory, XlPrimary);
                object ticks = axis == null ? null : WppCom.GetProperty(axis, "TickLabels");
                object font = ticks == null ? null : WppCom.GetProperty(ticks, "Font");
                if (font == null)
                {
                    return;
                }

                color = WppCom.GetProperty(font, "Color");
                size = WppCom.GetProperty(font, "Size");
            }
            catch (Exception)
            {
            }
        }

        private static void TryRestoreTickLabelFont(object chart, object color, object size)
        {
            if (color == null && size == null)
            {
                return;
            }

            try
            {
                object axis = TryInvoke(chart, "Axes", XlCategory, XlPrimary);
                object ticks = axis == null ? null : WppCom.GetProperty(axis, "TickLabels");
                object font = ticks == null ? null : WppCom.GetProperty(ticks, "Font");
                if (font == null)
                {
                    return;
                }

                if (color != null)
                {
                    WppCom.TrySetProperty(font, "Color", color);
                }

                if (size != null)
                {
                    WppCom.TrySetProperty(font, "Size", size);
                }
            }
            catch (Exception)
            {
            }
        }

        private static bool TryEnsureSeriesCount(object chart, int wantSeries, out string error)
        {
            error = null;
            if (wantSeries < 1)
            {
                return true;
            }

            try
            {
                int count = GetSeriesCount(chart);
                for (int i = count; i < wantSeries; i++)
                {
                    object sc = TryInvoke(chart, "SeriesCollection");
                    if (sc == null)
                    {
                        sc = WppCom.GetProperty(chart, "SeriesCollection");
                    }

                    if (TryInvoke(sc, "NewSeries") == null && GetSeriesCount(chart) <= i)
                    {
                        error = "无法把图表扩到 " + wantSeries + " 个系列";
                        return false;
                    }
                }

                return GetSeriesCount(chart) >= wantSeries;
            }
            catch (Exception ex)
            {
                error = "扩展图表系列失败: " + ex.Message;
                return false;
            }
        }

        private static bool TryPourSeriesLikeWord(object chart, PptHtmlChartGrid grid, out string error)
        {
            error = null;
            var cats = new List<string>();
            if (grid.Rows != null)
            {
                foreach (List<string> row in grid.Rows)
                {
                    cats.Add(row != null && row.Count > 0 ? (row[0] ?? "") : "");
                }
            }

            int si = 0;
            for (int c = 0; c < grid.Columns.Count; c++)
            {
                if (grid.Columns[c].Role == "category")
                {
                    continue;
                }

                si++;
                object series = GetSeries(chart, si);
                if (series == null)
                {
                    error = "无法访问图表系列 " + si;
                    return false;
                }

                var values = new List<double>();
                if (grid.Rows != null)
                {
                    foreach (List<string> row in grid.Rows)
                    {
                        string raw = row != null && c < row.Count ? row[c] : "";
                        if (!double.TryParse(raw, NumberStyles.Float, CultureInfo.InvariantCulture, out double n))
                        {
                            n = 0;
                        }

                        values.Add(n);
                    }
                }

                try
                {
                    WppCom.TrySetProperty(series, "Values", values.ToArray());
                    if (cats.Count > 0)
                    {
                        WppCom.TrySetProperty(series, "XValues", cats.ToArray());
                    }

                    if (!string.IsNullOrEmpty(grid.Columns[c].Name))
                    {
                        WppCom.TrySetProperty(series, "Name", grid.Columns[c].Name);
                    }
                }
                catch (Exception ex)
                {
                    error = "写入系列 " + si + " 失败: " + ex.Message;
                    return false;
                }
            }

            return true;
        }

        private static bool TrimExtraSeries(object chart, int wantSeries, out string error)
        {
            error = null;
            if (wantSeries < 1)
            {
                return true;
            }

            try
            {
                int total = GetSeriesCount(chart);
                for (int i = total; i > wantSeries; i--)
                {
                    try
                    {
                        object series = GetSeries(chart, i);
                        if (series != null)
                        {
                            WppCom.Invoke(series, "Delete");
                        }
                    }
                    catch (Exception)
                    {
                    }
                }

                return true;
            }
            catch (Exception ex)
            {
                error = "删除图表多余系列失败: " + ex.Message;
                return false;
            }
        }

        private static void TryClearSheet(object ws)
        {
            try
            {
                object block = TryGetExcelRange(ws, "A1", "H40");
                if (block == null)
                {
                    block = TryGetExcelRange(ws, "A1:H40", null);
                }

                TryInvoke(block, "Clear");
                TryInvoke(block, "ClearContents");
            }
            catch (Exception)
            {
            }

            try
            {
                object used = WppCom.GetProperty(ws, "UsedRange");
                TryInvoke(used, "Clear");
            }
            catch (Exception)
            {
            }
        }

        private static object GetCellObject(object ws, int row, int col)
        {
            try
            {
                object cells = WppCom.GetProperty(ws, "Cells");
                if (cells != null)
                {
                    return cells.GetType().InvokeMember(
                        "Item",
                        BindingFlags.GetProperty
                        | BindingFlags.InvokeMethod
                        | BindingFlags.Instance
                        | BindingFlags.Public,
                        null,
                        cells,
                        new object[] { row, col });
                }
            }
            catch (Exception)
            {
            }

            try
            {
                return WppCom.Invoke(ws, "Cells", row, col);
            }
            catch (Exception)
            {
                return null;
            }
        }

        private static void SetCell(object ws, int row, int col, object value)
        {
            object cell = GetCellObject(ws, row, col);
            WppCom.TrySetProperty(cell, "Value", value);
            WppCom.TrySetProperty(cell, "Value2", value);
        }

        private static object GetCell(object ws, int row, int col)
        {
            try
            {
                object cell = GetCellObject(ws, row, col);
                object v = WppCom.GetProperty(cell, "Value2");
                return v ?? WppCom.GetProperty(cell, "Value");
            }
            catch (Exception)
            {
                return null;
            }
        }

        private static object TryInvoke(object target, string name, params object[] args)
        {
            if (target == null)
            {
                return null;
            }

            try
            {
                return WppCom.Invoke(target, name, args);
            }
            catch (Exception)
            {
                return null;
            }
        }

        private static void TryDelete(object shape)
        {
            try
            {
                WppCom.Invoke(shape, "Delete");
            }
            catch (Exception)
            {
            }
        }

        private static List<string> ToStringList(object values)
        {
            var list = new List<string>();
            if (values == null)
            {
                return list;
            }

            if (values is Array arr)
            {
                foreach (object item in arr)
                {
                    list.Add(FormatCell(item));
                }

                return list;
            }

            list.Add(FormatCell(values));
            return list;
        }

        private static string FormatCell(object v)
        {
            if (v == null || v is DBNull)
            {
                return "";
            }

            if (v is double d)
            {
                return d.ToString("0.####", CultureInfo.InvariantCulture);
            }

            if (v is float f)
            {
                return f.ToString("0.####", CultureInfo.InvariantCulture);
            }

            if (v is decimal m)
            {
                return m.ToString("0.####", CultureInfo.InvariantCulture);
            }

            return Convert.ToString(v, CultureInfo.InvariantCulture) ?? "";
        }

        private static string TryReadSeriesColor(object series)
        {
            try
            {
                object fmt = WppCom.GetProperty(series, "Format");
                object fill = WppCom.GetProperty(fmt, "Fill");
                object fc = WppCom.GetProperty(fill, "ForeColor");
                object rgb = WppCom.GetProperty(fc, "RGB");
                if (rgb != null)
                {
                    return OfficeRgbToHex(Convert.ToInt32(Convert.ToDouble(rgb)));
                }
            }
            catch (Exception)
            {
            }

            return null;
        }

        private static bool TryParseHexToOffice(string hex, out int office)
        {
            office = 0;
            string t = NormalizeHexOrNull(hex);
            if (t == null || t.Length != 7)
            {
                return false;
            }

            if (!int.TryParse(t.Substring(1, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out int r)
                || !int.TryParse(t.Substring(3, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out int g)
                || !int.TryParse(t.Substring(5, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out int b))
            {
                return false;
            }

            office = r + (g << 8) + (b << 16);
            return true;
        }

        private static string OfficeRgbToHex(int office)
        {
            int r = office & 0xFF;
            int g = (office >> 8) & 0xFF;
            int b = (office >> 16) & 0xFF;
            return "#" + r.ToString("X2", CultureInfo.InvariantCulture)
                + g.ToString("X2", CultureInfo.InvariantCulture)
                + b.ToString("X2", CultureInfo.InvariantCulture);
        }

        private static string NormalizeHexOrNull(string raw)
        {
            if (string.IsNullOrWhiteSpace(raw))
            {
                return null;
            }

            string t = raw.Trim();
            if (t.StartsWith("#", StringComparison.Ordinal) && t.Length == 7)
            {
                return t.ToUpperInvariant();
            }

            return null;
        }

        private static int LegendToXl(string pos)
        {
            switch ((pos ?? "").Trim().ToLowerInvariant())
            {
                case "top":
                    return -4160;
                case "left":
                    return -4131;
                case "right":
                    return -4152;
                default:
                    return -4107;
            }
        }

        private static string LegendFromXl(int pos)
        {
            if (pos == -4160)
            {
                return "top";
            }

            if (pos == -4131)
            {
                return "left";
            }

            if (pos == -4152)
            {
                return "right";
            }

            return "bottom";
        }

        private static string ColLetter(int col)
        {
            string s = "";
            int n = col;
            while (n > 0)
            {
                int m = (n - 1) % 26;
                s = (char)('A' + m) + s;
                n = (n - 1) / 26;
            }

            return s;
        }

        private static bool IsTruthy(object v)
        {
            if (v == null)
            {
                return false;
            }

            if (v is bool b)
            {
                return b;
            }

            try
            {
                int n = Convert.ToInt32(v);
                return n != 0 && n != -4142;
            }
            catch (Exception)
            {
                return false;
            }
        }

        private static bool IsTrue(string raw)
        {
            return string.Equals(raw, "true", StringComparison.OrdinalIgnoreCase)
                || string.Equals(raw, "1", StringComparison.OrdinalIgnoreCase)
                || string.Equals(raw, "yes", StringComparison.OrdinalIgnoreCase);
        }

        private static void Warn(List<string> warnings, string field)
        {
            warnings?.Add("未能套用 " + field);
        }

        private static void WriteAttr(StringBuilder sb, string name, string value)
        {
            if (string.IsNullOrEmpty(value))
            {
                return;
            }

            sb.Append(" ").Append(name).Append("=\"").Append(EscapeAttr(value)).Append("\"");
        }

        private static string GetAttr(XElement el, string name)
        {
            if (el == null)
            {
                return null;
            }

            XAttribute a = el.Attribute(name);
            if (a == null)
            {
                a = el.Attributes().FirstOrDefault(x =>
                    string.Equals(x.Name.LocalName, name, StringComparison.OrdinalIgnoreCase));
            }

            return a == null ? null : a.Value;
        }

        private static string InnerText(XElement el)
        {
            if (el == null)
            {
                return "";
            }

            return string.Concat(el.DescendantNodes().OfType<XText>().Select(t => t.Value)).Trim();
        }

        private static string EscapeAttr(string text)
        {
            if (string.IsNullOrEmpty(text))
            {
                return "";
            }

            return text.Replace("&", "&amp;").Replace("\"", "&quot;").Replace("<", "&lt;").Replace(">", "&gt;");
        }

        private static string EscapeText(string text)
        {
            if (string.IsNullOrEmpty(text))
            {
                return "";
            }

            return text.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;");
        }
    }
}
