using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
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
        private const int XlPie = 5;
        private const int Xl3DPie = -4102;

        private const int XlCategory = 1;
        private const int XlValue = 2;
        private const int XlPrimary = 1;
        private const int XlSecondary = 2;

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

            object workbook = null;
            object excelApp = null;
            try
            {
                object chartData = WppCom.GetProperty(chart, "ChartData");
                if (chartData == null)
                {
                    error = "无法访问 ChartData";
                    return false;
                }

                TryInvoke(chartData, "Activate");
                workbook = WppCom.GetProperty(chartData, "Workbook");
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
                int wantSeries = 0;
                for (int c = 0; c < cols; c++)
                {
                    SetCell(ws, 1, c + 1, grid.Columns[c].Name ?? "");
                    if (grid.Columns[c].Role != "category")
                    {
                        wantSeries++;
                    }
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

                string lastCol = ColLetter(cols);
                string lastRow = (rows + 1).ToString(CultureInfo.InvariantCulture);
                object range = TryInvoke(ws, "Range", "A1", lastCol + lastRow);
                if (range == null)
                {
                    range = TryInvoke(ws, "Range", "A1:" + lastCol + lastRow);
                }

                if (range == null)
                {
                    error = "无法定位图表数据区域";
                    return false;
                }

                if (!TrySetSourceData(chart, range))
                {
                    error = "无法绑定图表数据区域";
                    return false;
                }

                if (!TrimExtraSeries(chart, wantSeries, out error))
                {
                    return false;
                }

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
                RestoreExcel(excelApp);
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
                if (!TryParseType(fmt.ChartType, out int xl, out _, out error))
                {
                    return false;
                }

                try
                {
                    WppCom.TrySetProperty(chart, "ChartType", xl);
                }
                catch (Exception ex)
                {
                    error = "无法改 data-chart-type: " + ex.Message;
                    return false;
                }
            }

            if (fmt.Title != null)
            {
                TrySetTitle(chart, fmt.Title, warnings);
            }

            TrySetOptional(chart, fmt, warnings);
            return true;
        }

        public static bool TryApplyToShape(
            object shape,
            PptHtmlChartGrid grid,
            PptHtmlChartFormat format,
            bool hasTextNoise,
            List<string> warnings,
            out string error)
        {
            error = null;
            object chart = TryGetChart(shape);
            if (chart == null)
            {
                error = "目标形状不是图表";
                return false;
            }

            if (grid != null)
            {
                if (!TryPourGrid(chart, grid, out error))
                {
                    return false;
                }
            }
            else if (hasTextNoise)
            {
                warnings?.Add("忽略对 chart 的正文（请改内嵌 <table>）");
            }

            return TryApplyFormat(chart, format, warnings, out error);
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
            out string error)
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
                    shape = WppCom.Invoke(shapes, "AddChart2", -1, xlType, left, top, width, height);
                }
                catch (Exception)
                {
                    shape = WppCom.Invoke(shapes, "AddChart", xlType, left, top, width, height);
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
            if (TryReadGridFromWorkbook(chart, out grid, out string wbErr) && grid != null && grid.IsPourable)
            {
                return true;
            }

            if (TryReadGridFromSeries(chart, out grid, out string seriesErr) && grid != null && grid.IsPourable)
            {
                return true;
            }

            error = string.IsNullOrEmpty(seriesErr)
                ? (string.IsNullOrEmpty(wbErr) ? "无法读取图表数据" : wbErr)
                : seriesErr;
            return false;
        }

        private static bool TryReadGridFromWorkbook(object chart, out PptHtmlChartGrid grid, out string error)
        {
            grid = null;
            error = null;
            object excelApp = null;
            try
            {
                object chartData = WppCom.GetProperty(chart, "ChartData");
                if (chartData == null)
                {
                    error = "无 ChartData";
                    return false;
                }

                TryInvoke(chartData, "Activate");
                object workbook = WppCom.GetProperty(chartData, "Workbook");
                if (workbook == null)
                {
                    error = "无内嵌工作簿";
                    return false;
                }

                excelApp = WppCom.GetProperty(workbook, "Application");
                SuppressExcel(excelApp);
                object sheets = WppCom.GetProperty(workbook, "Worksheets");
                object ws = WppCom.GetIndexed(sheets, 1);
                object used = WppCom.GetProperty(ws, "UsedRange");
                if (used == null)
                {
                    error = "内嵌表为空";
                    return false;
                }

                object rowsObj = WppCom.GetProperty(used, "Rows");
                object colsObj = WppCom.GetProperty(used, "Columns");
                int rowCount = Convert.ToInt32(WppCom.GetProperty(rowsObj, "Count"));
                int colCount = Convert.ToInt32(WppCom.GetProperty(colsObj, "Count"));
                if (rowCount < 2 || colCount < 2)
                {
                    error = "内嵌表行列不足";
                    return false;
                }

                bool truncated = rowCount > MaxRows || colCount > MaxCols;
                int useRows = Math.Min(rowCount, MaxRows);
                int useCols = Math.Min(colCount, MaxCols);
                var columns = new List<PptHtmlChartColumn>();
                for (int c = 1; c <= useCols; c++)
                {
                    columns.Add(new PptHtmlChartColumn
                    {
                        Role = c == 1 ? "category" : "value",
                        Name = Convert.ToString(GetCell(ws, 1, c) ?? "")
                    });
                }

                var data = new List<List<string>>();
                for (int r = 2; r <= useRows; r++)
                {
                    var row = new List<string>();
                    for (int c = 1; c <= useCols; c++)
                    {
                        object v = GetCell(ws, r, c);
                        row.Add(FormatCell(v));
                    }

                    data.Add(row);
                }

                TryReadSeriesMeta(chart, columns);
                grid = new PptHtmlChartGrid
                {
                    Columns = columns,
                    Rows = data,
                    Truncated = truncated
                };
                return true;
            }
            catch (Exception ex)
            {
                error = "读 ChartData 失败: " + ex.Message;
                return false;
            }
            finally
            {
                RestoreExcel(excelApp);
            }
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

                object s1 = WppCom.GetIndexed(sc, 1);
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
                    object series = WppCom.GetIndexed(sc, i);
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

        private static void TryReadSeriesMeta(object chart, List<PptHtmlChartColumn> columns)
        {
            try
            {
                object sc = TryInvoke(chart, "SeriesCollection");
                int count = Convert.ToInt32(WppCom.GetProperty(sc, "Count"));
                int si = 0;
                for (int i = 0; i < columns.Count; i++)
                {
                    if (columns[i].Role == "category")
                    {
                        continue;
                    }

                    si++;
                    if (si > count)
                    {
                        break;
                    }

                    object series = WppCom.GetIndexed(sc, si);
                    if (string.IsNullOrEmpty(columns[i].Color))
                    {
                        columns[i].Color = TryReadSeriesColor(series);
                    }
                }
            }
            catch (Exception)
            {
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
                    WppCom.TrySetProperty(chart, "ChartStyle", style);
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
                    TryInvoke(fill, "Solid");
                    object fc = WppCom.GetProperty(fill, "ForeColor");
                    WppCom.TrySetProperty(fc, "RGB", plotRgb);
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

        private static void RestoreExcel(object excelApp)
        {
            if (excelApp == null)
            {
                return;
            }

            try
            {
                WppCom.TrySetProperty(excelApp, "DisplayAlerts", false);
            }
            catch (Exception)
            {
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

        private static bool TrimExtraSeries(object chart, int wantSeries, out string error)
        {
            error = null;
            if (wantSeries < 1)
            {
                return true;
            }

            try
            {
                for (int guard = 0; guard < 16; guard++)
                {
                    object sc = TryInvoke(chart, "SeriesCollection");
                    if (sc == null)
                    {
                        sc = WppCom.GetProperty(chart, "SeriesCollection");
                    }

                    int count = Convert.ToInt32(WppCom.GetProperty(sc, "Count"));
                    if (count <= wantSeries)
                    {
                        return true;
                    }

                    object last = WppCom.GetIndexed(sc, count);
                    if (last == null)
                    {
                        last = TryInvoke(chart, "SeriesCollection", count);
                    }

                    if (last == null)
                    {
                        error = "无法删除图表多余系列（柱子会对不齐分类）";
                        return false;
                    }

                    WppCom.Invoke(last, "Delete");
                }

                error = "无法把图表系列数收到 " + wantSeries;
                return false;
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
                object block = TryInvoke(ws, "Range", "A1", "H40");
                if (block == null)
                {
                    block = TryInvoke(ws, "Range", "A1:H40");
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

        private static void SetCell(object ws, int row, int col, object value)
        {
            object cells = WppCom.GetProperty(ws, "Cells");
            object cell;
            try
            {
                cell = cells.GetType().InvokeMember(
                    "Item",
                    System.Reflection.BindingFlags.GetProperty
                    | System.Reflection.BindingFlags.InvokeMethod
                    | System.Reflection.BindingFlags.Instance
                    | System.Reflection.BindingFlags.Public,
                    null,
                    cells,
                    new object[] { row, col });
            }
            catch (Exception)
            {
                cell = WppCom.Invoke(ws, "Cells", row, col);
            }

            WppCom.TrySetProperty(cell, "Value", value);
            WppCom.TrySetProperty(cell, "Value2", value);
        }

        private static object GetCell(object ws, int row, int col)
        {
            try
            {
                object cells = WppCom.GetProperty(ws, "Cells");
                object cell = cells.GetType().InvokeMember(
                    "Item",
                    System.Reflection.BindingFlags.GetProperty
                    | System.Reflection.BindingFlags.InvokeMethod
                    | System.Reflection.BindingFlags.Instance
                    | System.Reflection.BindingFlags.Public,
                    null,
                    cells,
                    new object[] { row, col });
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
