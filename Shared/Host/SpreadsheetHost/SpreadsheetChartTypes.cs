using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Web.Script.Serialization;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;

namespace WordAddIn1.SpreadsheetHost
{
    internal sealed class SpreadsheetChartRequest
    {
        /// <summary>list | create | extract | apply | delete</summary>
        public string Action { get; set; }

        public string DestSheet { get; set; }

        public string DestRangeA1 { get; set; }

        public string ChartName { get; set; }

        /// <summary>format | data | all</summary>
        public string Scope { get; set; }

        public string XmlContent { get; set; }

        /// <summary>list 可选：只列该表；空=整簿。</summary>
        public string ListSheetFilter { get; set; }

        public ChartInfo Format { get; set; }

        public SpreadsheetChartDataBind DataBind { get; set; }
    }

    internal sealed class SpreadsheetChartDataBind
    {
        /// <summary>source | header</summary>
        public string Mode { get; set; }

        public string SourceSheet { get; set; }

        public string SourceRangeA1 { get; set; }

        /// <summary>columns | rows</summary>
        public string PlotBy { get; set; }

        public List<SpreadsheetChartColumnBind> Columns { get; set; }
    }

    internal sealed class SpreadsheetChartColumnBind
    {
        /// <summary>category | value</summary>
        public string Type { get; set; }

        public string Sheet { get; set; }

        public string RangeA1 { get; set; }

        public string Name { get; set; }

        public string Color { get; set; }
    }

    internal sealed class SpreadsheetChartInfo
    {
        public string ChartName { get; set; }

        public string Sheet { get; set; }

        public string Type { get; set; }

        public string DestRange { get; set; }

        public string SourceSheet { get; set; }

        public string SourceRange { get; set; }
    }

    internal sealed class SpreadsheetChartResult
    {
        public string ChannelId { get; set; }

        public string Kind { get; set; }

        public string Action { get; set; }

        public string ChartName { get; set; }

        public string Sheet { get; set; }

        public string DestRange { get; set; }

        public string Type { get; set; }

        public string Scope { get; set; }

        public string XmlContent { get; set; }

        public List<string> Warnings { get; set; }

        public List<SpreadsheetChartInfo> Charts { get; set; }

        public int ChartsCount { get; set; }
    }

    internal static class SpreadsheetChartLimits
    {
        public const int DestMaxRows = 80;
        public const int DestMaxCols = 40;
        public const int DestMaxCells = 2000;

        public const int SourceMaxRows = 5000;
        public const int SourceMaxCols = 50;
        public const int SourceMaxCells = 100000;

        public static bool TryCheckDestLimits(int rowCount, int colCount, out string error)
        {
            error = null;
            if (rowCount < 1 || colCount < 1)
            {
                error = "dest.range 无效";
                return false;
            }

            if (rowCount * colCount < 2)
            {
                error = "dest.range 须为至少 2 格的矩形（如 E2:K18）";
                return false;
            }

            if (rowCount > DestMaxRows
                || colCount > DestMaxCols
                || (long)rowCount * colCount > DestMaxCells)
            {
                error = "dest 放置区域过大（最多 "
                    + DestMaxRows + "×"
                    + DestMaxCols + " / "
                    + DestMaxCells + " 格）";
                return false;
            }

            return true;
        }

        public static bool TryCheckSourceLimits(int rowCount, int colCount, out string error)
        {
            error = null;
            if (rowCount > SourceMaxRows
                || colCount > SourceMaxCols
                || (long)rowCount * colCount > SourceMaxCells)
            {
                error = "图表源区域过大（最多 "
                    + SourceMaxRows + "×"
                    + SourceMaxCols + " / "
                    + SourceMaxCells + " 格）";
                return false;
            }

            return true;
        }
    }

    internal static class SpreadsheetChartParse
    {
        public const int XlRows = 1;
        public const int XlColumns = 2;

        public static bool TryParseRequest(
            Dictionary<string, object> args,
            out SpreadsheetChartRequest request,
            out string error)
        {
            request = null;
            error = null;
            if (args == null)
            {
                error = "须提供参数";
                return false;
            }

            string action = GetString(args, "action").ToLowerInvariant();
            if (action != "list"
                && action != "create"
                && action != "extract"
                && action != "apply"
                && action != "delete")
            {
                error = "action 须为 list|create|extract|apply|delete";
                return false;
            }

            string chartName = GetString(args, "chart_name");
            if (string.IsNullOrEmpty(chartName))
            {
                chartName = null;
            }

            string listFilter = GetString(args, "sheet");
            string destSheet = null;
            string destRange = null;
            if (args.ContainsKey("dest") && args["dest"] != null)
            {
                Dictionary<string, object> dest = CoerceDict(args["dest"]);
                if (dest == null)
                {
                    error = "dest 须为对象 { sheet, range }";
                    return false;
                }

                destSheet = GetString(dest, "sheet");
                destRange = GetString(dest, "range");
                if (string.IsNullOrEmpty(listFilter) && !string.IsNullOrEmpty(destSheet))
                {
                    listFilter = destSheet;
                }
            }

            if (action == "list")
            {
                request = new SpreadsheetChartRequest
                {
                    Action = action,
                    ChartName = chartName,
                    ListSheetFilter = string.IsNullOrEmpty(listFilter) ? null : listFilter
                };
                return true;
            }

            string scope = GetString(args, "scope").ToLowerInvariant();
            string xml = GetString(args, "xml_content");
            if (string.IsNullOrEmpty(xml) && args.ContainsKey("xml_content") && args["xml_content"] != null)
            {
                xml = Convert.ToString(args["xml_content"]) ?? "";
            }

            if (action == "create")
            {
                if (string.IsNullOrEmpty(destSheet) || string.IsNullOrEmpty(destRange))
                {
                    error = "create 必须提供 dest.sheet 与 dest.range";
                    return false;
                }

                if (destSheet.IndexOf('!') >= 0 || destRange.IndexOf('!') >= 0)
                {
                    error = "dest 须 sheet / range 分开，禁止含 !";
                    return false;
                }

                if (string.IsNullOrWhiteSpace(xml))
                {
                    error = "create 必须提供 xml_content";
                    return false;
                }

                if (!TryParseExcelChartXml(
                        xml,
                        requireData: true,
                        out ChartInfo format,
                        out SpreadsheetChartDataBind dataBind,
                        out string xmlName,
                        out error))
                {
                    return false;
                }

                if (string.IsNullOrEmpty(chartName) && !string.IsNullOrEmpty(xmlName))
                {
                    chartName = xmlName;
                }

                request = new SpreadsheetChartRequest
                {
                    Action = action,
                    DestSheet = destSheet,
                    DestRangeA1 = destRange,
                    ChartName = chartName,
                    XmlContent = xml,
                    Format = format,
                    DataBind = dataBind
                };
                return true;
            }

            if (action == "extract")
            {
                if (string.IsNullOrEmpty(scope))
                {
                    scope = "all";
                }

                if (scope != "format" && scope != "data" && scope != "all")
                {
                    error = "scope 须为 format|data|all";
                    return false;
                }

                if (string.IsNullOrEmpty(chartName)
                    && (string.IsNullOrEmpty(destSheet) || string.IsNullOrEmpty(destRange)))
                {
                    error = "extract 须提供 chart_name，或 dest.sheet + dest.range";
                    return false;
                }

                request = new SpreadsheetChartRequest
                {
                    Action = action,
                    DestSheet = destSheet,
                    DestRangeA1 = destRange,
                    ChartName = chartName,
                    Scope = scope
                };
                return true;
            }

            if (action == "apply")
            {
                if (string.IsNullOrEmpty(scope)
                    || (scope != "format" && scope != "data" && scope != "all"))
                {
                    error = "apply 必须显式提供 scope=format|data|all";
                    return false;
                }

                if (string.IsNullOrEmpty(chartName)
                    && (string.IsNullOrEmpty(destSheet) || string.IsNullOrEmpty(destRange)))
                {
                    error = "apply 须提供 chart_name，或 dest.sheet + dest.range";
                    return false;
                }

                if (string.IsNullOrWhiteSpace(xml))
                {
                    error = "apply 必须提供 xml_content";
                    return false;
                }

                bool needData = scope == "data" || scope == "all";
                bool needFormat = scope == "format" || scope == "all";
                if (!TryParseExcelChartXml(
                        xml,
                        requireData: needData,
                        out ChartInfo format,
                        out SpreadsheetChartDataBind dataBind,
                        out string xmlName,
                        out error))
                {
                    return false;
                }

                if (!needFormat)
                {
                    format = null;
                }

                if (!needData)
                {
                    dataBind = null;
                }

                if (string.IsNullOrEmpty(chartName) && !string.IsNullOrEmpty(xmlName))
                {
                    chartName = xmlName;
                }

                request = new SpreadsheetChartRequest
                {
                    Action = action,
                    DestSheet = destSheet,
                    DestRangeA1 = destRange,
                    ChartName = chartName,
                    Scope = scope,
                    XmlContent = xml,
                    Format = format,
                    DataBind = dataBind
                };
                return true;
            }

            // delete
            if (string.IsNullOrEmpty(chartName)
                && (string.IsNullOrEmpty(destSheet) || string.IsNullOrEmpty(destRange)))
            {
                error = "delete 须提供 chart_name，或 dest.sheet + dest.range";
                return false;
            }

            request = new SpreadsheetChartRequest
            {
                Action = action,
                DestSheet = destSheet,
                DestRangeA1 = destRange,
                ChartName = chartName
            };
            return true;
        }

        public static bool TryParseExcelChartXml(
            string xmlContent,
            bool requireData,
            out ChartInfo format,
            out SpreadsheetChartDataBind dataBind,
            out string chartName,
            out string error)
        {
            format = null;
            dataBind = null;
            chartName = null;
            error = null;
            if (string.IsNullOrWhiteSpace(xmlContent))
            {
                error = "xml_content 为空";
                return false;
            }

            XDocument doc;
            try
            {
                doc = XDocument.Parse(xmlContent);
            }
            catch (Exception ex)
            {
                error = "xml_content 解析失败: " + ex.Message;
                return false;
            }

            XElement root = doc.Root;
            if (root == null || !string.Equals(root.Name.LocalName, "ChartConfig", StringComparison.OrdinalIgnoreCase))
            {
                error = "根元素须为 ChartConfig";
                return false;
            }

            XElement chartEl = root.Element("Chart")
                ?? root.Elements().FirstOrDefault(e =>
                    string.Equals(e.Name.LocalName, "Chart", StringComparison.OrdinalIgnoreCase));
            if (chartEl == null)
            {
                error = "缺少 Chart 元素";
                return false;
            }

            chartName = Attr(chartEl, "name");
            if (string.IsNullOrEmpty(chartName))
            {
                chartName = Attr(root, "name");
            }

            format = new ChartInfo
            {
                Type = El(chartEl, "Type") ?? "",
                Title = El(chartEl, "Title") ?? "",
                Theme = El(chartEl, "Theme") ?? "",
                Width = ParseInt(El(chartEl, "Width")),
                Height = ParseInt(El(chartEl, "Height")),
                ShowDataLabels = ParseBool(El(chartEl, "ShowDataLabels")),
                LegendPosition = El(chartEl, "LegendPosition") ?? "",
                PlotColor = El(chartEl, "PlotColor") ?? "",
                TitleFontSize = ParseInt(El(chartEl, "TitleFontSize")),
                TitleFontBold = ParseBool(El(chartEl, "TitleFontBold")),
                TitleFontColor = El(chartEl, "TitleFontColor") ?? "",
                Gridlines = El(chartEl, "Gridlines") ?? "",
                GapWidth = ParseInt(El(chartEl, "GapWidth")),
                DataMarkers = ParseBool(El(chartEl, "DataMarkers")),
                MarkerSize = ParseInt(El(chartEl, "MarkerSize")),
                LineWeight = ParseDouble(El(chartEl, "LineWeight")),
                YAxisMin = ParseDouble(El(chartEl, "YAxisMin")),
                YAxisMax = ParseDouble(El(chartEl, "YAxisMax")),
                YAxisMajorUnit = ParseDouble(El(chartEl, "YAxisMajorUnit")),
                DataLabelType = El(chartEl, "DataLabelType") ?? "",
                Explosion = ParseInt(El(chartEl, "Explosion")),
                FillMissingValues = true,
                SeriesCollection = new List<SeriesInfo>()
            };

            XElement fillMissing = Child(chartEl, "FillMissingValues");
            if (fillMissing != null)
            {
                format.FillMissingValues = ParseBool(fillMissing.Value);
            }

            XElement axisX = Child(chartEl, "AxisX");
            if (axisX != null)
            {
                format.AxisX = new AxisXInfo
                {
                    Title = AxisField(axisX, "Title"),
                    Type = AxisField(axisX, "Type") ?? "category",
                    Format = AxisField(axisX, "Format"),
                    TickLabelCount = ParseNullableInt(AxisField(axisX, "TickLabelCount")),
                    TickLabelSpacing = ParseNullableInt(AxisField(axisX, "TickLabelSpacing")),
                    BetweenCategories = ParseNullableBool(AxisField(axisX, "BetweenCategories"))
                };
            }

            XElement axisY = Child(chartEl, "AxisY");
            if (axisY != null)
            {
                format.AxisY = new AxisYInfo
                {
                    PrimaryTitle = AxisField(axisY, "PrimaryTitle"),
                    SecondaryTitle = AxisField(axisY, "SecondaryTitle")
                };
            }

            XElement dataTable = Child(chartEl, "DataTable");
            if (dataTable != null)
            {
                if (Child(dataTable, "Rows") != null || Child(dataTable, "DataFile") != null)
                {
                    error = "Excel 图禁止 Rows/DataFile 灌点值，请用 Source 或 Header/Column 单元格绑定";
                    return false;
                }

                XElement source = Child(dataTable, "Source");
                XElement header = Child(dataTable, "Header");
                if (source != null && header != null)
                {
                    error = "DataTable 中 Source 与 Header 只能二选一";
                    return false;
                }

                if (source != null)
                {
                    string sheet = Attr(source, "sheet");
                    string range = Attr(source, "range");
                    if (string.IsNullOrWhiteSpace(sheet) || string.IsNullOrWhiteSpace(range))
                    {
                        error = "Source 须带 sheet 与 range 属性";
                        return false;
                    }

                    if (sheet.IndexOf('!') >= 0 || range.IndexOf('!') >= 0)
                    {
                        error = "Source 须 sheet/range 分开，禁止含 !";
                        return false;
                    }

                    string plotBy = (Attr(source, "plotBy") ?? "columns").Trim().ToLowerInvariant();
                    if (plotBy != "columns" && plotBy != "rows")
                    {
                        error = "plotBy 须为 columns|rows";
                        return false;
                    }

                    if (!ValidateRangeLimits(range, out error))
                    {
                        return false;
                    }

                    dataBind = new SpreadsheetChartDataBind
                    {
                        Mode = "source",
                        SourceSheet = sheet.Trim(),
                        SourceRangeA1 = range.Trim(),
                        PlotBy = plotBy,
                        Columns = new List<SpreadsheetChartColumnBind>()
                    };
                }
                else if (header != null)
                {
                    var columns = new List<SpreadsheetChartColumnBind>();
                    foreach (XElement col in header.Elements()
                                 .Where(e => string.Equals(e.Name.LocalName, "Column", StringComparison.OrdinalIgnoreCase)))
                    {
                        string type = (Attr(col, "type") ?? "value").Trim().ToLowerInvariant();
                        if (type != "category" && type != "value")
                        {
                            error = "Column type 须为 category|value";
                            return false;
                        }

                        string sheet = Attr(col, "sheet");
                        string range = Attr(col, "range");
                        if (string.IsNullOrWhiteSpace(sheet) || string.IsNullOrWhiteSpace(range))
                        {
                            error = "Column 须带 sheet 与 range";
                            return false;
                        }

                        if (sheet.IndexOf('!') >= 0 || range.IndexOf('!') >= 0)
                        {
                            error = "Column 须 sheet/range 分开，禁止含 !";
                            return false;
                        }

                        if (!ValidateRangeLimits(range, out error))
                        {
                            return false;
                        }

                        columns.Add(new SpreadsheetChartColumnBind
                        {
                            Type = type,
                            Sheet = sheet.Trim(),
                            RangeA1 = range.Trim(),
                            Name = (col.Value ?? "").Trim(),
                            Color = Attr(col, "color")
                        });
                    }

                    if (columns.Count(c => c.Type == "category") > 1)
                    {
                        error = "Header 中 category 列至多一列";
                        return false;
                    }

                    if (columns.Count(c => c.Type == "value") < 1)
                    {
                        error = "Header 中至少需要一个 value 列";
                        return false;
                    }

                    dataBind = new SpreadsheetChartDataBind
                    {
                        Mode = "header",
                        PlotBy = "columns",
                        Columns = columns
                    };
                }
                else if (requireData)
                {
                    error = "DataTable 须含 Source 或 Header";
                    return false;
                }
            }
            else if (requireData)
            {
                error = "须提供 DataTable（Source 或 Header 单元格绑定）";
                return false;
            }

            if (requireData && dataBind == null)
            {
                error = "须提供有效的数据绑定";
                return false;
            }

            if (!string.IsNullOrWhiteSpace(format.Type)
                && !TryMapChartType(format.Type, out _, out string typeError))
            {
                error = typeError;
                return false;
            }

            if (requireData && string.IsNullOrWhiteSpace(format.Type))
            {
                error = "Chart/Type 不能为空";
                return false;
            }

            return true;
        }

        public static bool TryMapChartType(string type, out Excel.XlChartType xlType, out string error)
        {
            xlType = Excel.XlChartType.xlColumnClustered;
            error = null;
            string t = (type ?? "").Trim().ToLowerInvariant();
            switch (t)
            {
                case "column":
                case "column_clustered":
                    xlType = Excel.XlChartType.xlColumnClustered;
                    return true;
                case "line":
                    xlType = Excel.XlChartType.xlLine;
                    return true;
                case "pie":
                    xlType = Excel.XlChartType.xlPie;
                    return true;
                case "pie3d":
                    xlType = Excel.XlChartType.xl3DPie;
                    return true;
                default:
                    error = "未知图表类型: " + type + "（支持 column_clustered|line|pie|pie3d，column 为别名）";
                    return false;
            }
        }

        public static string ChartTypeToWire(Excel.XlChartType xlType)
        {
            switch (xlType)
            {
                case Excel.XlChartType.xlLine:
                    return "line";
                case Excel.XlChartType.xlPie:
                    return "pie";
                case Excel.XlChartType.xl3DPie:
                    return "pie3d";
                default:
                    return "column_clustered";
            }
        }

        public static int ExpectedSeriesCount(SpreadsheetChartDataBind bind)
        {
            if (bind == null)
            {
                return -1;
            }

            if (bind.Mode == "header" && bind.Columns != null)
            {
                return bind.Columns.Count(c => c.Type == "value");
            }

            if (bind.Mode == "source"
                && A1Address.TryParseRange(
                    bind.SourceRangeA1,
                    out int fr,
                    out int fc,
                    out int lr,
                    out int lc,
                    out _))
            {
                int rows = lr - fr + 1;
                int cols = lc - fc + 1;
                if (string.Equals(bind.PlotBy, "rows", StringComparison.OrdinalIgnoreCase))
                {
                    return Math.Max(0, rows - 1);
                }

                return Math.Max(0, cols - 1);
            }

            return -1;
        }

        public static Dictionary<string, object> ChartInfoToWire(SpreadsheetChartInfo info)
        {
            var d = new Dictionary<string, object>
            {
                ["chart_name"] = info.ChartName ?? "",
                ["sheet"] = info.Sheet ?? "",
                ["type"] = info.Type ?? "",
                ["dest_range"] = info.DestRange ?? ""
            };
            if (!string.IsNullOrEmpty(info.SourceSheet))
            {
                d["source_sheet"] = info.SourceSheet;
            }

            if (!string.IsNullOrEmpty(info.SourceRange))
            {
                d["source_range"] = info.SourceRange;
            }

            return d;
        }

        public static string BuildExtractXml(
            ChartInfo format,
            SpreadsheetChartDataBind dataBind,
            string scope,
            string chartName)
        {
            scope = (scope ?? "all").ToLowerInvariant();
            var chart = new XElement("Chart");
            if (!string.IsNullOrEmpty(chartName))
            {
                chart.SetAttributeValue("name", chartName);
            }

            if (scope == "format" || scope == "all")
            {
                if (format != null)
                {
                    if (!string.IsNullOrEmpty(format.Type))
                    {
                        chart.Add(new XElement("Type", format.Type));
                    }

                    if (!string.IsNullOrEmpty(format.Title))
                    {
                        chart.Add(new XElement("Title", format.Title));
                    }

                    if (!string.IsNullOrEmpty(format.Theme))
                    {
                        chart.Add(new XElement("Theme", format.Theme));
                    }

                    if (!string.IsNullOrEmpty(format.LegendPosition))
                    {
                        chart.Add(new XElement("LegendPosition", format.LegendPosition));
                    }

                    if (!string.IsNullOrEmpty(format.Gridlines))
                    {
                        chart.Add(new XElement("Gridlines", format.Gridlines));
                    }

                    if (format.ShowDataLabels)
                    {
                        chart.Add(new XElement("ShowDataLabels", "true"));
                    }

                    if (!string.IsNullOrEmpty(format.DataLabelType))
                    {
                        chart.Add(new XElement("DataLabelType", format.DataLabelType));
                    }

                    if (format.AxisX != null)
                    {
                        var ax = new XElement("AxisX");
                        if (!string.IsNullOrEmpty(format.AxisX.Title))
                        {
                            ax.SetAttributeValue("Title", format.AxisX.Title);
                        }

                        if (!string.IsNullOrEmpty(format.AxisX.Type))
                        {
                            ax.SetAttributeValue("Type", format.AxisX.Type);
                        }

                        chart.Add(ax);
                    }

                    if (format.AxisY != null
                        && (!string.IsNullOrEmpty(format.AxisY.PrimaryTitle)
                            || !string.IsNullOrEmpty(format.AxisY.SecondaryTitle)))
                    {
                        var ay = new XElement("AxisY");
                        if (!string.IsNullOrEmpty(format.AxisY.PrimaryTitle))
                        {
                            ay.SetAttributeValue("PrimaryTitle", format.AxisY.PrimaryTitle);
                        }

                        if (!string.IsNullOrEmpty(format.AxisY.SecondaryTitle))
                        {
                            ay.SetAttributeValue("SecondaryTitle", format.AxisY.SecondaryTitle);
                        }

                        chart.Add(ay);
                    }
                }
            }

            if ((scope == "data" || scope == "all") && dataBind != null)
            {
                var dataTable = new XElement("DataTable");
                if (dataBind.Mode == "source")
                {
                    var source = new XElement("Source");
                    source.SetAttributeValue("sheet", dataBind.SourceSheet ?? "");
                    source.SetAttributeValue("range", dataBind.SourceRangeA1 ?? "");
                    source.SetAttributeValue("plotBy", dataBind.PlotBy ?? "columns");
                    dataTable.Add(source);
                }
                else if (dataBind.Columns != null)
                {
                    var header = new XElement("Header");
                    foreach (SpreadsheetChartColumnBind col in dataBind.Columns)
                    {
                        var c = new XElement("Column", col.Name ?? "");
                        c.SetAttributeValue("type", col.Type ?? "value");
                        c.SetAttributeValue("sheet", col.Sheet ?? "");
                        c.SetAttributeValue("range", col.RangeA1 ?? "");
                        if (!string.IsNullOrEmpty(col.Color))
                        {
                            c.SetAttributeValue("color", col.Color);
                        }

                        header.Add(c);
                    }

                    dataTable.Add(header);
                }

                chart.Add(dataTable);
            }

            return new XElement("ChartConfig", chart).ToString(SaveOptions.DisableFormatting);
        }

        private static bool ValidateRangeLimits(string rangeA1, out string error)
        {
            error = null;
            if (!A1Address.TryParseRange(rangeA1, out int fr, out int fc, out int lr, out int lc, out error))
            {
                return false;
            }

            return SpreadsheetChartLimits.TryCheckSourceLimits(lr - fr + 1, lc - fc + 1, out error);
        }

        private static string El(XElement parent, string name)
        {
            return Child(parent, name)?.Value;
        }

        private static XElement Child(XElement parent, string name)
        {
            if (parent == null)
            {
                return null;
            }

            return parent.Elements().FirstOrDefault(e =>
                string.Equals(e.Name.LocalName, name, StringComparison.OrdinalIgnoreCase));
        }

        private static string Attr(XElement el, string name)
        {
            if (el == null)
            {
                return null;
            }

            XAttribute a = el.Attributes().FirstOrDefault(x =>
                string.Equals(x.Name.LocalName, name, StringComparison.OrdinalIgnoreCase));
            return a?.Value;
        }

        private static string AxisField(XElement axis, string name)
        {
            string fromAttr = Attr(axis, name);
            if (!string.IsNullOrEmpty(fromAttr))
            {
                return fromAttr;
            }

            return El(axis, name);
        }

        private static int ParseInt(string s)
        {
            return int.TryParse(s, NumberStyles.Integer, CultureInfo.InvariantCulture, out int v) ? v : 0;
        }

        private static double ParseDouble(string s)
        {
            return double.TryParse(s, NumberStyles.Float, CultureInfo.InvariantCulture, out double v) ? v : 0;
        }

        private static bool ParseBool(string s)
        {
            return bool.TryParse(s, out bool v) && v;
        }

        private static int? ParseNullableInt(string s)
        {
            if (string.IsNullOrWhiteSpace(s))
            {
                return null;
            }

            return int.TryParse(s, NumberStyles.Integer, CultureInfo.InvariantCulture, out int v)
                ? (int?)v
                : null;
        }

        private static bool? ParseNullableBool(string s)
        {
            if (string.IsNullOrWhiteSpace(s))
            {
                return null;
            }

            return bool.TryParse(s, out bool v) ? (bool?)v : null;
        }

        private static string GetString(Dictionary<string, object> args, string key)
        {
            if (args == null || !args.ContainsKey(key) || args[key] == null)
            {
                return "";
            }

            return Convert.ToString(args[key])?.Trim() ?? "";
        }

        private static Dictionary<string, object> CoerceDict(object raw)
        {
            if (raw is Dictionary<string, object> d)
            {
                return d;
            }

            if (raw is IDictionary idict)
            {
                var result = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);
                foreach (DictionaryEntry entry in idict)
                {
                    result[Convert.ToString(entry.Key) ?? ""] = entry.Value;
                }

                return result;
            }

            try
            {
                var ser = new JavaScriptSerializer();
                return ser.ConvertToType<Dictionary<string, object>>(raw);
            }
            catch (Exception)
            {
                return null;
            }
        }
    }
}
