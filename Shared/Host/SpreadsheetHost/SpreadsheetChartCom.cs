using System;
using System.Collections.Generic;
using System.Globalization;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using Excel = Microsoft.Office.Interop.Excel;
using WordAddIn1.OpenFiles;

namespace WordAddIn1.SpreadsheetHost
{
    /// <summary>
    /// 将 ChartInfo 套到 Excel.Chart；套不上的字段进 warnings（I21）。
    /// </summary>
    internal static class SpreadsheetChartFormat
    {
        public static void ApplyToExcelChart(Excel.Chart chart, ChartInfo info, List<string> warnings)
        {
            if (chart == null || info == null)
            {
                return;
            }

            warnings = warnings ?? new List<string>();

            if (info.Width > 0 || info.Height > 0)
            {
                warnings.Add("Width/Height 已忽略（Excel 以 dest 格子矩形为准）");
            }

            if (!string.IsNullOrWhiteSpace(info.Type)
                && SpreadsheetChartParse.TryMapChartType(info.Type, out Excel.XlChartType xlType, out _))
            {
                try
                {
                    chart.ChartType = xlType;
                }
                catch (Exception ex)
                {
                    warnings.Add("Type: " + ex.Message);
                }
            }

            if (!string.IsNullOrWhiteSpace(info.Theme))
            {
                try
                {
                    chart.ChartStyle = ChartEngine.GetChartStyle(info.Theme);
                }
                catch (Exception ex)
                {
                    warnings.Add("Theme: " + ex.Message);
                }
            }

            if (!string.IsNullOrWhiteSpace(info.Title))
            {
                try
                {
                    chart.HasTitle = true;
                    chart.ChartTitle.Text = info.Title.Trim();
                    if (info.TitleFontSize > 0)
                    {
                        chart.ChartTitle.Format.TextFrame2.TextRange.Font.Size = info.TitleFontSize;
                    }

                    if (info.TitleFontBold)
                    {
                        try
                        {
                            chart.ChartTitle.Font.Bold = true;
                        }
                        catch (Exception)
                        {
                            warnings.Add("TitleFontBold 未应用");
                        }
                    }

                    if (!string.IsNullOrWhiteSpace(info.TitleFontColor))
                    {
                        TrySetTitleColor(chart, info.TitleFontColor, warnings);
                    }
                }
                catch (Exception ex)
                {
                    warnings.Add("Title: " + ex.Message);
                }
            }

            if (!string.IsNullOrWhiteSpace(info.LegendPosition))
            {
                try
                {
                    chart.HasLegend = true;
                    chart.Legend.Position = (Excel.XlLegendPosition)ChartEngine.GetLegendPosition(info.LegendPosition);
                }
                catch (Exception ex)
                {
                    warnings.Add("LegendPosition: " + ex.Message);
                }
            }

            if (!string.IsNullOrWhiteSpace(info.Gridlines))
            {
                TrySetGridlines(chart, info.Gridlines, warnings);
            }

            TryApplyAxes(chart, info, warnings);
            TryApplySeriesFormatting(chart, info, warnings);
            TryApplyChartSpecific(chart, info, warnings);
        }

        private static void TrySetTitleColor(Excel.Chart chart, string hex, List<string> warnings)
        {
            try
            {
                if (!TryParseRgb(hex, out int rgb))
                {
                    warnings.Add("TitleFontColor 无效");
                    return;
                }

                chart.ChartTitle.Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = rgb;
            }
            catch (Exception ex)
            {
                warnings.Add("TitleFontColor: " + ex.Message);
            }
        }

        private static void TrySetGridlines(Excel.Chart chart, string gridlines, List<string> warnings)
        {
            try
            {
                Excel.Axis valueAxis = (Excel.Axis)chart.Axes(Excel.XlAxisType.xlValue, Excel.XlAxisGroup.xlPrimary);
                string g = (gridlines ?? "").Trim().ToLowerInvariant();
                valueAxis.HasMajorGridlines = g == "major" || g == "both";
                valueAxis.HasMinorGridlines = g == "minor" || g == "both";
                if (g == "none")
                {
                    valueAxis.HasMajorGridlines = false;
                    valueAxis.HasMinorGridlines = false;
                }
            }
            catch (Exception ex)
            {
                warnings.Add("Gridlines: " + ex.Message);
            }
        }

        private static void TryApplyAxes(Excel.Chart chart, ChartInfo info, List<string> warnings)
        {
            if (info.AxisX != null && !string.IsNullOrWhiteSpace(info.AxisX.Title))
            {
                try
                {
                    Excel.Axis x = (Excel.Axis)chart.Axes(Excel.XlAxisType.xlCategory, Excel.XlAxisGroup.xlPrimary);
                    x.HasTitle = true;
                    x.AxisTitle.Text = info.AxisX.Title.Trim();
                }
                catch (Exception ex)
                {
                    warnings.Add("AxisX.Title: " + ex.Message);
                }
            }

            if (info.AxisY != null && !string.IsNullOrWhiteSpace(info.AxisY.PrimaryTitle))
            {
                try
                {
                    Excel.Axis y = (Excel.Axis)chart.Axes(Excel.XlAxisType.xlValue, Excel.XlAxisGroup.xlPrimary);
                    y.HasTitle = true;
                    y.AxisTitle.Text = info.AxisY.PrimaryTitle.Trim();
                }
                catch (Exception ex)
                {
                    warnings.Add("AxisY.PrimaryTitle: " + ex.Message);
                }
            }

            if (info.YAxisMin != 0 || info.YAxisMax != 0 || info.YAxisMajorUnit != 0)
            {
                try
                {
                    Excel.Axis y = (Excel.Axis)chart.Axes(Excel.XlAxisType.xlValue, Excel.XlAxisGroup.xlPrimary);
                    if (info.YAxisMin != 0 || info.YAxisMax != 0)
                    {
                        y.MinimumScaleIsAuto = false;
                        y.MaximumScaleIsAuto = false;
                        if (info.YAxisMin != 0)
                        {
                            y.MinimumScale = info.YAxisMin;
                        }

                        if (info.YAxisMax != 0)
                        {
                            y.MaximumScale = info.YAxisMax;
                        }
                    }

                    if (info.YAxisMajorUnit != 0)
                    {
                        y.MajorUnitIsAuto = false;
                        y.MajorUnit = info.YAxisMajorUnit;
                    }
                }
                catch (Exception ex)
                {
                    warnings.Add("YAxis scale: " + ex.Message);
                }
            }
        }

        private static void TryApplySeriesFormatting(Excel.Chart chart, ChartInfo info, List<string> warnings)
        {
            if (info.SeriesCollection == null || info.SeriesCollection.Count == 0)
            {
                if (!string.IsNullOrWhiteSpace(info.PlotColor))
                {
                    TryColorSeries(chart, 1, info.PlotColor, warnings);
                }

                return;
            }

            for (int i = 0; i < info.SeriesCollection.Count; i++)
            {
                SeriesInfo s = info.SeriesCollection[i];
                if (!string.IsNullOrWhiteSpace(s.Color))
                {
                    TryColorSeries(chart, i + 1, s.Color, warnings);
                }
            }
        }

        private static void TryColorSeries(Excel.Chart chart, int index1Based, string hex, List<string> warnings)
        {
            try
            {
                if (!TryParseRgb(hex, out int rgb))
                {
                    warnings.Add("系列色无效: " + hex);
                    return;
                }

                Excel.Series series = (Excel.Series)chart.SeriesCollection(index1Based);
                series.Format.Fill.ForeColor.RGB = rgb;
            }
            catch (Exception ex)
            {
                warnings.Add("Series color: " + ex.Message);
            }
        }

        private static void TryApplyChartSpecific(Excel.Chart chart, ChartInfo info, List<string> warnings)
        {
            try
            {
                if (info.ShowDataLabels || !string.IsNullOrWhiteSpace(info.DataLabelType))
                {
                    Excel.SeriesCollection sc = (Excel.SeriesCollection)chart.SeriesCollection();
                    for (int i = 1; i <= sc.Count; i++)
                    {
                        Excel.Series s = sc.Item(i);
                        s.HasDataLabels = info.ShowDataLabels
                            || !string.IsNullOrWhiteSpace(info.DataLabelType);
                        if (string.Equals(info.DataLabelType, "percent", StringComparison.OrdinalIgnoreCase))
                        {
                            Excel.DataLabels labels = (Excel.DataLabels)s.DataLabels();
                            labels.ShowPercentage = true;
                            labels.ShowValue = false;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                warnings.Add("DataLabels: " + ex.Message);
            }

            try
            {
                if (info.GapWidth > 0)
                {
                    Excel.ChartGroup group = (Excel.ChartGroup)chart.ChartGroups(1);
                    group.GapWidth = info.GapWidth;
                }
            }
            catch (Exception ex)
            {
                warnings.Add("GapWidth: " + ex.Message);
            }

            try
            {
                if (info.DataMarkers || info.MarkerSize > 0 || info.LineWeight > 0)
                {
                    Excel.SeriesCollection sc = (Excel.SeriesCollection)chart.SeriesCollection();
                    for (int i = 1; i <= sc.Count; i++)
                    {
                        Excel.Series s = sc.Item(i);
                        if (info.DataMarkers || info.MarkerSize > 0)
                        {
                            s.MarkerStyle = Excel.XlMarkerStyle.xlMarkerStyleAutomatic;
                            if (info.MarkerSize > 0)
                            {
                                s.MarkerSize = info.MarkerSize;
                            }
                        }

                        if (info.LineWeight > 0)
                        {
                            s.Format.Line.Weight = (float)info.LineWeight;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                warnings.Add("Markers/LineWeight: " + ex.Message);
            }

            try
            {
                if (info.Explosion > 0)
                {
                    Excel.Series s = (Excel.Series)chart.SeriesCollection(1);
                    s.Explosion = info.Explosion;
                }
            }
            catch (Exception ex)
            {
                warnings.Add("Explosion: " + ex.Message);
            }
        }

        public static bool TryParseRgb(string hex, out int rgb)
        {
            rgb = 0;
            if (string.IsNullOrWhiteSpace(hex))
            {
                return false;
            }

            string h = hex.Trim();
            if (h.StartsWith("#", StringComparison.Ordinal))
            {
                h = h.Substring(1);
            }

            if (h.Length != 6)
            {
                return false;
            }

            if (!int.TryParse(h.Substring(0, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out int r)
                || !int.TryParse(h.Substring(2, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out int g)
                || !int.TryParse(h.Substring(4, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out int b))
            {
                return false;
            }

            rgb = r + (g << 8) + (b << 16);
            return true;
        }

        public static ChartInfo ReadFormatFromExcelChart(Excel.Chart chart)
        {
            var info = new ChartInfo
            {
                SeriesCollection = new List<SeriesInfo>()
            };
            if (chart == null)
            {
                return info;
            }

            try
            {
                info.Type = SpreadsheetChartParse.ChartTypeToWire(chart.ChartType);
            }
            catch (Exception)
            {
            }

            try
            {
                if (chart.HasTitle)
                {
                    info.Title = chart.ChartTitle.Text;
                }
            }
            catch (Exception)
            {
            }

            try
            {
                if (chart.HasLegend)
                {
                    switch (chart.Legend.Position)
                    {
                        case Excel.XlLegendPosition.xlLegendPositionTop:
                            info.LegendPosition = "top";
                            break;
                        case Excel.XlLegendPosition.xlLegendPositionLeft:
                            info.LegendPosition = "left";
                            break;
                        case Excel.XlLegendPosition.xlLegendPositionRight:
                            info.LegendPosition = "right";
                            break;
                        default:
                            info.LegendPosition = "bottom";
                            break;
                    }
                }
            }
            catch (Exception)
            {
            }

            return info;
        }
    }

    internal static class SpreadsheetChartCom
    {
        private static readonly Regex SeriesFormulaRegex = new Regex(
            @"^=SERIES\((?<name>[^,]*),(?<cats>[^,]*),(?<vals>[^,]*),",
            RegexOptions.IgnoreCase | RegexOptions.Compiled);

        private static string FormatError(Exception ex)
        {
            Exception cur = ex;
            while (cur is TargetInvocationException tie && tie.InnerException != null)
            {
                cur = tie.InnerException;
            }

            string msg = cur?.Message ?? "未知错误";
            if (cur is COMException com)
            {
                msg += " (HRESULT=0x" + unchecked((uint)com.ErrorCode).ToString("X8") + ")";
            }

            return msg;
        }

        public static bool TryExecuteOnExcelWorkbook(
            Excel.Workbook book,
            SpreadsheetChartRequest request,
            out SpreadsheetChartResult partial,
            out string error)
        {
            partial = new SpreadsheetChartResult
            {
                Action = request.Action,
                Charts = new List<SpreadsheetChartInfo>(),
                Warnings = new List<string>()
            };
            error = null;

            try
            {
                switch (request.Action)
                {
                    case "list":
                        return TryListExcel(book, request.ListSheetFilter, partial, out error);
                    case "create":
                        return TryCreateExcel(book, request, partial, out error);
                    case "extract":
                        return TryExtractExcel(book, request, partial, out error);
                    case "apply":
                        return TryApplyExcel(book, request, partial, out error);
                    case "delete":
                        return TryDeleteExcel(book, request, partial, out error);
                    default:
                        error = "未知 action";
                        return false;
                }
            }
            catch (Exception ex)
            {
                error = "图表操作失败: " + FormatError(ex);
                return false;
            }
        }

        public static bool TryExecuteOnEtWorkbook(
            object book,
            SpreadsheetChartRequest request,
            out SpreadsheetChartResult partial,
            out string error)
        {
            partial = new SpreadsheetChartResult
            {
                Action = request.Action,
                Charts = new List<SpreadsheetChartInfo>(),
                Warnings = new List<string>()
            };
            error = null;

            try
            {
                switch (request.Action)
                {
                    case "list":
                        return TryListEt(book, request.ListSheetFilter, partial, out error);
                    case "create":
                        return TryCreateEt(book, request, partial, out error);
                    case "extract":
                        return TryExtractEt(book, request, partial, out error);
                    case "apply":
                        return TryApplyEt(book, request, partial, out error);
                    case "delete":
                        return TryDeleteEt(book, request, partial, out error);
                    default:
                        error = "未知 action";
                        return false;
                }
            }
            catch (Exception ex)
            {
                error = "图表操作失败: " + FormatError(ex);
                return false;
            }
        }

        private static bool TryListExcel(
            Excel.Workbook book,
            string sheetFilter,
            SpreadsheetChartResult partial,
            out string error)
        {
            error = null;
            foreach (object raw in book.Sheets)
            {
                var sheet = raw as Excel.Worksheet;
                if (sheet == null)
                {
                    continue;
                }

                string sheetName;
                try
                {
                    sheetName = sheet.Name;
                }
                catch (Exception)
                {
                    continue;
                }

                if (!string.IsNullOrEmpty(sheetFilter)
                    && !string.Equals(sheetName, sheetFilter, StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                Excel.ChartObjects cos;
                try
                {
                    cos = (Excel.ChartObjects)sheet.ChartObjects();
                }
                catch (Exception)
                {
                    continue;
                }

                for (int i = 1; i <= cos.Count; i++)
                {
                    Excel.ChartObject co = cos.Item(i);
                    if (IsPivotChart(co.Chart))
                    {
                        continue;
                    }

                    partial.Charts.Add(ReadChartInfoExcel(sheetName, co));
                }
            }

            partial.ChartsCount = partial.Charts.Count;
            return true;
        }

        private static SpreadsheetChartInfo ReadChartInfoExcel(string sheetName, Excel.ChartObject co)
        {
            var info = new SpreadsheetChartInfo
            {
                Sheet = sheetName,
                ChartName = co.Name
            };
            try
            {
                info.Type = SpreadsheetChartParse.ChartTypeToWire(co.Chart.ChartType);
            }
            catch (Exception)
            {
            }

            try
            {
                Excel.Range tl = co.TopLeftCell;
                Excel.Range br = co.BottomRightCell;
                info.DestRange = tl.Address[false, false]
                    + ":"
                    + br.Address[false, false];
            }
            catch (Exception)
            {
            }

            TryFillSourceFromSeries(co.Chart, info);
            return info;
        }

        private static void TryFillSourceFromSeries(Excel.Chart chart, SpreadsheetChartInfo info)
        {
            try
            {
                Excel.Series s = (Excel.Series)chart.SeriesCollection(1);
                string formula = s.Formula;
                Match m = SeriesFormulaRegex.Match(formula ?? "");
                if (!m.Success)
                {
                    return;
                }

                string vals = m.Groups["vals"].Value.Trim();
                if (TrySplitSheetRange(vals, out string sheet, out string range))
                {
                    info.SourceSheet = sheet;
                    info.SourceRange = range;
                }
            }
            catch (Exception)
            {
            }
        }

        private static bool TryCreateExcel(
            Excel.Workbook book,
            SpreadsheetChartRequest request,
            SpreadsheetChartResult partial,
            out string error)
        {
            if (!TryGetExcelWorksheet(book, request.DestSheet, out Excel.Worksheet destSheet, out error))
            {
                return false;
            }

            if (!A1Address.TryParseRange(
                    request.DestRangeA1,
                    out int fr,
                    out int fc,
                    out int lr,
                    out int lc,
                    out error))
            {
                return false;
            }

            if (!SpreadsheetChartLimits.TryCheckDestLimits(lr - fr + 1, lc - fc + 1, out error))
            {
                return false;
            }

            if (!string.IsNullOrEmpty(request.ChartName)
                && TryFindExcelChartByName(book, request.ChartName, out _, out _))
            {
                error = "已存在同名图表，请先 delete: " + request.ChartName;
                return false;
            }

            if (TryFindExcelChartIntersecting(destSheet, fr, fc, lr, lc, out string hitName))
            {
                error = "dest 与已有图表相交，请先 delete: " + hitName;
                return false;
            }

            if (!SpreadsheetChartParse.TryMapChartType(request.Format.Type, out Excel.XlChartType xlType, out error))
            {
                return false;
            }

            Excel.Range destRange = destSheet.Range[A1Address.Range(fr, fc, lr, lc)];
            Excel.ChartObjects cos = (Excel.ChartObjects)destSheet.ChartObjects();
            Excel.ChartObject co = cos.Add(
                Convert.ToDouble(destRange.Left),
                Convert.ToDouble(destRange.Top),
                Convert.ToDouble(destRange.Width),
                Convert.ToDouble(destRange.Height));
            Excel.Chart chart = co.Chart;

            try
            {
                if (IsPivotChart(chart))
                {
                    co.Delete();
                    error = "不支持 PivotChart";
                    return false;
                }

                chart.ChartType = xlType;
                if (!TryBindDataExcel(book, chart, request.DataBind, out error))
                {
                    try
                    {
                        co.Delete();
                    }
                    catch (Exception)
                    {
                    }

                    return false;
                }

                SpreadsheetChartFormat.ApplyToExcelChart(chart, request.Format, partial.Warnings);
                ApplyColumnColors(book, chart, request.DataBind, partial.Warnings);

                if (!string.IsNullOrEmpty(request.ChartName))
                {
                    try
                    {
                        co.Name = request.ChartName;
                    }
                    catch (Exception ex)
                    {
                        partial.Warnings.Add("设置图名失败: " + ex.Message);
                    }
                }

                FillResultFromExcel(partial, destSheet.Name, co);
                return true;
            }
            catch (Exception ex)
            {
                try
                {
                    co.Delete();
                }
                catch (Exception)
                {
                }

                error = "create 失败: " + FormatError(ex);
                return false;
            }
        }

        private static bool TryBindDataExcel(
            Excel.Workbook book,
            Excel.Chart chart,
            SpreadsheetChartDataBind bind,
            out string error)
        {
            error = null;
            if (bind == null)
            {
                error = "缺少数据绑定";
                return false;
            }

            if (bind.Mode == "source")
            {
                if (!TryGetExcelWorksheet(book, bind.SourceSheet, out Excel.Worksheet srcSheet, out error))
                {
                    return false;
                }

                Excel.Range src = srcSheet.Range[bind.SourceRangeA1];
                Excel.XlRowCol plotBy = string.Equals(bind.PlotBy, "rows", StringComparison.OrdinalIgnoreCase)
                    ? Excel.XlRowCol.xlRows
                    : Excel.XlRowCol.xlColumns;
                chart.SetSourceData(src, plotBy);
                return true;
            }

            // header mode
            while (chart.SeriesCollection().Count > 0)
            {
                ((Excel.Series)chart.SeriesCollection(1)).Delete();
            }

            SpreadsheetChartColumnBind category = null;
            foreach (SpreadsheetChartColumnBind col in bind.Columns)
            {
                if (col.Type == "category")
                {
                    category = col;
                    break;
                }
            }

            foreach (SpreadsheetChartColumnBind col in bind.Columns)
            {
                if (col.Type != "value")
                {
                    continue;
                }

                if (!TryGetExcelWorksheet(book, col.Sheet, out Excel.Worksheet vs, out error))
                {
                    return false;
                }

                Excel.Series series = (Excel.Series)chart.SeriesCollection().NewSeries();
                series.Values = vs.Range[col.RangeA1];
                if (!string.IsNullOrEmpty(col.Name))
                {
                    series.Name = col.Name;
                }

                if (category != null)
                {
                    if (!TryGetExcelWorksheet(book, category.Sheet, out Excel.Worksheet cs, out error))
                    {
                        return false;
                    }

                    series.XValues = cs.Range[category.RangeA1];
                }
            }

            return true;
        }

        private static void ApplyColumnColors(
            Excel.Workbook book,
            Excel.Chart chart,
            SpreadsheetChartDataBind bind,
            List<string> warnings)
        {
            if (bind?.Mode != "header" || bind.Columns == null)
            {
                return;
            }

            int seriesIndex = 1;
            foreach (SpreadsheetChartColumnBind col in bind.Columns)
            {
                if (col.Type != "value")
                {
                    continue;
                }

                if (!string.IsNullOrWhiteSpace(col.Color))
                {
                    try
                    {
                        if (SpreadsheetChartFormat.TryParseRgb(col.Color, out int rgb))
                        {
                            ((Excel.Series)chart.SeriesCollection(seriesIndex)).Format.Fill.ForeColor.RGB = rgb;
                        }
                        else
                        {
                            warnings.Add("Column color 无效: " + col.Color);
                        }
                    }
                    catch (Exception ex)
                    {
                        warnings.Add("Column color: " + ex.Message);
                    }
                }

                seriesIndex++;
            }
        }

        private static bool TryExtractExcel(
            Excel.Workbook book,
            SpreadsheetChartRequest request,
            SpreadsheetChartResult partial,
            out string error)
        {
            if (!TryResolveExcelChart(book, request, out Excel.Worksheet sheet, out Excel.ChartObject co, out error))
            {
                return false;
            }

            if (IsPivotChart(co.Chart))
            {
                error = "不支持 PivotChart";
                return false;
            }

            ChartInfo format = SpreadsheetChartFormat.ReadFormatFromExcelChart(co.Chart);
            SpreadsheetChartDataBind dataBind = TryReadDataBindFromExcel(co.Chart, sheet.Name);
            partial.Scope = request.Scope;
            partial.XmlContent = SpreadsheetChartParse.BuildExtractXml(
                format,
                dataBind,
                request.Scope,
                co.Name);
            FillResultFromExcel(partial, sheet.Name, co);
            return true;
        }

        private static SpreadsheetChartDataBind TryReadDataBindFromExcel(Excel.Chart chart, string fallbackSheet)
        {
            var columns = new List<SpreadsheetChartColumnBind>();
            try
            {
                Excel.SeriesCollection sc = (Excel.SeriesCollection)chart.SeriesCollection();
                for (int i = 1; i <= sc.Count; i++)
                {
                    Excel.Series s = sc.Item(i);
                    string formula = s.Formula ?? "";
                    Match m = SeriesFormulaRegex.Match(formula);
                    if (!m.Success)
                    {
                        continue;
                    }

                    if (i == 1)
                    {
                        string cats = m.Groups["cats"].Value.Trim();
                        if (TrySplitSheetRange(cats, out string cSheet, out string cRange))
                        {
                            columns.Add(new SpreadsheetChartColumnBind
                            {
                                Type = "category",
                                Sheet = cSheet,
                                RangeA1 = cRange,
                                Name = "类别"
                            });
                        }
                    }

                    string vals = m.Groups["vals"].Value.Trim();
                    if (TrySplitSheetRange(vals, out string vSheet, out string vRange))
                    {
                        columns.Add(new SpreadsheetChartColumnBind
                        {
                            Type = "value",
                            Sheet = vSheet,
                            RangeA1 = vRange,
                            Name = s.Name
                        });
                    }
                }
            }
            catch (Exception)
            {
            }

            if (columns.Count == 0)
            {
                return null;
            }

            return new SpreadsheetChartDataBind
            {
                Mode = "header",
                Columns = columns,
                PlotBy = "columns"
            };
        }

        private static bool TryApplyExcel(
            Excel.Workbook book,
            SpreadsheetChartRequest request,
            SpreadsheetChartResult partial,
            out string error)
        {
            if (!TryResolveExcelChart(book, request, out Excel.Worksheet sheet, out Excel.ChartObject co, out error))
            {
                return false;
            }

            if (IsPivotChart(co.Chart))
            {
                error = "不支持 PivotChart";
                return false;
            }

            Excel.Chart chart = co.Chart;
            string scope = request.Scope;

            if (scope == "data" || scope == "all")
            {
                int current = 0;
                try
                {
                    current = chart.SeriesCollection().Count;
                }
                catch (Exception)
                {
                }

                int expected = SpreadsheetChartParse.ExpectedSeriesCount(request.DataBind);
                if (expected >= 0 && current != expected)
                {
                    error = "系列数不一致（现有 "
                        + current
                        + "，请求 "
                        + expected
                        + "），请先 delete 再 create";
                    return false;
                }

                if (!TryBindDataExcel(book, chart, request.DataBind, out error))
                {
                    return false;
                }

                ApplyColumnColors(book, chart, request.DataBind, partial.Warnings);
            }

            if (scope == "format" || scope == "all")
            {
                SpreadsheetChartFormat.ApplyToExcelChart(chart, request.Format, partial.Warnings);
            }

            FillResultFromExcel(partial, sheet.Name, co);
            return true;
        }

        private static bool TryDeleteExcel(
            Excel.Workbook book,
            SpreadsheetChartRequest request,
            SpreadsheetChartResult partial,
            out string error)
        {
            if (!TryResolveExcelChart(book, request, out Excel.Worksheet sheet, out Excel.ChartObject co, out error))
            {
                return false;
            }

            FillResultFromExcel(partial, sheet.Name, co);
            co.Delete();
            return true;
        }

        private static void FillResultFromExcel(
            SpreadsheetChartResult partial,
            string sheetName,
            Excel.ChartObject co)
        {
            partial.ChartName = co.Name;
            partial.Sheet = sheetName;
            try
            {
                partial.Type = SpreadsheetChartParse.ChartTypeToWire(co.Chart.ChartType);
            }
            catch (Exception)
            {
            }

            try
            {
                Excel.Range tl = co.TopLeftCell;
                Excel.Range br = co.BottomRightCell;
                partial.DestRange = tl.Address[false, false] + ":" + br.Address[false, false];
            }
            catch (Exception)
            {
            }
        }

        private static bool TryResolveExcelChart(
            Excel.Workbook book,
            SpreadsheetChartRequest request,
            out Excel.Worksheet sheet,
            out Excel.ChartObject chartObject,
            out string error)
        {
            sheet = null;
            chartObject = null;
            error = null;

            if (!string.IsNullOrEmpty(request.ChartName))
            {
                if (TryFindExcelChartByName(book, request.ChartName, out sheet, out chartObject))
                {
                    return true;
                }

                error = "未找到图表: " + request.ChartName;
                return false;
            }

            if (!TryGetExcelWorksheet(book, request.DestSheet, out sheet, out error))
            {
                return false;
            }

            if (!A1Address.TryParseRange(
                    request.DestRangeA1,
                    out int fr,
                    out int fc,
                    out int lr,
                    out int lc,
                    out error))
            {
                return false;
            }

            var hits = new List<Excel.ChartObject>();
            Excel.ChartObjects cos = (Excel.ChartObjects)sheet.ChartObjects();
            for (int i = 1; i <= cos.Count; i++)
            {
                Excel.ChartObject co = cos.Item(i);
                if (ChartIntersects(co, fr, fc, lr, lc))
                {
                    hits.Add(co);
                }
            }

            if (hits.Count == 0)
            {
                error = "未找到与 dest 相交的图表";
                return false;
            }

            if (hits.Count > 1)
            {
                error = "dest 与多张图表相交，请用 chart_name 定位";
                return false;
            }

            chartObject = hits[0];
            return true;
        }

        private static bool TryFindExcelChartByName(
            Excel.Workbook book,
            string name,
            out Excel.Worksheet sheet,
            out Excel.ChartObject chartObject)
        {
            sheet = null;
            chartObject = null;
            foreach (object raw in book.Sheets)
            {
                var ws = raw as Excel.Worksheet;
                if (ws == null)
                {
                    continue;
                }

                try
                {
                    Excel.ChartObjects cos = (Excel.ChartObjects)ws.ChartObjects();
                    for (int i = 1; i <= cos.Count; i++)
                    {
                        Excel.ChartObject co = cos.Item(i);
                        if (string.Equals(co.Name, name, StringComparison.Ordinal))
                        {
                            sheet = ws;
                            chartObject = co;
                            return true;
                        }
                    }
                }
                catch (Exception)
                {
                }
            }

            return false;
        }

        private static bool TryFindExcelChartIntersecting(
            Excel.Worksheet sheet,
            int fr,
            int fc,
            int lr,
            int lc,
            out string hitName)
        {
            hitName = null;
            Excel.ChartObjects cos = (Excel.ChartObjects)sheet.ChartObjects();
            for (int i = 1; i <= cos.Count; i++)
            {
                Excel.ChartObject co = cos.Item(i);
                if (ChartIntersects(co, fr, fc, lr, lc))
                {
                    hitName = co.Name;
                    return true;
                }
            }

            return false;
        }

        private static bool ChartIntersects(Excel.ChartObject co, int fr, int fc, int lr, int lc)
        {
            try
            {
                Excel.Range tl = co.TopLeftCell;
                Excel.Range br = co.BottomRightCell;
                int a1 = tl.Row;
                int a2 = tl.Column;
                int b1 = br.Row;
                int b2 = br.Column;
                return !(lr < a1 || fr > b1 || lc < a2 || fc > b2);
            }
            catch (Exception)
            {
                return false;
            }
        }

        private static bool IsPivotChart(Excel.Chart chart)
        {
            try
            {
                object pl = chart.PivotLayout;
                return pl != null;
            }
            catch (Exception)
            {
                return false;
            }
        }

        private static bool TryGetExcelWorksheet(
            Excel.Workbook book,
            string sheetName,
            out Excel.Worksheet sheet,
            out string error)
        {
            sheet = null;
            error = null;
            if (string.IsNullOrWhiteSpace(sheetName))
            {
                error = "须提供工作表名";
                return false;
            }

            foreach (object raw in book.Sheets)
            {
                string name;
                try
                {
                    name = ((dynamic)raw).Name;
                }
                catch (Exception)
                {
                    continue;
                }

                if (!string.Equals(name, sheetName.Trim(), StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                if (raw is Excel.Chart)
                {
                    error = "sheet_type=chart，请换工作表";
                    return false;
                }

                sheet = raw as Excel.Worksheet;
                if (sheet == null)
                {
                    error = "sheet_type=chart，请换工作表";
                    return false;
                }

                return true;
            }

            error = "未找到工作表: " + sheetName;
            return false;
        }

        private static bool TrySplitSheetRange(string refText, out string sheet, out string range)
        {
            sheet = null;
            range = null;
            if (string.IsNullOrWhiteSpace(refText))
            {
                return false;
            }

            string t = refText.Trim().Trim('(').Trim(')');
            int bang = t.LastIndexOf('!');
            if (bang <= 0)
            {
                return false;
            }

            sheet = t.Substring(0, bang).Trim('\'');
            range = t.Substring(bang + 1).Replace("$", "");
            return !string.IsNullOrEmpty(sheet) && !string.IsNullOrEmpty(range);
        }

        // ——— et late-bind ———

        private static bool TryListEt(
            object book,
            string sheetFilter,
            SpreadsheetChartResult partial,
            out string error)
        {
            error = null;
            foreach (object sheet in EtCom.EnumerateSheets(book))
            {
                if (!IsEtWorksheet(sheet))
                {
                    continue;
                }

                string sheetName = EtCom.TryReadName(sheet) ?? "";
                if (!string.IsNullOrEmpty(sheetFilter)
                    && !string.Equals(sheetName, sheetFilter, StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                object cos = TryGetChartObjects(sheet);
                if (cos == null)
                {
                    continue;
                }

                int count;
                try
                {
                    count = Convert.ToInt32(EtCom.GetProperty(cos, "Count"));
                }
                catch (Exception)
                {
                    continue;
                }

                for (int i = 1; i <= count; i++)
                {
                    object co = EtCom.GetIndexed(cos, i);
                    if (co == null)
                    {
                        continue;
                    }

                    partial.Charts.Add(ReadChartInfoEt(sheetName, co));
                }
            }

            partial.ChartsCount = partial.Charts.Count;
            return true;
        }

        private static SpreadsheetChartInfo ReadChartInfoEt(string sheetName, object co)
        {
            var info = new SpreadsheetChartInfo
            {
                Sheet = sheetName,
                ChartName = Convert.ToString(EtCom.GetProperty(co, "Name")) ?? ""
            };
            try
            {
                object chart = EtCom.GetProperty(co, "Chart");
                object typeObj = EtCom.GetProperty(chart, "ChartType");
                if (typeObj != null)
                {
                    info.Type = SpreadsheetChartParse.ChartTypeToWire(
                        (Excel.XlChartType)Convert.ToInt32(typeObj));
                }

                object tl = EtCom.GetProperty(co, "TopLeftCell");
                object br = EtCom.GetProperty(co, "BottomRightCell");
                string a = Convert.ToString(EtCom.Invoke(tl, "Address", false, false));
                string b = Convert.ToString(EtCom.Invoke(br, "Address", false, false));
                if (!string.IsNullOrEmpty(a) && !string.IsNullOrEmpty(b))
                {
                    info.DestRange = a + ":" + b;
                }
            }
            catch (Exception)
            {
            }

            return info;
        }

        private static bool TryCreateEt(
            object book,
            SpreadsheetChartRequest request,
            SpreadsheetChartResult partial,
            out string error)
        {
            if (!TryGetEtWorksheet(book, request.DestSheet, out object destSheet, out error))
            {
                return false;
            }

            if (!A1Address.TryParseRange(
                    request.DestRangeA1,
                    out int fr,
                    out int fc,
                    out int lr,
                    out int lc,
                    out error))
            {
                return false;
            }

            if (!SpreadsheetChartLimits.TryCheckDestLimits(lr - fr + 1, lc - fc + 1, out error))
            {
                return false;
            }

            if (!string.IsNullOrEmpty(request.ChartName)
                && TryFindEtChartByName(book, request.ChartName, out _, out _))
            {
                error = "已存在同名图表，请先 delete: " + request.ChartName;
                return false;
            }

            if (TryFindEtChartIntersecting(destSheet, fr, fc, lr, lc, out string hit))
            {
                error = "dest 与已有图表相交，请先 delete: " + hit;
                return false;
            }

            if (!SpreadsheetChartParse.TryMapChartType(request.Format.Type, out Excel.XlChartType xlType, out error))
            {
                return false;
            }

            object destRange = EtGetSheetRange(destSheet, A1Address.Range(fr, fc, lr, lc));
            double left = Convert.ToDouble(EtCom.GetProperty(destRange, "Left"));
            double top = Convert.ToDouble(EtCom.GetProperty(destRange, "Top"));
            double width = Convert.ToDouble(EtCom.GetProperty(destRange, "Width"));
            double height = Convert.ToDouble(EtCom.GetProperty(destRange, "Height"));

            object cos = TryGetChartObjects(destSheet);
            if (cos == null)
            {
                error = "无法访问 ChartObjects";
                return false;
            }

            object co = EtCom.Invoke(cos, "Add", left, top, width, height);
            if (co == null)
            {
                error = "ChartObjects.Add 失败";
                return false;
            }

            try
            {
                object chart = EtCom.GetProperty(co, "Chart");
                EtCom.TrySetProperty(chart, "ChartType", (int)xlType);
                if (!TryBindDataEt(book, chart, request.DataBind, out error))
                {
                    try
                    {
                        EtCom.Invoke(co, "Delete");
                    }
                    catch (Exception)
                    {
                    }

                    return false;
                }

                // et：尽量套常用格式；失败进 warnings
                TryApplyFormatEt(chart, request.Format, partial.Warnings);

                if (!string.IsNullOrEmpty(request.ChartName))
                {
                    try
                    {
                        EtCom.TrySetProperty(co, "Name", request.ChartName);
                    }
                    catch (Exception ex)
                    {
                        partial.Warnings.Add("设置图名失败: " + ex.Message);
                    }
                }

                FillResultFromEt(partial, EtCom.TryReadName(destSheet), co);
                return true;
            }
            catch (Exception ex)
            {
                try
                {
                    EtCom.Invoke(co, "Delete");
                }
                catch (Exception)
                {
                }

                error = "create 失败: " + FormatError(ex);
                return false;
            }
        }

        private static bool TryBindDataEt(
            object book,
            object chart,
            SpreadsheetChartDataBind bind,
            out string error)
        {
            error = null;
            if (bind == null)
            {
                error = "缺少数据绑定";
                return false;
            }

            if (bind.Mode == "source")
            {
                if (!TryGetEtWorksheet(book, bind.SourceSheet, out object srcSheet, out error))
                {
                    return false;
                }

                object src = EtGetSheetRange(srcSheet, bind.SourceRangeA1);
                int plotBy = string.Equals(bind.PlotBy, "rows", StringComparison.OrdinalIgnoreCase)
                    ? SpreadsheetChartParse.XlRows
                    : SpreadsheetChartParse.XlColumns;
                EtCom.Invoke(chart, "SetSourceData", src, plotBy);
                return true;
            }

            object seriesColl = EtCom.GetProperty(chart, "SeriesCollection");
            // clear existing
            try
            {
                while (Convert.ToInt32(EtCom.GetProperty(seriesColl, "Count")) > 0)
                {
                    object s0 = EtCom.GetIndexed(seriesColl, 1);
                    EtCom.Invoke(s0, "Delete");
                }
            }
            catch (Exception)
            {
            }

            SpreadsheetChartColumnBind category = null;
            foreach (SpreadsheetChartColumnBind col in bind.Columns)
            {
                if (col.Type == "category")
                {
                    category = col;
                    break;
                }
            }

            foreach (SpreadsheetChartColumnBind col in bind.Columns)
            {
                if (col.Type != "value")
                {
                    continue;
                }

                if (!TryGetEtWorksheet(book, col.Sheet, out object vs, out error))
                {
                    return false;
                }

                object series = EtCom.Invoke(seriesColl, "NewSeries");
                EtCom.TrySetProperty(series, "Values", EtGetSheetRange(vs, col.RangeA1));
                if (!string.IsNullOrEmpty(col.Name))
                {
                    EtCom.TrySetProperty(series, "Name", col.Name);
                }

                if (category != null)
                {
                    if (!TryGetEtWorksheet(book, category.Sheet, out object cs, out error))
                    {
                        return false;
                    }

                    EtCom.TrySetProperty(series, "XValues", EtGetSheetRange(cs, category.RangeA1));
                }
            }

            return true;
        }

        private static void TryApplyFormatEt(object chart, ChartInfo info, List<string> warnings)
        {
            if (info == null)
            {
                return;
            }

            if (info.Width > 0 || info.Height > 0)
            {
                warnings.Add("Width/Height 已忽略（Excel 以 dest 格子矩形为准）");
            }

            try
            {
                if (!string.IsNullOrWhiteSpace(info.Title))
                {
                    EtCom.TrySetProperty(chart, "HasTitle", true);
                    object title = EtCom.GetProperty(chart, "ChartTitle");
                    EtCom.TrySetProperty(title, "Text", info.Title.Trim());
                }
            }
            catch (Exception ex)
            {
                warnings.Add("Title: " + ex.Message);
            }

            try
            {
                if (!string.IsNullOrWhiteSpace(info.LegendPosition))
                {
                    EtCom.TrySetProperty(chart, "HasLegend", true);
                    object legend = EtCom.GetProperty(chart, "Legend");
                    EtCom.TrySetProperty(legend, "Position", ChartEngine.GetLegendPosition(info.LegendPosition));
                }
            }
            catch (Exception ex)
            {
                warnings.Add("LegendPosition: " + ex.Message);
            }

            try
            {
                if (!string.IsNullOrWhiteSpace(info.Theme))
                {
                    EtCom.TrySetProperty(chart, "ChartStyle", ChartEngine.GetChartStyle(info.Theme));
                }
            }
            catch (Exception ex)
            {
                warnings.Add("Theme: " + ex.Message);
            }
        }

        private static bool TryExtractEt(
            object book,
            SpreadsheetChartRequest request,
            SpreadsheetChartResult partial,
            out string error)
        {
            if (!TryResolveEtChart(book, request, out object sheet, out object co, out error))
            {
                return false;
            }

            object chart = EtCom.GetProperty(co, "Chart");
            var format = new ChartInfo { Type = "", SeriesCollection = new List<SeriesInfo>() };
            try
            {
                object typeObj = EtCom.GetProperty(chart, "ChartType");
                if (typeObj != null)
                {
                    format.Type = SpreadsheetChartParse.ChartTypeToWire(
                        (Excel.XlChartType)Convert.ToInt32(typeObj));
                }

                if (Convert.ToBoolean(EtCom.GetProperty(chart, "HasTitle")))
                {
                    object title = EtCom.GetProperty(chart, "ChartTitle");
                    format.Title = Convert.ToString(EtCom.GetProperty(title, "Text")) ?? "";
                }
            }
            catch (Exception)
            {
            }

            SpreadsheetChartDataBind dataBind = null;
            string name = Convert.ToString(EtCom.GetProperty(co, "Name")) ?? "";
            partial.Scope = request.Scope;
            partial.XmlContent = SpreadsheetChartParse.BuildExtractXml(format, dataBind, request.Scope, name);
            FillResultFromEt(partial, EtCom.TryReadName(sheet), co);
            return true;
        }

        private static bool TryApplyEt(
            object book,
            SpreadsheetChartRequest request,
            SpreadsheetChartResult partial,
            out string error)
        {
            if (!TryResolveEtChart(book, request, out object sheet, out object co, out error))
            {
                return false;
            }

            object chart = EtCom.GetProperty(co, "Chart");
            string scope = request.Scope;

            if (scope == "data" || scope == "all")
            {
                int current = 0;
                try
                {
                    object sc = EtCom.GetProperty(chart, "SeriesCollection");
                    current = Convert.ToInt32(EtCom.GetProperty(sc, "Count"));
                }
                catch (Exception)
                {
                }

                int expected = SpreadsheetChartParse.ExpectedSeriesCount(request.DataBind);
                if (expected >= 0 && current != expected)
                {
                    error = "系列数不一致（现有 "
                        + current
                        + "，请求 "
                        + expected
                        + "），请先 delete 再 create";
                    return false;
                }

                if (!TryBindDataEt(book, chart, request.DataBind, out error))
                {
                    return false;
                }
            }

            if (scope == "format" || scope == "all")
            {
                TryApplyFormatEt(chart, request.Format, partial.Warnings);
            }

            FillResultFromEt(partial, EtCom.TryReadName(sheet), co);
            return true;
        }

        private static bool TryDeleteEt(
            object book,
            SpreadsheetChartRequest request,
            SpreadsheetChartResult partial,
            out string error)
        {
            if (!TryResolveEtChart(book, request, out object sheet, out object co, out error))
            {
                return false;
            }

            FillResultFromEt(partial, EtCom.TryReadName(sheet), co);
            EtCom.Invoke(co, "Delete");
            return true;
        }

        private static void FillResultFromEt(SpreadsheetChartResult partial, string sheetName, object co)
        {
            partial.Sheet = sheetName ?? "";
            try
            {
                partial.ChartName = Convert.ToString(EtCom.GetProperty(co, "Name")) ?? "";
                object chart = EtCom.GetProperty(co, "Chart");
                object typeObj = EtCom.GetProperty(chart, "ChartType");
                if (typeObj != null)
                {
                    partial.Type = SpreadsheetChartParse.ChartTypeToWire(
                        (Excel.XlChartType)Convert.ToInt32(typeObj));
                }

                object tl = EtCom.GetProperty(co, "TopLeftCell");
                object br = EtCom.GetProperty(co, "BottomRightCell");
                string a = Convert.ToString(EtCom.Invoke(tl, "Address", false, false));
                string b = Convert.ToString(EtCom.Invoke(br, "Address", false, false));
                if (!string.IsNullOrEmpty(a) && !string.IsNullOrEmpty(b))
                {
                    partial.DestRange = a + ":" + b;
                }
            }
            catch (Exception)
            {
            }
        }

        private static bool TryResolveEtChart(
            object book,
            SpreadsheetChartRequest request,
            out object sheet,
            out object chartObject,
            out string error)
        {
            sheet = null;
            chartObject = null;
            error = null;

            if (!string.IsNullOrEmpty(request.ChartName))
            {
                if (TryFindEtChartByName(book, request.ChartName, out sheet, out chartObject))
                {
                    return true;
                }

                error = "未找到图表: " + request.ChartName;
                return false;
            }

            if (!TryGetEtWorksheet(book, request.DestSheet, out sheet, out error))
            {
                return false;
            }

            if (!A1Address.TryParseRange(
                    request.DestRangeA1,
                    out int fr,
                    out int fc,
                    out int lr,
                    out int lc,
                    out error))
            {
                return false;
            }

            var hits = new List<object>();
            object cos = TryGetChartObjects(sheet);
            if (cos == null)
            {
                error = "无法访问 ChartObjects";
                return false;
            }

            int count = Convert.ToInt32(EtCom.GetProperty(cos, "Count"));
            for (int i = 1; i <= count; i++)
            {
                object co = EtCom.GetIndexed(cos, i);
                if (co != null && ChartIntersectsEt(co, fr, fc, lr, lc))
                {
                    hits.Add(co);
                }
            }

            if (hits.Count == 0)
            {
                error = "未找到与 dest 相交的图表";
                return false;
            }

            if (hits.Count > 1)
            {
                error = "dest 与多张图表相交，请用 chart_name 定位";
                return false;
            }

            chartObject = hits[0];
            return true;
        }

        private static bool TryFindEtChartByName(
            object book,
            string name,
            out object sheet,
            out object chartObject)
        {
            sheet = null;
            chartObject = null;
            foreach (object candidate in EtCom.EnumerateSheets(book))
            {
                if (!IsEtWorksheet(candidate))
                {
                    continue;
                }

                object cos = TryGetChartObjects(candidate);
                if (cos == null)
                {
                    continue;
                }

                int count;
                try
                {
                    count = Convert.ToInt32(EtCom.GetProperty(cos, "Count"));
                }
                catch (Exception)
                {
                    continue;
                }

                for (int i = 1; i <= count; i++)
                {
                    object co = EtCom.GetIndexed(cos, i);
                    if (co == null)
                    {
                        continue;
                    }

                    string n = Convert.ToString(EtCom.GetProperty(co, "Name")) ?? "";
                    if (string.Equals(n, name, StringComparison.Ordinal))
                    {
                        sheet = candidate;
                        chartObject = co;
                        return true;
                    }
                }
            }

            return false;
        }

        private static bool TryFindEtChartIntersecting(
            object sheet,
            int fr,
            int fc,
            int lr,
            int lc,
            out string hitName)
        {
            hitName = null;
            object cos = TryGetChartObjects(sheet);
            if (cos == null)
            {
                return false;
            }

            int count = Convert.ToInt32(EtCom.GetProperty(cos, "Count"));
            for (int i = 1; i <= count; i++)
            {
                object co = EtCom.GetIndexed(cos, i);
                if (co != null && ChartIntersectsEt(co, fr, fc, lr, lc))
                {
                    hitName = Convert.ToString(EtCom.GetProperty(co, "Name")) ?? "";
                    return true;
                }
            }

            return false;
        }

        private static bool ChartIntersectsEt(object co, int fr, int fc, int lr, int lc)
        {
            try
            {
                object tl = EtCom.GetProperty(co, "TopLeftCell");
                object br = EtCom.GetProperty(co, "BottomRightCell");
                int a1 = Convert.ToInt32(EtCom.GetProperty(tl, "Row"));
                int a2 = Convert.ToInt32(EtCom.GetProperty(tl, "Column"));
                int b1 = Convert.ToInt32(EtCom.GetProperty(br, "Row"));
                int b2 = Convert.ToInt32(EtCom.GetProperty(br, "Column"));
                return !(lr < a1 || fr > b1 || lc < a2 || fc > b2);
            }
            catch (Exception)
            {
                return false;
            }
        }

        private static object TryGetChartObjects(object sheet)
        {
            try
            {
                return EtCom.Invoke(sheet, "ChartObjects");
            }
            catch (Exception)
            {
                try
                {
                    return EtCom.GetProperty(sheet, "ChartObjects");
                }
                catch (Exception)
                {
                    return null;
                }
            }
        }

        private static bool IsEtWorksheet(object sheet)
        {
            try
            {
                object typeObj = EtCom.GetProperty(sheet, "Type");
                if (typeObj != null && Convert.ToInt32(typeObj) == -4109)
                {
                    return false;
                }
            }
            catch (Exception)
            {
            }

            return true;
        }

        private static bool TryGetEtWorksheet(
            object book,
            string sheetName,
            out object sheet,
            out string error)
        {
            sheet = null;
            error = null;
            if (string.IsNullOrWhiteSpace(sheetName))
            {
                error = "须提供工作表名";
                return false;
            }

            foreach (object candidate in EtCom.EnumerateSheets(book))
            {
                string name = EtCom.TryReadName(candidate) ?? "";
                if (!string.Equals(name, sheetName.Trim(), StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                if (!IsEtWorksheet(candidate))
                {
                    error = "sheet_type=chart，请换工作表";
                    return false;
                }

                sheet = candidate;
                return true;
            }

            error = "未找到工作表: " + sheetName;
            return false;
        }

        private static object EtGetSheetRange(object sheet, string a1)
        {
            if (sheet == null || string.IsNullOrWhiteSpace(a1))
            {
                return null;
            }

            try
            {
                return sheet.GetType().InvokeMember(
                    "Range",
                    BindingFlags.GetProperty
                        | BindingFlags.InvokeMethod
                        | BindingFlags.Instance
                        | BindingFlags.Public,
                    null,
                    sheet,
                    new object[] { a1 });
            }
            catch (Exception)
            {
            }

            try
            {
                return EtCom.InvokeFlex(sheet, "Range", a1);
            }
            catch (Exception)
            {
            }

            return null;
        }
    }
}
