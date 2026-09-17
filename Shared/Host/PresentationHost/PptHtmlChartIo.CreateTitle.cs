using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using System.Xml.Linq;
using WordAddIn1.OpenFiles;


namespace WordAddIn1.PresentationHost
{
    internal static partial class PptHtmlChartIo
    {
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
            bool newLayout = true,
            bool applyHtmlChrome = true)
        {
            shape = null;
            error = null;
            if (grid == null || !grid.IsPourable)
            {
                error = "新建 chart 必须在节点内嵌 <table> 灌数，不能建空图";
                return false;
            }

            if (chartStyle < 0
                && format != null
                && int.TryParse(format.ChartStyle, NumberStyles.Integer, CultureInfo.InvariantCulture, out int htmlStyle)
                && htmlStyle > 0)
            {
                chartStyle = htmlStyle;
            }

            PrepareChartExcelUiSuppression();
            try
            {
                PourLog(warnings, "即将 AddChart2 style=" + chartStyle
                    + " xl=" + xlType
                    + " newLayout=" + newLayout
                    + " box=" + left.ToString("0.#", CultureInfo.InvariantCulture) + ","
                    + top.ToString("0.#", CultureInfo.InvariantCulture) + " "
                    + width.ToString("0.#", CultureInfo.InvariantCulture) + "x"
                    + height.ToString("0.#", CultureInfo.InvariantCulture));
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
                    PourLog(warnings, "AddChart2(newLayout) 返回 " + (shape == null ? "null" : "ok"));
                }
                catch (Exception ex1)
                {
                    PourLog(warnings, "AddChart2(newLayout) 失败: " + ex1.Message);
                    try
                    {
                        shape = WppCom.Invoke(shapes, "AddChart2", chartStyle, xlType, left, top, width, height);
                        PourLog(warnings, "AddChart2 返回 " + (shape == null ? "null" : "ok"));
                    }
                    catch (Exception ex2)
                    {
                        PourLog(warnings, "AddChart2 失败: " + ex2.Message);
                        shape = WppCom.Invoke(shapes, "AddChart", xlType, left, top, width, height);
                        PourLog(warnings, "AddChart 返回 " + (shape == null ? "null" : "ok"));
                    }
                }
            }
            catch (Exception ex)
            {
                error = "创建图表失败: " + ex.Message;
                PourLog(warnings, error);
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

            PourLog(warnings, "建图后取 Chart " + (chart == null ? "null" : "ok")
                + " | " + DescribeLiveSeries(chart));
            DismissChartExcelUiForChart(chart);
            try
            {
                if (!TryPourGrid(chart, grid, out error, warnings))
                {
                    TryDelete(shape);
                    shape = null;
                    return false;
                }

                if (applyHtmlChrome)
                {
                    if (!TryApplyFormat(chart, format, warnings, out error))
                    {
                        TryDelete(shape);
                        shape = null;
                        return false;
                    }

                    ChartStyleSnap htmlSnap = SnapFromFormat(format, grid);
                    // 新建无旧图：展示开关按「未提默认关」，避免 AddChart2 默认 ChartTitle=系列名
                    EnsureDisplaySwitches(oldSnap: null, format, htmlSnap, warnings);
                    FinishLineChartLayout(chart, xlType, format, warnings);
                    TryApplyStyleSnap(chart, htmlSnap, warnings, grid, format);
                    TryInheritAxisChrome(chart, htmlSnap, warnings);
                    EnsureNewPieVariesByCategory(chart, grid, warnings);
                    // VaryByCategory 可能重建瓣点，爆炸须在其后按当前点数再写。
                    TryApplyPieExplosionFromFormat(chart, format, warnings);
                }
                else
                {
                    EnsureCategoryAxisLabels(chart, grid);
                }

                return true;
            }
            finally
            {
                DismissChartExcelUiForChart(chart);
            }
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
            // I8：未外链时优先 ChartData 内嵌表；外链图禁止开 Workbook（会弹「链接的文件不可用」），改读 Series。
            string embeddedErr = null;
            if (!IsChartDataLinked(chart)
                && TryReadGridFromEmbeddedSheetAuto(chart, out grid, out embeddedErr)
                && grid != null
                && grid.IsPourable)
            {
                return true;
            }

            if (TryReadGridFromSeries(chart, out grid, out error) && grid != null && grid.IsPourable)
            {
                return true;
            }

            if (!string.IsNullOrEmpty(embeddedErr))
            {
                error = embeddedErr;
            }
            else if (string.IsNullOrEmpty(error))
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
                List<string> cats = ToStringList(WppCom.GetProperty(s1, "XValues"));
                if (cats.Count == 0)
                {
                    // 外链/部分图类别只在轴上，XValues 为空。
                    cats = TryReadCategoryAxisNames(chart);
                }

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

                    if (!IsPieChart(chart) && !string.IsNullOrEmpty(col.Color))
                    {
                        try
                        {
                            object fmt = WppCom.GetProperty(series, "Format");
                            object fill = WppCom.GetProperty(fmt, "Fill");
                            if (string.Equals(col.Color, "none", StringComparison.OrdinalIgnoreCase)
                                || string.Equals(col.Color, "transparent", StringComparison.OrdinalIgnoreCase))
                            {
                                WppCom.TrySetProperty(fill, "Visible", 0);
                            }
                            else if (TryParseHexToOffice(col.Color, out int rgb))
                            {
                                WppCom.TrySetProperty(fill, "Visible", -1);
                                object fc = WppCom.GetProperty(fill, "ForeColor");
                                WppCom.TrySetProperty(fc, "RGB", rgb);
                            }
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

                    if (IsSeriesAxisY2(col.AxisY))
                    {
                        WppCom.TrySetProperty(series, "AxisGroup", XlSecondary);
                    }
                    else if (IsSeriesAxisY(col.AxisY))
                    {
                        WppCom.TrySetProperty(series, "AxisGroup", XlPrimary);
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

        private static void TryReadValueScale(object chart, PptHtmlChartFormat format)
        {
            try
            {
                object y = TryInvoke(chart, "Axes", XlValue, XlPrimary);
                if (y == null)
                {
                    return;
                }

                if (!IsTruthy(WppCom.GetProperty(y, "MinimumScaleIsAuto")))
                {
                    object mn = WppCom.GetProperty(y, "MinimumScale");
                    if (mn != null)
                    {
                        format.AxisYMin = Convert.ToDouble(mn).ToString("0.##", CultureInfo.InvariantCulture);
                    }
                }

                if (!IsTruthy(WppCom.GetProperty(y, "MaximumScaleIsAuto")))
                {
                    object mx = WppCom.GetProperty(y, "MaximumScale");
                    if (mx != null)
                    {
                        format.AxisYMax = Convert.ToDouble(mx).ToString("0.##", CultureInfo.InvariantCulture);
                    }
                }

                if (!IsTruthy(WppCom.GetProperty(y, "MajorUnitIsAuto")))
                {
                    object un = WppCom.GetProperty(y, "MajorUnit");
                    if (un != null)
                    {
                        format.AxisYMajorUnit = Convert.ToDouble(un).ToString("0.##", CultureInfo.InvariantCulture);
                    }
                }
            }
            catch (Exception)
            {
            }
        }

        private static void TryReadExplosion(object chart, PptHtmlChartFormat format)
        {
            try
            {
                // 仅各瓣相同才报整图值；不对称不输出（换数也不继承）。
                int? uniform = TryReadUniformPieExplosion(chart);
                if (uniform.HasValue)
                {
                    format.Explosion = uniform.Value.ToString(CultureInfo.InvariantCulture);
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
                    try
                    {
                        if (IsTruthy(WppCom.GetProperty(chart, "HasTitle")))
                        {
                            object leftover = WppCom.GetProperty(chart, "ChartTitle");
                            if (leftover != null)
                            {
                                WppCom.TrySetProperty(leftover, "Text", "");
                            }

                            WppCom.TrySetProperty(chart, "HasTitle", false);
                        }
                    }
                    catch (Exception)
                    {
                    }

                    return;
                }

                WppCom.TrySetProperty(chart, "HasTitle", true);
                object ct = WppCom.GetProperty(chart, "ChartTitle");
                WppCom.TrySetProperty(ct, "Text", title);
            }
            catch (Exception ex)
            {
                warnings?.Add("未能套用 data-title: " + ex.Message);
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

            if (!string.IsNullOrWhiteSpace(fmt.ShowValue) || !string.IsNullOrWhiteSpace(fmt.ShowPercentage))
            {
                try
                {
                    bool? showValue = ParseOptionalBool(fmt.ShowValue);
                    bool? showPct = ParseOptionalBool(fmt.ShowPercentage);
                    object sc = TryInvoke(chart, "SeriesCollection");
                    int n = Convert.ToInt32(WppCom.GetProperty(sc, "Count"));
                    for (int i = 1; i <= n; i++)
                    {
                        object s = WppCom.GetIndexed(sc, i);
                        if (showValue == true || showPct == true)
                        {
                            object has = WppCom.GetProperty(s, "HasDataLabels");
                            if (!IsTruthy(has))
                            {
                                TryInvoke(s, "ApplyDataLabels", XlDataLabelsShowValue);
                            }
                        }

                        object dls = TryGetDataLabels(s);
                        if (dls == null)
                        {
                            continue;
                        }

                        if (showValue.HasValue)
                        {
                            WppCom.TrySetProperty(dls, "ShowValue", showValue.Value);
                        }

                        if (showPct.HasValue)
                        {
                            WppCom.TrySetProperty(dls, "ShowPercentage", showPct.Value);
                        }
                    }
                }
                catch (Exception)
                {
                    Warn(warnings, "data-show-value/percentage");
                }
            }

            if (string.Equals(fmt.PlotColor, "none", StringComparison.OrdinalIgnoreCase)
                || string.Equals(fmt.PlotColor, "transparent", StringComparison.OrdinalIgnoreCase))
            {
                try
                {
                    TryWriteAreaFill(chart, "PlotArea", false, null);
                }
                catch (Exception)
                {
                    Warn(warnings, "data-plot-color");
                }
            }
            else if (!string.IsNullOrWhiteSpace(fmt.PlotColor) && TryParseHexToOffice(fmt.PlotColor, out int plotRgb))
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
                TryApplyValueAxisScaleFromFormat(chart, fmt, warnings);
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
                TryApplyPieExplosion(chart, exp, warnings);
            }
        }

        /// <summary>饼/环：灌数后点数已定，给当前全部瓣写同一爆炸值。</summary>
        private static void TryApplyPieExplosion(object chart, int explosion, List<string> warnings)
        {
            if (chart == null || !IsPieChart(chart))
            {
                return;
            }

            try
            {
                object sc = TryInvoke(chart, "SeriesCollection");
                if (sc == null)
                {
                    sc = WppCom.GetProperty(chart, "SeriesCollection");
                }

                object s = sc == null ? null : WppCom.GetIndexed(sc, 1);
                if (s == null)
                {
                    return;
                }

                // 系列级顺带写一下（部分环境会带动点）；读回以点为准。
                try
                {
                    WppCom.TrySetProperty(s, "Explosion", explosion);
                }
                catch (Exception)
                {
                }

                object pts = WppCom.GetProperty(s, "Points");
                if (pts == null)
                {
                    pts = TryInvoke(s, "Points");
                }

                if (pts == null)
                {
                    return;
                }

                int n = Convert.ToInt32(WppCom.GetProperty(pts, "Count"));
                for (int i = 1; i <= n; i++)
                {
                    object pt = WppCom.GetIndexed(pts, i);
                    if (pt == null)
                    {
                        pt = TryInvoke(pts, "Item", i);
                    }

                    if (pt != null)
                    {
                        WppCom.TrySetProperty(pt, "Explosion", explosion);
                    }
                }
            }
            catch (Exception)
            {
                Warn(warnings, "data-explosion");
            }
        }

        private static void TryApplyPieExplosionFromFormat(
            object chart,
            PptHtmlChartFormat fmt,
            List<string> warnings)
        {
            if (fmt == null || string.IsNullOrWhiteSpace(fmt.Explosion))
            {
                return;
            }

            if (!int.TryParse(fmt.Explosion, NumberStyles.Integer, CultureInfo.InvariantCulture, out int exp))
            {
                return;
            }

            TryApplyPieExplosion(chart, exp, warnings);
        }

        /// <summary>
        /// 各瓣爆炸值相同则返回该值；不对称或非饼返回 null（换数不继承不对称）。
        /// </summary>
        private static int? TryReadUniformPieExplosion(object chart)
        {
            if (chart == null || !IsPieChart(chart))
            {
                return null;
            }

            try
            {
                object sc = TryInvoke(chart, "SeriesCollection");
                if (sc == null)
                {
                    sc = WppCom.GetProperty(chart, "SeriesCollection");
                }

                object s = sc == null ? null : WppCom.GetIndexed(sc, 1);
                object pts = s == null ? null : WppCom.GetProperty(s, "Points");
                if (pts == null && s != null)
                {
                    pts = TryInvoke(s, "Points");
                }

                if (pts == null)
                {
                    return null;
                }

                int n = Convert.ToInt32(WppCom.GetProperty(pts, "Count"));
                if (n < 1)
                {
                    return null;
                }

                int? uniform = null;
                for (int i = 1; i <= n; i++)
                {
                    object pt = WppCom.GetIndexed(pts, i) ?? TryInvoke(pts, "Item", i);
                    object raw = pt == null ? null : WppCom.GetProperty(pt, "Explosion");
                    int v = raw == null ? 0 : Convert.ToInt32(Convert.ToDouble(raw));
                    if (uniform == null)
                    {
                        uniform = v;
                    }
                    else if (uniform.Value != v)
                    {
                        return null;
                    }
                }

                return uniform;
            }
            catch (Exception)
            {
                return null;
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

    }
}
