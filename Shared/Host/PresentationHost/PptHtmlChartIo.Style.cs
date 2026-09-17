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
                object color = WppCom.GetProperty(chart, "ChartColor");
                if (color != null)
                {
                    snap.ChartColor = Convert.ToInt32(color);
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
                    if (title != null)
                    {
                        snap.Title = Convert.ToString(WppCom.GetProperty(title, "Text") ?? "");
                    }

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
            NormalizeAxisChrome(snap);

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
                    one.PointFills = TryCapturePointFills(series, warnings, (tag ?? "图") + " S" + i);
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

        /// <summary>
        /// 继承只走这一次：Refresh 已在调用方做完。顺序粗→细，后面盖前面。
        /// 结构 → 主题盘 → 标题/图例/区 → 轴结构+网格 → 系列填/扇区 → 线/标记 → 标签。
        /// 轴皮（字色/轴线）不在这里写，整次套回最后一层。
        /// </summary>
        private static void TryApplyStyleSnap(
            object chart,
            ChartStyleSnap snap,
            List<string> warnings,
            PptHtmlChartGrid grid = null,
            PptHtmlChartFormat format = null,
            bool deferTitleOff = false)
        {
            if (chart == null || snap == null)
            {
                StyleLog(warnings, "套回跳过 chart=" + (chart == null ? "null" : "ok")
                    + " snap=" + (snap == null ? "null" : "ok"));
                return;
            }

            StyleLog(warnings, "开始套回 " + DescribeSnap("快照", snap)
                + (deferTitleOff ? " deferTitleOff=True" : ""));
            try
            {
                TryInheritSeriesStructure(chart, snap);
                TryApplyChartTheme(chart, snap);
                TryInheritTitleAndLegend(chart, snap, warnings, deferTitleOff);
                TryWriteAreaFill(chart, "ChartArea", snap.ChartAreaFillVisible, snap.ChartAreaFillRgb);
                TryWriteAreaFill(chart, "PlotArea", snap.PlotFillVisible, snap.PlotFillRgb);
                TryApplyChartGroup(chart, snap);
                TryInheritAxes(chart, snap, grid, format, warnings);
                TryInheritColors(chart, snap, warnings);
                TryInheritSeriesLineAndMarker(chart, snap, warnings);
                TryRestorePieChartType(chart, snap);
                TryApplyPlotLayout(chart, snap);
                TryInheritDataLabels(chart, snap, warnings);
                // 主题/系列/标签可能把 AddChart2 默认标题再打开，正文和开关最后再钉一次
                TryInheritTitleAndLegend(chart, snap, warnings, deferTitleOff);
            }
            catch (Exception ex)
            {
                StyleLog(warnings, "套回原图样式部分失败: " + ex.Message);
            }
        }

        private static void TryInheritSeriesStructure(object chart, ChartStyleSnap snap)
        {
            if (snap.Series == null)
            {
                return;
            }

            for (int i = 0; i < snap.Series.Count; i++)
            {
                object series = GetSeries(chart, i + 1);
                SeriesStyleSnap one = snap.Series[i];
                if (series == null || one == null)
                {
                    continue;
                }

                if (one.ChartType.HasValue
                    && !(IsPieXl(one.ChartType.Value) && IsPieChart(chart)))
                {
                    WppCom.TrySetProperty(series, "ChartType", one.ChartType.Value);
                }

                if (one.AxisGroup.HasValue)
                {
                    WppCom.TrySetProperty(series, "AxisGroup", one.AxisGroup.Value);
                }
            }
        }

        private static void TryInheritTitleAndLegend(
            object chart,
            ChartStyleSnap snap,
            List<string> warnings,
            bool deferTitleOff = false)
        {
            if (snap == null)
            {
                return;
            }

            StyleLog(warnings, "套标题前 snap.HasTitle=" + Convert.ToString(snap.HasTitle)
                + " snap.Title=[" + (snap.Title ?? "(null)") + "] deferTitleOff=" + deferTitleOff
                + " " + DescribeLiveTitle(chart));
            bool explicitOff = snap.HasTitle == false
                || (snap.Title != null && snap.Title.Length == 0);
            if (explicitOff && deferTitleOff)
            {
                StyleLog(warnings, "关标题延后到 Dismiss Excel 后，此处跳过");
            }
            else if (explicitOff)
            {
                TrySetTitle(chart, "", warnings);
                try
                {
                    WppCom.TrySetProperty(chart, "HasTitle", false);
                }
                catch (Exception)
                {
                }
            }
            else if (!string.IsNullOrEmpty(snap.Title))
            {
                TrySetTitle(chart, snap.Title, warnings);
            }
            else if (snap.HasTitle.HasValue)
            {
                WppCom.TrySetProperty(chart, "HasTitle", snap.HasTitle.Value);
            }

            StyleLog(warnings, "套标题后 explicitOff=" + explicitOff + " " + DescribeLiveTitle(chart));
            bool titleOn = !explicitOff
                && (snap.HasTitle == true
                    || !string.IsNullOrEmpty(snap.Title)
                    || (snap.HasTitle == null && IsTruthy(WppCom.GetProperty(chart, "HasTitle"))));
            if (titleOn)
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

            if (snap.HasLegend.HasValue)
            {
                WppCom.TrySetProperty(chart, "HasLegend", snap.HasLegend.Value);
            }

            bool legendOn = snap.HasLegend == true
                || (snap.HasLegend == null && IsTruthy(WppCom.GetProperty(chart, "HasLegend")));
            if (!legendOn)
            {
                return;
            }

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

        private static void TryInheritAxes(
            object chart,
            ChartStyleSnap snap,
            PptHtmlChartGrid grid,
            PptHtmlChartFormat format,
            List<string> warnings)
        {
            HideDeletedAxes(chart, snap);
            TryApplyAxis(chart, XlCategory, XlPrimary, snap.Category);
            TryApplyAxis(chart, XlValue, XlPrimary, snap.Value);
            TryApplyAxis(chart, XlValue, XlSecondary, snap.ValueSecondary);
            if (snap.Category != null && snap.Category.Deleted != true)
            {
                EnsureCategoryAxisLabels(chart, grid);
            }

            EnsureGridlinesMatchSnap(chart, snap, format, warnings);
        }

        private static void TryInheritSeriesLineAndMarker(object chart, ChartStyleSnap snap, List<string> warnings)
        {
            if (snap.Series == null)
            {
                return;
            }

            for (int i = 0; i < snap.Series.Count; i++)
            {
                object series = GetSeries(chart, i + 1);
                if (series == null || snap.Series[i] == null)
                {
                    continue;
                }

                TryApplyLine(series, snap.Series[i].Line, warnings, "S" + (i + 1));
                TryApplyMarker(series, snap.Series[i]);
            }
        }

        private static void TryInheritDataLabels(object chart, ChartStyleSnap snap, List<string> warnings)
        {
            if (snap.Series == null)
            {
                return;
            }

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
                    snap.TickFontColor = TryReadTickLabelsRgb(ticks, font);

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
                        TryCaptureGridlineLine(axis, snap);
                        if (!GridlinesVisuallyOn(snap))
                        {
                            snap.HasMajorGridlines = false;
                        }
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
                    // 主题白/自动色 Type≠RGB，只认显式色会拍成 null，套回后新图轴线就没了。
                    snap.LineRgb = TryReadResolvedRgb(axisColor);
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
                    else if (snap.HasTitle.Value == false)
                    {
                        object at = WppCom.GetProperty(axis, "AxisTitle");
                        TryInvoke(at, "Delete");
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
                // 只写轴结构。字色/轴线整次套回最后一层，避免 NumberFormat 之后的 COM 按风格掀掉。
                if (ticks != null)
                {
                    if (axisType == XlCategory)
                    {
                        WppCom.TrySetProperty(ticks, "NumberFormat", "@");
                    }
                    else if (!string.IsNullOrEmpty(snap.NumberFormat))
                    {
                        WppCom.TrySetProperty(ticks, "NumberFormat", snap.NumberFormat);
                    }
                }

                if (snap.HasMajorGridlines.HasValue || snap.GridlineVisible.HasValue)
                {
                    bool visualOn = GridlinesVisuallyOn(snap);
                    WppCom.TrySetProperty(axis, "HasMajorGridlines", visualOn);
                    if (visualOn)
                    {
                        TryApplyGridlineLine(axis, snap);
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
            FillSnap snap = TryCaptureFormatFill(series, warnings, prefix);
            if (snap != null && (snap.SolidRgb.HasValue || (snap.Stops != null && snap.Stops.Count >= 2)))
            {
                return snap;
            }

            FillSnap interior = TryCaptureInteriorFill(series);
            if (interior != null)
            {
                StyleLog(warnings, prefix + " Interior.Color=" + HexOf(interior.SolidRgb));
                return interior;
            }

            if (snap == null)
            {
                StyleLog(warnings, prefix + " 拍填充失败");
            }

            return snap;
        }

        private static FillSnap TryCaptureFormatFill(object target, List<string> warnings, string prefix)
        {
            object fill;
            try
            {
                object format = WppCom.GetProperty(target, "Format");
                fill = format == null ? null : WppCom.GetProperty(format, "Fill");
            }
            catch (Exception ex)
            {
                StyleLog(warnings, prefix + " Format.Fill 不可用: " + ex.Message);
                return null;
            }

            if (fill == null)
            {
                return null;
            }

            var snap = new FillSnap();
            try
            {
                object vis = WppCom.GetProperty(fill, "Visible");
                if (vis != null)
                {
                    snap.Visible = Convert.ToInt32(vis) != 0;
                }
            }
            catch (Exception)
            {
            }

            try
            {
                object fillType = WppCom.GetProperty(fill, "Type");
                if (fillType != null)
                {
                    snap.FillType = Convert.ToInt32(fillType);
                }
            }
            catch (Exception)
            {
            }

            try
            {
                object angle = WppCom.GetProperty(fill, "GradientAngle");
                if (angle != null)
                {
                    snap.Angle = Convert.ToDouble(angle);
                }
            }
            catch (Exception)
            {
            }

            try
            {
                snap.Stops = TryReadGradientStops(fill, warnings, prefix);
            }
            catch (Exception)
            {
            }

            try
            {
                if (snap.Stops == null || snap.Stops.Count < 2)
                {
                    snap.SolidRgb = TryReadResolvedRgb(WppCom.GetProperty(fill, "ForeColor"));
                }
                else
                {
                    snap.SolidRgb = PickSolidFromStops(snap.Stops);
                }
            }
            catch (Exception)
            {
            }

            if (snap.SolidRgb.HasValue || (snap.Stops != null && snap.Stops.Count >= 2) || snap.Visible.HasValue)
            {
                StyleLog(warnings, prefix + " Fill.Visible=" + snap.Visible
                    + " Type=" + snap.FillType
                    + " stops=" + DescribeStops(snap.Stops)
                    + " solid=" + HexOf(snap.SolidRgb));
                return snap;
            }

            return null;
        }

        private static FillSnap TryCaptureInteriorFill(object target)
        {
            if (target == null)
            {
                return null;
            }

            try
            {
                object interior = WppCom.GetProperty(target, "Interior");
                if (interior == null)
                {
                    return null;
                }

                object color = WppCom.GetProperty(interior, "Color");
                if (color == null)
                {
                    return null;
                }

                int rgb = Convert.ToInt32(Convert.ToDouble(color)) & 0x00FFFFFF;
                return new FillSnap { Visible = true, SolidRgb = rgb };
            }
            catch (Exception)
            {
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

            if (TryApplyFormatFill(series, snap, warnings, prefix))
            {
                return;
            }

            if (snap.SolidRgb.HasValue && TryApplyInteriorFill(series, snap.SolidRgb.Value))
            {
                StyleLog(warnings, prefix + " Interior 实色 " + HexOf(snap.SolidRgb));
                return;
            }

            StyleLog(warnings, prefix + " 3D/COM 套不上");
        }

        private static bool TryApplyFormatFill(object target, FillSnap snap, List<string> warnings, string prefix)
        {
            try
            {
                object format = WppCom.GetProperty(target, "Format");
                object fill = format == null ? null : WppCom.GetProperty(format, "Fill");
                if (fill == null)
                {
                    return false;
                }

                if (snap.Visible == false && (snap.Stops == null || snap.Stops.Count < 2))
                {
                    WppCom.TrySetProperty(fill, "Visible", 0);
                    StyleLog(warnings, prefix + " 隐藏填充");
                    return true;
                }

                WppCom.TrySetProperty(fill, "Visible", -1);
                if (snap.Stops != null && snap.Stops.Count >= 2)
                {
                    bool ok = TryWriteGradientStops(fill, snap.Stops, snap.Angle);
                    StyleLog(warnings, prefix + " 渐变 " + DescribeStops(snap.Stops) + " ok=" + ok);
                    if (ok)
                    {
                        return true;
                    }
                }

                if (snap.SolidRgb.HasValue)
                {
                    TryInvoke(fill, "Solid");
                    object fc = WppCom.GetProperty(fill, "ForeColor");
                    WppCom.TrySetProperty(fc, "Type", 1);
                    WppCom.TrySetProperty(fc, "RGB", snap.SolidRgb.Value);
                    StyleLog(warnings, prefix + " 实色 " + HexOf(snap.SolidRgb));
                    return true;
                }

                return snap.Visible.HasValue;
            }
            catch (Exception ex)
            {
                StyleLog(warnings, prefix + " Format.Fill 异常: " + ex.Message);
                return false;
            }
        }

        private static bool TryApplyInteriorFill(object target, int rgb)
        {
            try
            {
                object interior = WppCom.GetProperty(target, "Interior");
                if (interior == null)
                {
                    return false;
                }

                WppCom.TrySetProperty(interior, "Color", rgb);
                object wrote = WppCom.GetProperty(interior, "Color");
                if (wrote == null)
                {
                    return false;
                }

                int got = Convert.ToInt32(Convert.ToDouble(wrote)) & 0x00FFFFFF;
                return got == (rgb & 0x00FFFFFF);
            }
            catch (Exception)
            {
                return false;
            }
        }

        private static List<FillSnap> TryCapturePointFills(
            object series,
            List<string> warnings,
            string tag)
        {
            object pts = TryGetPoints(series);
            if (pts == null)
            {
                return null;
            }

            try
            {
                int n = Convert.ToInt32(WppCom.GetProperty(pts, "Count"));
                if (n <= 0)
                {
                    return null;
                }

                if (n > 30)
                {
                    n = 30;
                }

                var list = new List<FillSnap>();
                int got = 0;
                for (int i = 1; i <= n; i++)
                {
                    object pt = WppCom.GetIndexed(pts, i);
                    FillSnap fill = pt == null
                        ? null
                        : TryCaptureFill(pt, warnings, (tag ?? "S") + " P" + i);
                    list.Add(fill);
                    if (fill != null)
                    {
                        got++;
                    }
                }

                if (got == 0)
                {
                    StyleLog(warnings, (tag ?? "S") + " 点填充 0/" + n);
                    return null;
                }

                StyleLog(warnings, (tag ?? "S") + " 点填充 " + got + "/" + n);
                return list;
            }
            catch (Exception ex)
            {
                StyleLog(warnings, (tag ?? "S") + " 拍点填充失败: " + ex.Message);
                return null;
            }
        }

        private static void TryApplyPointFills(
            object series,
            List<FillSnap> fills,
            List<string> warnings,
            string tag)
        {
            if (fills == null || fills.Count == 0)
            {
                return;
            }

            object pts = TryGetPoints(series);
            if (pts == null)
            {
                return;
            }

            try
            {
                int live = Convert.ToInt32(WppCom.GetProperty(pts, "Count"));
                int n = Math.Min(live, fills.Count);
                for (int i = 1; i <= n; i++)
                {
                    FillSnap fill = fills[i - 1];
                    if (fill == null)
                    {
                        continue;
                    }

                    object pt = WppCom.GetIndexed(pts, i);
                    if (pt == null)
                    {
                        continue;
                    }

                    TryApplyFill(pt, fill, warnings, (tag ?? "S") + " P" + i);
                }
            }
            catch (Exception ex)
            {
                StyleLog(warnings, "套点填充 " + (tag ?? "") + " 异常: " + ex.Message);
            }
        }

        private static object TryGetPoints(object series)
        {
            if (series == null)
            {
                return null;
            }

            object pts = TryInvoke(series, "Points");
            if (pts != null)
            {
                return pts;
            }

            try
            {
                return WppCom.GetProperty(series, "Points");
            }
            catch (Exception)
            {
                return null;
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
                bool? showValue = TryGetBoolProperty(dls, "ShowValue");
                if (showValue.HasValue)
                {
                    one.ShowValue = showValue;
                }

                bool? showPct = TryGetBoolProperty(dls, "ShowPercentage");
                if (showPct.HasValue)
                {
                    one.ShowPercentage = showPct;
                }

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

        private static bool? TryGetBoolProperty(object target, string name)
        {
            if (target == null || string.IsNullOrEmpty(name))
            {
                return null;
            }

            try
            {
                object v = WppCom.GetProperty(target, name);
                if (v == null)
                {
                    return null;
                }

                return IsTruthy(v);
            }
            catch (Exception)
            {
                return null;
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
                || one.ShowValue.HasValue
                || one.ShowPercentage.HasValue
                || one.DataLabelPosition.HasValue
                || one.DataLabelFontColor.HasValue
                || !string.IsNullOrEmpty(one.DataLabelNumberFormat);
            StyleLog(warnings, prefix + " HasDataLabels=" + one.HasDataLabels
                + " want=" + want
                + " showValue=" + one.ShowValue
                + " showPct=" + one.ShowPercentage
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

            // 有快照则听快照；未提的内容开关不硬改（跟 ApplyDataLabels 默认 / 旧图继承）
            if (one.ShowValue.HasValue)
            {
                WppCom.TrySetProperty(dls, "ShowValue", one.ShowValue.Value);
            }

            if (one.ShowPercentage.HasValue)
            {
                WppCom.TrySetProperty(dls, "ShowPercentage", one.ShowPercentage.Value);
            }

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
                object pts = TryGetPoints(series);
                if (pts == null)
                {
                    return null;
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

        /// <summary>
        /// 改已有图时 HTML 只负责灌数。类型/轴/标题以旧图快照为准，避免 HTML 多写或少写。
        /// </summary>
        private static bool IsInsideGroup(object shape)
        {
            if (shape == null)
            {
                return false;
            }

            try
            {
                return WppCom.GetProperty(shape, "ParentGroup") != null;
            }
            catch (Exception)
            {
                return false;
            }
        }

        private static void MergeGridlinesFromFormat(ChartStyleSnap snap, PptHtmlChartFormat format)
        {
            if (snap == null || format == null || string.IsNullOrWhiteSpace(format.Gridlines))
            {
                return;
            }

            if (snap.Value == null)
            {
                snap.Value = new AxisStyleSnap();
            }

            snap.Value.HasMajorGridlines = IsTrue(format.Gridlines);
        }

        private static void EnsureGridlinesMatchSnap(
            object chart,
            ChartStyleSnap snap,
            PptHtmlChartFormat format,
            List<string> warnings)
        {
            bool formatOff = format != null
                && !string.IsNullOrWhiteSpace(format.Gridlines)
                && !IsTrue(format.Gridlines);
            bool snapOff = snap != null && snap.Value != null
                && (snap.Value.HasMajorGridlines == false
                    || snap.Value.GridlineVisible == false);
            if (!formatOff && !snapOff)
            {
                return;
            }

            ForceAxisGridlinesOff(chart, XlValue, XlPrimary);
            ForceAxisGridlinesOff(chart, XlValue, XlSecondary);
            StyleLog(warnings, "按原图/稿关闭网格线");
        }

        private static void ForceAxisGridlinesOff(object chart, int axisType, int group)
        {
            if (chart == null)
            {
                return;
            }

            try
            {
                object axis = TryInvoke(chart, "Axes", axisType, group);
                if (axis == null)
                {
                    return;
                }

                WppCom.TrySetProperty(axis, "HasMajorGridlines", false);
                object gl = WppCom.GetProperty(axis, "MajorGridlines");
                object glLine = gl == null
                    ? null
                    : WppCom.GetProperty(WppCom.GetProperty(gl, "Format"), "Line");
                if (glLine != null)
                {
                    WppCom.TrySetProperty(glLine, "Visible", 0);
                    WppCom.TrySetProperty(glLine, "Transparency", 1.0);
                }
            }
            catch (Exception)
            {
            }
        }

        private static void SyncFormatToOldSnap(
            PptHtmlChartFormat format,
            int xlType,
            List<string> warnings)
        {
            if (format == null)
            {
                return;
            }

            string canon = CanonicalTypeFromXl(xlType);
            if (!string.Equals(format.ChartType, canon, StringComparison.OrdinalIgnoreCase))
            {
                StyleLog(warnings, "format.ChartType 按旧图 " + (format.ChartType ?? "") + " → " + canon);
                format.ChartType = canon;
            }

            format.AxisX = null;
            format.AxisY = null;
            format.AxisYSecondary = null;
            StyleLog(warnings, "SyncFormatToOldSnap 清空 format.Title=[" + (format.Title ?? "(null)") + "] Mentioned=" + format.TitleMentioned);
            format.Title = null;
        }

        private static void TryWriteTickFont(object ticks, AxisStyleSnap snap)
        {
            if (ticks == null || snap == null)
            {
                return;
            }

            object font = WppCom.GetProperty(ticks, "Font");
            if (font == null)
            {
                return;
            }

            if (!string.IsNullOrEmpty(snap.TickFontName))
            {
                WppCom.TrySetProperty(font, "Name", snap.TickFontName);
            }

            if (snap.TickFontColor.HasValue)
            {
                TryWriteFontRgb(font, snap.TickFontColor.Value);
                TryWriteTickLabelsRgb(ticks, snap.TickFontColor.Value);
            }

            if (snap.TickFontSize.HasValue)
            {
                WppCom.TrySetProperty(font, "Size", snap.TickFontSize.Value);
            }
        }

        /// <summary>
        /// 轴皮最后一层：只写字色/轴线，不再走 NumberFormat / 分类名。
        /// </summary>
        private static void TryInheritAxisChrome(object chart, ChartStyleSnap snap, List<string> warnings)
        {
            if (chart == null || snap == null)
            {
                return;
            }

            StyleLog(warnings, "轴皮最后一层 "
                + "catTick=" + (snap.Category == null ? "-" : HexOf(snap.Category.TickFontColor))
                + " valTick=" + (snap.Value == null ? "-" : HexOf(snap.Value.TickFontColor)));
            TryInheritOneAxisChrome(chart, XlCategory, XlPrimary, snap.Category);
            TryInheritOneAxisChrome(chart, XlValue, XlPrimary, snap.Value);
            TryInheritOneAxisChrome(chart, XlValue, XlSecondary, snap.ValueSecondary);
        }

        private static void TryInheritOneAxisChrome(object chart, int axisType, int group, AxisStyleSnap snap)
        {
            if (chart == null || snap == null || snap.Deleted == true)
            {
                return;
            }

            try
            {
                object axis = TryInvoke(chart, "Axes", axisType, group);
                if (axis == null)
                {
                    return;
                }

                object ticks = WppCom.GetProperty(axis, "TickLabels");
                TryWriteTickFont(ticks, snap);
                TryApplyAxisLine(axis, snap);
            }
            catch (Exception)
            {
            }
        }

        private static void TryApplyAxisLine(object axis, AxisStyleSnap snap)
        {
            if (axis == null || snap == null)
            {
                return;
            }

            object axisLine = WppCom.GetProperty(WppCom.GetProperty(axis, "Format"), "Line");
            if (axisLine == null)
            {
                return;
            }

            if (snap.LineVisible == false)
            {
                WppCom.TrySetProperty(axisLine, "Visible", 0);
                return;
            }

            // AddChart2 + ChartStyle 默认常无线；Visible 为 true/未拍到时都要把线打开。
            WppCom.TrySetProperty(axisLine, "Visible", -1);
            int rgb = snap.LineRgb ?? 0x00FFFFFF;
            TryWriteLineRgb(axisLine, rgb);
            if (snap.LineWeight.HasValue && !IsPhantomWeight(snap.LineWeight.Value))
            {
                WppCom.TrySetProperty(axisLine, "Weight", snap.LineWeight.Value);
            }
        }

        /// <summary>
        /// 横轴年能拍到白色、纵轴主题色拍空时，两边轴线/刻度字互相补，避免一边白一边棕橙。
        /// </summary>
        private static void NormalizeAxisChrome(ChartStyleSnap snap)
        {
            if (snap == null)
            {
                return;
            }

            InheritAxisChrome(snap.Category, snap.Value);
            InheritAxisChrome(snap.Value, snap.Category);
            EnsureDefaultTickColor(snap.Category);
            EnsureDefaultTickColor(snap.Value);
        }

        private static void EnsureDefaultTickColor(AxisStyleSnap ax)
        {
            if (ax == null || ax.Deleted == true || ax.TickFontColor.HasValue)
            {
                return;
            }

            // 图表刻度常是主题色，COM 解析不成 RGB。深色页上回落白，避免落到主题灰/棕橙。
            ax.TickFontColor = 0x00FFFFFF;
        }

        private static void InheritAxisChrome(AxisStyleSnap dest, AxisStyleSnap src)
        {
            if (dest == null || src == null || dest.Deleted == true)
            {
                return;
            }

            if (!dest.TickFontColor.HasValue && src.TickFontColor.HasValue)
            {
                dest.TickFontColor = src.TickFontColor;
            }

            if (string.IsNullOrEmpty(dest.TickFontName) && !string.IsNullOrEmpty(src.TickFontName))
            {
                dest.TickFontName = src.TickFontName;
            }

            if (!dest.TickFontSize.HasValue && src.TickFontSize.HasValue)
            {
                dest.TickFontSize = src.TickFontSize;
            }

            if (dest.LineVisible != false && !dest.LineRgb.HasValue && src.LineRgb.HasValue)
            {
                dest.LineRgb = src.LineRgb;
            }
        }

        private static void HideDeletedAxes(object chart, ChartStyleSnap snap)
        {
            if (chart == null || snap == null)
            {
                return;
            }

            if (snap.Category == null || snap.Category.Deleted == true)
            {
                TryHideAxis(chart, XlCategory, XlPrimary);
            }

            if (snap.Value == null || snap.Value.Deleted == true)
            {
                TryHideAxis(chart, XlValue, XlPrimary);
            }

            if (snap.ValueSecondary == null || snap.ValueSecondary.Deleted == true)
            {
                TryHideAxis(chart, XlValue, XlSecondary);
            }
        }

        /// <summary>
        /// 折线灌数/改类型后 COM 偶发不重算坐标，点贴在基线像没应用完。强制 Refresh + 自动刻度。
        /// </summary>
        private static void FinishLineChartLayout(
            object chart,
            int xlType,
            PptHtmlChartFormat format,
            List<string> warnings)
        {
            if (chart == null || !IsLineLike(xlType))
            {
                TryInvoke(chart, "Refresh");
                return;
            }

            try
            {
                object current = WppCom.GetProperty(chart, "ChartType");
                if (current != null
                    && !string.Equals(
                        CanonicalTypeFromXl(Convert.ToInt32(current)),
                        CanonicalTypeFromXl(xlType),
                        StringComparison.OrdinalIgnoreCase))
                {
                    StyleLog(warnings, "收尾改回折线 ChartType=" + current + " → " + xlType);
                    WppCom.TrySetProperty(chart, "ChartType", xlType);
                }
            }
            catch (Exception)
            {
            }

            TryInvoke(chart, "Refresh");

            bool pinScale = format != null
                && (!string.IsNullOrWhiteSpace(format.AxisYMin)
                    || !string.IsNullOrWhiteSpace(format.AxisYMax));
            if (pinScale)
            {
                return;
            }

            foreach (int group in new[] { XlPrimary, XlSecondary })
            {
                try
                {
                    object y = TryInvoke(chart, "Axes", XlValue, group);
                    if (y == null)
                    {
                        continue;
                    }

                    WppCom.TrySetProperty(y, "MinimumScaleIsAuto", true);
                    WppCom.TrySetProperty(y, "MaximumScaleIsAuto", true);
                }
                catch (Exception)
                {
                }
            }

            TryInvoke(chart, "Refresh");
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

        private static bool GridlinesVisuallyOn(AxisStyleSnap snap)
        {
            if (snap == null)
            {
                return false;
            }

            if (snap.HasMajorGridlines == false)
            {
                return false;
            }

            if (snap.GridlineVisible == false)
            {
                return false;
            }

            if (snap.GridlineTransparency.HasValue && IsFullyTransparent(snap.GridlineTransparency.Value))
            {
                return false;
            }

            return snap.HasMajorGridlines == true;
        }

        private static bool IsFullyTransparent(double transparency)
        {
            if (transparency > 1.0)
            {
                return transparency >= 95.0;
            }

            return transparency >= 0.95;
        }

        private static void TryCaptureGridlineLine(object axis, AxisStyleSnap snap)
        {
            if (axis == null || snap == null)
            {
                return;
            }

            try
            {
                object gl = WppCom.GetProperty(axis, "MajorGridlines");
                object glLine = gl == null ? null : WppCom.GetProperty(WppCom.GetProperty(gl, "Format"), "Line");
                if (glLine == null)
                {
                    return;
                }

                object vis = WppCom.GetProperty(glLine, "Visible");
                if (vis != null)
                {
                    snap.GridlineVisible = Convert.ToInt32(vis) != 0;
                }

                object weight = WppCom.GetProperty(glLine, "Weight");
                if (weight != null)
                {
                    snap.GridlineWeight = Convert.ToDouble(weight);
                }

                object trans = WppCom.GetProperty(glLine, "Transparency");
                if (trans != null)
                {
                    snap.GridlineTransparency = Convert.ToDouble(trans);
                }

                object glColor = WppCom.GetProperty(glLine, "ForeColor");
                snap.MajorGridlineRgb = TryReadExplicitRgb(glColor);
            }
            catch (Exception)
            {
            }
        }

        private static void TryApplyGridlineLine(object axis, AxisStyleSnap snap)
        {
            if (axis == null || snap == null)
            {
                return;
            }

            try
            {
                object gl = WppCom.GetProperty(axis, "MajorGridlines");
                object glLine = gl == null ? null : WppCom.GetProperty(WppCom.GetProperty(gl, "Format"), "Line");
                if (glLine == null)
                {
                    return;
                }

                if (snap.GridlineVisible.HasValue)
                {
                    WppCom.TrySetProperty(glLine, "Visible", snap.GridlineVisible.Value ? -1 : 0);
                }

                if (snap.MajorGridlineRgb.HasValue)
                {
                    object glColor = WppCom.GetProperty(glLine, "ForeColor");
                    WppCom.TrySetProperty(glColor, "RGB", snap.MajorGridlineRgb.Value);
                }

                if (snap.GridlineWeight.HasValue
                    && snap.GridlineWeight.Value > 0
                    && snap.GridlineWeight.Value <= 40)
                {
                    WppCom.TrySetProperty(glLine, "Weight", snap.GridlineWeight.Value);
                }

                if (snap.GridlineTransparency.HasValue)
                {
                    WppCom.TrySetProperty(glLine, "Transparency", snap.GridlineTransparency.Value);
                }
            }
            catch (Exception)
            {
            }
        }

        private static void TryHideAxis(object chart, int axisType, int group)
        {
            try
            {
                object axis = TryInvoke(chart, "Axes", axisType, group);
                if (axis != null)
                {
                    WppCom.TrySetProperty(axis, "HasTitle", false);
                    object title = WppCom.GetProperty(axis, "AxisTitle");
                    TryInvoke(title, "Delete");
                }
            }
            catch (Exception)
            {
            }

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

        private static void PourLog(List<string> warnings, string message)
        {
            if (string.IsNullOrEmpty(message))
            {
                return;
            }

            string line = "[PptChartPour] " + message;
            try
            {
                EasyWriteLog.WriteLine(line);
                EasyWriteLog.Flush();
            }
            catch (Exception)
            {
            }

            warnings?.Add(line);
        }

        private static string DescribeWantGrid(PptHtmlChartGrid grid)
        {
            return DescribeGridCompact(grid, "稿");
        }

        private static string DescribeGridCompact(PptHtmlChartGrid grid, string label)
        {
            if (grid == null || grid.Rows == null || grid.Columns == null)
            {
                return label + "=null";
            }

            var sb = new StringBuilder();
            sb.Append(label).Append(' ').Append(grid.Rows.Count).Append('x').Append(grid.Columns.Count);
            sb.Append(" cols=[");
            for (int c = 0; c < grid.Columns.Count; c++)
            {
                PptHtmlChartColumn col = grid.Columns[c];
                if (c > 0)
                {
                    sb.Append('|');
                }

                sb.Append(col?.Role ?? "?");
                if (!string.IsNullOrEmpty(col?.Name))
                {
                    sb.Append(':').Append(col.Name);
                }
            }

            sb.Append(']');

            var cats = new List<string>();
            for (int r = 0; r < grid.Rows.Count; r++)
            {
                List<string> row = grid.Rows[r];
                cats.Add(row != null && row.Count > 0 ? row[0] ?? "" : "");
            }

            sb.Append(" cat=[").Append(string.Join(",", cats)).Append(']');

            for (int c = 0; c < grid.Columns.Count; c++)
            {
                if (grid.Columns[c]?.Role == "category")
                {
                    continue;
                }

                var vals = new List<string>();
                for (int r = 0; r < grid.Rows.Count; r++)
                {
                    List<string> row = grid.Rows[r];
                    vals.Add(row != null && c < row.Count ? row[c] ?? "" : "");
                }

                string colName = grid.Columns[c]?.Name ?? ("col" + c);
                sb.Append(' ').Append(colName).Append("=[").Append(string.Join(",", vals)).Append(']');
            }

            return sb.ToString();
        }

        private static string DescribeSeriesColumnMap(PptHtmlChartGrid live)
        {
            if (live?.Columns == null)
            {
                return "映射=?";
            }

            var parts = new List<string>();
            int si = 0;
            for (int c = 0; c < live.Columns.Count; c++)
            {
                if (live.Columns[c]?.Role == "category")
                {
                    continue;
                }

                si++;
                parts.Add("S" + si + "→列" + (c + 1) + "「" + (live.Columns[c].Name ?? "") + "」");
            }

            return parts.Count == 0 ? "映射=无value列" : string.Join(" ", parts);
        }

        private static string FindFirstGridMismatch(
            PptHtmlChartGrid want,
            PptHtmlChartGrid got,
            bool compareNames)
        {
            if (want == null || got == null)
            {
                return "want或got为null";
            }

            if (want.Columns == null || got.Columns == null)
            {
                return "列定义缺失 want="
                    + (want.Columns == null ? "null" : want.Columns.Count.ToString(CultureInfo.InvariantCulture))
                    + " got="
                    + (got.Columns == null ? "null" : got.Columns.Count.ToString(CultureInfo.InvariantCulture));
            }

            if (want.Rows == null || got.Rows == null)
            {
                return "行缺失 want="
                    + (want.Rows == null ? "null" : want.Rows.Count.ToString(CultureInfo.InvariantCulture))
                    + " got="
                    + (got.Rows == null ? "null" : got.Rows.Count.ToString(CultureInfo.InvariantCulture));
            }

            if (want.Columns.Count != got.Columns.Count)
            {
                return "列数不一致 want=" + want.Columns.Count + " got=" + got.Columns.Count
                    + " | want=" + DescribeGridColumnRoles(want)
                    + " | got=" + DescribeGridColumnRoles(got);
            }

            if (want.Rows.Count != got.Rows.Count)
            {
                return "行数不一致 want=" + want.Rows.Count + " got=" + got.Rows.Count;
            }

            for (int c = 0; c < want.Columns.Count; c++)
            {
                string role = want.Columns[c].Role ?? "";
                if (!string.Equals(got.Columns[c].Role ?? "", role, StringComparison.OrdinalIgnoreCase))
                {
                    return "列" + (c + 1) + " role 不一致 want=" + role + " got=" + (got.Columns[c].Role ?? "");
                }

                if (compareNames
                    && role != "category"
                    && !string.Equals(got.Columns[c].Name ?? "", want.Columns[c].Name ?? "", StringComparison.Ordinal))
                {
                    return "列" + (c + 1) + " 名不一致 want「" + (want.Columns[c].Name ?? "")
                        + "」got「" + (got.Columns[c].Name ?? "") + "」";
                }
            }

            for (int r = 0; r < want.Rows.Count; r++)
            {
                List<string> a = want.Rows[r];
                List<string> b = got.Rows[r];
                if (a == null || b == null)
                {
                    return "行" + (r + 1) + " 缺失 want=" + (a == null ? "null" : "ok")
                        + " got=" + (b == null ? "null" : "ok");
                }

                if (a.Count < want.Columns.Count || b.Count < got.Columns.Count)
                {
                    return "行" + (r + 1) + " 列数不足 wantCells=" + a.Count
                        + " gotCells=" + b.Count + " wantCols=" + want.Columns.Count;
                }

                for (int c = 0; c < want.Columns.Count; c++)
                {
                    string colLabel = DescribeGridCellLabel(want, c);
                    if (want.Columns[c].Role == "value")
                    {
                        string wantRaw = a[c] ?? "";
                        string gotRaw = b[c] ?? "";
                        bool wantOk = TryParseLooseDouble(wantRaw, out double x);
                        bool gotOk = TryParseLooseDouble(gotRaw, out double y);
                        if (!wantOk || !gotOk)
                        {
                            return "行" + (r + 1) + " " + colLabel + " 数值解析失败 want「" + wantRaw
                                + "」parse=" + wantOk + " got「" + gotRaw + "」parse=" + gotOk;
                        }

                        if (Math.Abs(x - y) > 0.0001)
                        {
                            return "行" + (r + 1) + " " + colLabel + " 数值不一致 want=" + x.ToString("0.####", CultureInfo.InvariantCulture)
                                + " got=" + y.ToString("0.####", CultureInfo.InvariantCulture)
                                + " Δ=" + Math.Abs(x - y).ToString("0.####", CultureInfo.InvariantCulture)
                                + " (raw want「" + wantRaw + "」got「" + gotRaw + "」)";
                        }
                    }
                    else if (!string.Equals(a[c] ?? "", b[c] ?? "", StringComparison.Ordinal))
                    {
                        return "行" + (r + 1) + " " + colLabel + " 类别不一致 want「" + (a[c] ?? "")
                            + "」got「" + (b[c] ?? "") + "」";
                    }
                }
            }

            return null;
        }

        private static string DescribeGridColumnRoles(PptHtmlChartGrid grid)
        {
            if (grid?.Columns == null)
            {
                return "[]";
            }

            var parts = new List<string>();
            foreach (PptHtmlChartColumn col in grid.Columns)
            {
                string role = col?.Role ?? "?";
                if (string.IsNullOrEmpty(col?.Name))
                {
                    parts.Add(role);
                }
                else
                {
                    parts.Add(role + ":" + col.Name);
                }
            }

            return "[" + string.Join("|", parts) + "]";
        }

        private static string DescribeGridCellLabel(PptHtmlChartGrid grid, int colIndex)
        {
            if (grid?.Columns == null || colIndex < 0 || colIndex >= grid.Columns.Count)
            {
                return "列" + (colIndex + 1);
            }

            PptHtmlChartColumn col = grid.Columns[colIndex];
            string role = col?.Role ?? "?";
            if (string.IsNullOrEmpty(col?.Name))
            {
                return "列" + (colIndex + 1) + "(" + role + ")";
            }

            return "列" + (colIndex + 1) + "「" + col.Name + "」(" + role + ")";
        }

        private static string DescribeLiveSeries(object chart)
        {
            if (chart == null)
            {
                return "图=null";
            }

            try
            {
                object t = WppCom.GetProperty(chart, "ChartType");
                int n = GetSeriesCount(chart);
                var sb = new StringBuilder();
                sb.Append("xl=").Append(t == null ? "?" : t.ToString());
                sb.Append(" series=").Append(n);
                int use = Math.Min(n, 3);
                for (int i = 1; i <= use; i++)
                {
                    object series = GetSeries(chart, i);
                    if (series == null)
                    {
                        sb.Append(" | S").Append(i).Append("=null");
                        continue;
                    }

                    List<string> xs = ToStringList(WppCom.GetProperty(series, "XValues"));
                    List<string> ys = ToStringList(WppCom.GetProperty(series, "Values"));
                    sb.Append(" | S").Append(i)
                        .Append(" pts=").Append(TryGetPointCount(series))
                        .Append(" X=[").Append(string.Join(",", xs)).Append("]")
                        .Append(" Y=[").Append(string.Join(",", ys)).Append("]");
                }

                return sb.ToString();
            }
            catch (Exception ex)
            {
                return "图读失败: " + ex.Message;
            }
        }

        private static string TryPropString(object target, string name)
        {
            try
            {
                object v = WppCom.GetProperty(target, name);
                return v == null ? null : Convert.ToString(v);
            }
            catch (Exception)
            {
                return null;
            }
        }

        private static string TryRangeAddress(object range)
        {
            if (range == null)
            {
                return "null";
            }

            return TryPropString(range, "Address")
                ?? TryPropString(range, "AddressLocal")
                ?? "ok";
        }

        private static object TryGetProp(object target, string name)
        {
            try
            {
                return WppCom.GetProperty(target, name);
            }
            catch (Exception)
            {
                return null;
            }
        }

        private static int ReadListObjectCount(object ws)
        {
            try
            {
                object lists = WppCom.GetProperty(ws, "ListObjects");
                return lists == null ? -1 : Convert.ToInt32(WppCom.GetProperty(lists, "Count"));
            }
            catch (Exception)
            {
                return -1;
            }
        }

        private static string DescribeSheetCells(object ws, int lastRow, int lastCol)
        {
            if (ws == null)
            {
                return "sheet=null";
            }

            var sb = new StringBuilder();
            sb.Append("name=").Append(TryPropString(ws, "Name") ?? "?");
            sb.Append(" lists=").Append(ReadListObjectCount(ws));
            sb.Append(" used=").Append(TryRangeAddress(TryGetProp(ws, "UsedRange")));
            sb.Append(" cells=");
            for (int r = 1; r <= lastRow; r++)
            {
                if (r > 1)
                {
                    sb.Append(" / ");
                }

                for (int c = 1; c <= lastCol; c++)
                {
                    if (c > 1)
                    {
                        sb.Append("|");
                    }

                    sb.Append(FormatCell(GetCell(ws, r, c)));
                }
            }

            return sb.ToString();
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
                .Append(" chartColor=").Append(snap.ChartColor)
                .Append(" title=").Append(snap.HasTitle)
                .Append("/").Append(snap.Title ?? "")
                .Append(" legend=").Append(snap.HasLegend)
                .Append(" plotFill=").Append(snap.PlotFillVisible)
                .Append(" gap=").Append(snap.GapWidth)
                .Append(" overlap=").Append(snap.Overlap)
                .Append(" catDel=").Append(snap.Category == null ? "null" : Convert.ToString(snap.Category.Deleted))
                .Append(" valDel=").Append(snap.Value == null ? "null" : Convert.ToString(snap.Value.Deleted))
                .Append(" ").Append(DescribeAxisChrome("cat", snap.Category))
                .Append(" ").Append(DescribeAxisChrome("val", snap.Value));
            if (snap.Series != null)
            {
                for (int i = 0; i < snap.Series.Count; i++)
                {
                    sb.Append(" | ").Append(DescribeSeries(i + 1, snap.Series[i]));
                }
            }

            return sb.ToString();
        }

        private static string DescribeAxisChrome(string tag, AxisStyleSnap ax)
        {
            if (ax == null)
            {
                return tag + "=null";
            }

            return tag
                + "Line=" + (ax.LineVisible == false ? "off" : "on")
                + "/" + HexOf(ax.LineRgb)
                + " tick=" + HexOf(ax.TickFontColor);
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
                + " fillStops=" + DescribeStops(one.Fill == null ? null : one.Fill.Stops)
                + " points=" + (one.PointFills == null ? 0 : one.PointFills.Count);
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
            int? rgb = TryReadTickFontRgb(font);
            return rgb.HasValue ? OfficeRgbToHex(rgb.Value) : null;
        }

        private static int? TryReadTickLabelsRgb(object ticks, object font)
        {
            int? rgb = TryReadTickFontRgb(font);
            return rgb.HasValue ? rgb : TryReadResolvedRgb(TryGetTickForeColor(ticks));
        }

        private static object TryGetTickForeColor(object ticks)
        {
            try
            {
                object format = ticks == null ? null : WppCom.GetProperty(ticks, "Format");
                object tf2 = format == null ? null : WppCom.GetProperty(format, "TextFrame2");
                object tr = tf2 == null ? null : WppCom.GetProperty(tf2, "TextRange");
                object f2 = tr == null ? null : WppCom.GetProperty(tr, "Font");
                object fill = f2 == null ? null : WppCom.GetProperty(f2, "Fill");
                return fill == null ? null : WppCom.GetProperty(fill, "ForeColor");
            }
            catch (Exception)
            {
                return null;
            }
        }

        private static int? TryReadTickFontRgb(object font)
        {
            if (font == null)
            {
                return null;
            }

            try
            {
                int? viaFormat = TryReadResolvedRgb(WppCom.GetProperty(font, "ColorFormat"));
                if (viaFormat.HasValue)
                {
                    return viaFormat;
                }
            }
            catch (Exception)
            {
            }

            try
            {
                object color = WppCom.GetProperty(font, "Color");
                int? rgb = TryReadResolvedRgb(color);
                if (rgb.HasValue)
                {
                    return rgb;
                }

                if (color == null)
                {
                    return null;
                }

                int value = Convert.ToInt32(Convert.ToDouble(color)) & 0x00FFFFFF;
                // 主题色索引通常很小；把它当 RGB 会写成棕橙。
                if (value <= 80)
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

        private static void TryWriteFontRgb(object font, int rgb)
        {
            if (font == null)
            {
                return;
            }

            WppCom.TrySetProperty(font, "Color", rgb);
            try
            {
                object color = WppCom.GetProperty(font, "Color");
                WppCom.TrySetProperty(color, "Type", 1);
                WppCom.TrySetProperty(color, "RGB", rgb);
            }
            catch (Exception)
            {
            }

            try
            {
                object cf = WppCom.GetProperty(font, "ColorFormat");
                WppCom.TrySetProperty(cf, "Type", 1);
                WppCom.TrySetProperty(cf, "RGB", rgb);
            }
            catch (Exception)
            {
            }
        }

        private static void TryWriteTickLabelsRgb(object ticks, int rgb)
        {
            object fc = TryGetTickForeColor(ticks);
            if (fc == null)
            {
                return;
            }

            WppCom.TrySetProperty(fc, "Type", 1);
            WppCom.TrySetProperty(fc, "RGB", rgb);
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

    }
}
