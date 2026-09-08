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
    internal sealed class PptHtmlChartColumn
    {
        public string Role { get; set; }

        public string Name { get; set; }

        public string Color { get; set; }

        public string SeriesType { get; set; }

        public string AxisY { get; set; }

        public string ShowDataLabels { get; set; }

        public string FillGradient { get; set; }

        public string FillAngle { get; set; }

        public string Line { get; set; }

        public string LineWeight { get; set; }

        public string Marker { get; set; }

        public string MarkerSize { get; set; }

        public string MarkerColor { get; set; }

        public string MarkerFill { get; set; }

        public string LabelPosition { get; set; }

        public string LabelFont { get; set; }

        public string LabelSize { get; set; }

        public string LabelColor { get; set; }

        public string LabelFormat { get; set; }
    }

    internal sealed class PptHtmlAxisExtras
    {
        public string Visible { get; set; }

        public string TickFont { get; set; }

        public string TickColor { get; set; }

        public string TickSize { get; set; }

        public string TickPosition { get; set; }

        public string MajorTick { get; set; }

        public string MinorTick { get; set; }

        public string Format { get; set; }

        public string Grid { get; set; }

        public string GridColor { get; set; }

        public string Line { get; set; }

        public string LineWeight { get; set; }
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

        public string ChartStyle { get; set; }

        public string LegendFontColor { get; set; }

        public string ChartAreaColor { get; set; }

        public string Overlap { get; set; }

        public string PlotBox { get; set; }

        public string PlotInside { get; set; }

        public PptHtmlAxisExtras AxisXStyle { get; set; }

        public PptHtmlAxisExtras AxisYStyle { get; set; }

        public PptHtmlAxisExtras AxisY2Style { get; set; }
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

            public int? ChartColor { get; set; }

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

            public bool? GridlineVisible { get; set; }

            public double? GridlineWeight { get; set; }

            public double? GridlineTransparency { get; set; }

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

            /// <summary>饼图等扇区色在 Point 上；Series.Fill 常拍失败。</summary>
            public List<FillSnap> PointFills { get; set; }
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

        private static void ProjectSnapToFormat(PptHtmlChartFormat format, ChartStyleSnap snap)
        {
            if (format == null || snap == null)
            {
                return;
            }

            if (snap.ChartStyle.HasValue)
            {
                format.ChartStyle = snap.ChartStyle.Value.ToString(CultureInfo.InvariantCulture);
            }

            if (snap.HasLegend == false)
            {
                format.Legend = "none";
            }
            else if (snap.LegendPosition.HasValue)
            {
                format.Legend = LegendFromXl(snap.LegendPosition.Value);
            }

            if (!string.IsNullOrEmpty(snap.LegendFontColor))
            {
                format.LegendFontColor = snap.LegendFontColor;
            }

            if (!string.IsNullOrEmpty(snap.TitleFontSize))
            {
                format.TitleFontSize = snap.TitleFontSize;
            }

            if (snap.TitleFontBold.HasValue)
            {
                format.TitleFontBold = snap.TitleFontBold.Value ? "true" : "false";
            }

            if (!string.IsNullOrEmpty(snap.TitleFontColor))
            {
                format.TitleFontColor = snap.TitleFontColor;
            }

            format.ChartAreaColor = AreaColorFromSnap(snap.ChartAreaFillVisible, snap.ChartAreaFillRgb);
            format.PlotColor = AreaColorFromSnap(snap.PlotFillVisible, snap.PlotFillRgb);
            if (snap.GapWidth.HasValue)
            {
                format.GapWidth = snap.GapWidth.Value.ToString(CultureInfo.InvariantCulture);
            }

            if (snap.Overlap.HasValue)
            {
                format.Overlap = snap.Overlap.Value.ToString(CultureInfo.InvariantCulture);
            }

            format.PlotBox = BoxFromSnap(
                snap.PlotLeft, snap.PlotTop, snap.PlotWidth, snap.PlotHeight);
            format.PlotInside = BoxFromSnap(
                snap.PlotInsideLeft, snap.PlotInsideTop, snap.PlotInsideWidth, snap.PlotInsideHeight);
            format.AxisXStyle = AxisExtrasFromSnap(snap.Category);
            format.AxisYStyle = AxisExtrasFromSnap(snap.Value);
            format.AxisY2Style = AxisExtrasFromSnap(snap.ValueSecondary);
            if (format.AxisXStyle != null && !string.IsNullOrEmpty(format.AxisXStyle.Format))
            {
                format.AxisXFormat = format.AxisXStyle.Format;
            }

            if (snap.Category != null && snap.Category.HasTitle == true
                && !string.IsNullOrEmpty(snap.Category.Title))
            {
                format.AxisX = snap.Category.Title;
            }

            if (snap.Value != null && snap.Value.HasTitle == true
                && !string.IsNullOrEmpty(snap.Value.Title))
            {
                format.AxisY = snap.Value.Title;
            }

            if (snap.ValueSecondary != null && snap.ValueSecondary.HasTitle == true
                && !string.IsNullOrEmpty(snap.ValueSecondary.Title))
            {
                format.AxisYSecondary = snap.ValueSecondary.Title;
            }

            if (snap.Value != null && snap.Value.HasMajorGridlines.HasValue)
            {
                format.Gridlines = snap.Value.HasMajorGridlines.Value ? "true" : "false";
            }
        }

        private static void ProjectSnapToColumns(PptHtmlChartGrid grid, ChartStyleSnap snap)
        {
            if (grid == null || grid.Columns == null || snap == null || snap.Series == null)
            {
                return;
            }

            int si = 0;
            for (int i = 0; i < grid.Columns.Count && si < snap.Series.Count; i++)
            {
                PptHtmlChartColumn col = grid.Columns[i];
                if (col == null || col.Role == "category")
                {
                    continue;
                }

                SeriesStyleSnap one = snap.Series[si++];
                if (one.ChartType.HasValue)
                {
                    col.SeriesType = SeriesTypeFromXl(one.ChartType.Value);
                }

                if (one.AxisGroup == XlSecondary)
                {
                    col.AxisY = "secondary";
                }
                else if (one.AxisGroup == XlPrimary)
                {
                    col.AxisY = "primary";
                }

                if (one.Fill != null)
                {
                    if (one.Fill.Stops != null && one.Fill.Stops.Count > 0)
                    {
                        col.FillGradient = EncodeGradient(one.Fill.Stops);
                        if (one.Fill.Angle.HasValue)
                        {
                            col.FillAngle = one.Fill.Angle.Value.ToString("0.##", CultureInfo.InvariantCulture);
                        }
                    }
                    else if (one.Fill.SolidRgb.HasValue)
                    {
                        col.Color = OfficeRgbToHex(one.Fill.SolidRgb.Value);
                    }
                    else if (one.Fill.Visible == false)
                    {
                        col.Color = "none";
                    }
                }

                if (one.Line != null)
                {
                    if (one.Line.Visible == false)
                    {
                        col.Line = "none";
                    }
                    else if (one.Line.Rgb.HasValue)
                    {
                        col.Line = OfficeRgbToHex(one.Line.Rgb.Value);
                    }

                    if (one.Line.Weight.HasValue && !IsPhantomWeight(one.Line.Weight.Value))
                    {
                        col.LineWeight = one.Line.Weight.Value.ToString("0.##", CultureInfo.InvariantCulture);
                    }
                }

                if (one.MarkerStyle.HasValue)
                {
                    col.Marker = MarkerFromXl(one.MarkerStyle.Value);
                }

                if (one.MarkerSize.HasValue)
                {
                    col.MarkerSize = one.MarkerSize.Value.ToString(CultureInfo.InvariantCulture);
                }

                if (one.MarkerForeRgb.HasValue)
                {
                    col.MarkerColor = OfficeRgbToHex(one.MarkerForeRgb.Value);
                }

                if (one.MarkerBackRgb.HasValue)
                {
                    col.MarkerFill = OfficeRgbToHex(one.MarkerBackRgb.Value);
                }

                if (one.HasDataLabels.HasValue)
                {
                    col.ShowDataLabels = one.HasDataLabels.Value ? "true" : "false";
                }

                if (one.DataLabelPosition.HasValue)
                {
                    col.LabelPosition = LabelPosFromXl(one.DataLabelPosition.Value);
                }

                if (!string.IsNullOrEmpty(one.DataLabelFontName))
                {
                    col.LabelFont = one.DataLabelFontName;
                }

                if (one.DataLabelFontSize.HasValue)
                {
                    col.LabelSize = one.DataLabelFontSize.Value.ToString("0.##", CultureInfo.InvariantCulture);
                }

                if (one.DataLabelFontColor.HasValue)
                {
                    col.LabelColor = OfficeRgbToHex(one.DataLabelFontColor.Value);
                }

                if (!string.IsNullOrEmpty(one.DataLabelNumberFormat)
                    && one.DataLabelNumberFormat != ";;;")
                {
                    col.LabelFormat = one.DataLabelNumberFormat;
                }
            }
        }

        private static ChartStyleSnap SnapFromFormat(PptHtmlChartFormat format, PptHtmlChartGrid grid)
        {
            var snap = new ChartStyleSnap { Series = new List<SeriesStyleSnap>() };
            if (format != null)
            {
                if (int.TryParse(format.ChartStyle, NumberStyles.Integer, CultureInfo.InvariantCulture, out int style))
                {
                    snap.ChartStyle = style;
                }

                if (format.Legend != null)
                {
                    if (string.Equals(format.Legend, "none", StringComparison.OrdinalIgnoreCase))
                    {
                        snap.HasLegend = false;
                    }
                    else if (!string.IsNullOrWhiteSpace(format.Legend))
                    {
                        snap.HasLegend = true;
                        snap.LegendPosition = LegendToXl(format.Legend);
                    }
                }

                if (!string.IsNullOrWhiteSpace(format.LegendFontColor))
                {
                    snap.LegendFontColor = format.LegendFontColor;
                }

                if (format.Title != null)
                {
                    snap.HasTitle = !string.IsNullOrEmpty(format.Title);
                }

                if (!string.IsNullOrWhiteSpace(format.TitleFontSize))
                {
                    snap.TitleFontSize = format.TitleFontSize;
                }

                if (!string.IsNullOrWhiteSpace(format.TitleFontBold))
                {
                    snap.TitleFontBold = IsTrue(format.TitleFontBold);
                }

                if (!string.IsNullOrWhiteSpace(format.TitleFontColor))
                {
                    snap.TitleFontColor = format.TitleFontColor;
                }

                ApplyAreaColor(format.ChartAreaColor, out bool? areaVis, out int? areaRgb);
                snap.ChartAreaFillVisible = areaVis;
                snap.ChartAreaFillRgb = areaRgb;
                ApplyAreaColor(format.PlotColor, out bool? plotVis, out int? plotRgb);
                snap.PlotFillVisible = plotVis;
                snap.PlotFillRgb = plotRgb;
                if (int.TryParse(format.GapWidth, NumberStyles.Integer, CultureInfo.InvariantCulture, out int gap))
                {
                    snap.GapWidth = gap;
                }

                if (int.TryParse(format.Overlap, NumberStyles.Integer, CultureInfo.InvariantCulture, out int overlap))
                {
                    snap.Overlap = overlap;
                }

                ParseBox(format.PlotBox, out float? l, out float? t, out float? w, out float? h);
                snap.PlotLeft = l;
                snap.PlotTop = t;
                snap.PlotWidth = w;
                snap.PlotHeight = h;
                ParseBox(format.PlotInside, out float? il, out float? it, out float? iw, out float? ih);
                snap.PlotInsideLeft = il;
                snap.PlotInsideTop = it;
                snap.PlotInsideWidth = iw;
                snap.PlotInsideHeight = ih;
                snap.Category = AxisSnapFromExtras(format.AxisXStyle, format.AxisX, format.AxisXFormat);
                snap.Value = AxisSnapFromExtras(format.AxisYStyle, format.AxisY, null);
                snap.ValueSecondary = AxisSnapFromExtras(format.AxisY2Style, format.AxisYSecondary, null);
                if (!string.IsNullOrWhiteSpace(format.Gridlines) && snap.Value != null
                    && snap.Value.HasMajorGridlines == null)
                {
                    snap.Value.HasMajorGridlines = IsTrue(format.Gridlines);
                }
            }

            if (grid != null && grid.Columns != null)
            {
                foreach (PptHtmlChartColumn col in grid.Columns)
                {
                    if (col == null || col.Role == "category")
                    {
                        continue;
                    }

                    snap.Series.Add(SeriesSnapFromColumn(col));
                }
            }

            return snap;
        }

        private static ChartStyleSnap OverlaySnap(ChartStyleSnap oldSnap, ChartStyleSnap htmlSnap)
        {
            if (oldSnap == null)
            {
                return htmlSnap;
            }

            if (htmlSnap == null)
            {
                return oldSnap;
            }

            if (htmlSnap.ChartStyle.HasValue)
            {
                oldSnap.ChartStyle = htmlSnap.ChartStyle;
            }

            if (htmlSnap.ChartColor.HasValue)
            {
                oldSnap.ChartColor = htmlSnap.ChartColor;
            }

            if (htmlSnap.HasTitle.HasValue)
            {
                oldSnap.HasTitle = htmlSnap.HasTitle;
            }

            if (!string.IsNullOrEmpty(htmlSnap.TitleFontColor))
            {
                oldSnap.TitleFontColor = htmlSnap.TitleFontColor;
            }

            if (!string.IsNullOrEmpty(htmlSnap.TitleFontSize))
            {
                oldSnap.TitleFontSize = htmlSnap.TitleFontSize;
            }

            if (htmlSnap.TitleFontBold.HasValue)
            {
                oldSnap.TitleFontBold = htmlSnap.TitleFontBold;
            }

            if (htmlSnap.HasLegend.HasValue)
            {
                oldSnap.HasLegend = htmlSnap.HasLegend;
            }

            if (htmlSnap.LegendPosition.HasValue)
            {
                oldSnap.LegendPosition = htmlSnap.LegendPosition;
            }

            if (!string.IsNullOrEmpty(htmlSnap.LegendFontColor))
            {
                oldSnap.LegendFontColor = htmlSnap.LegendFontColor;
            }

            if (htmlSnap.ChartAreaFillVisible.HasValue)
            {
                oldSnap.ChartAreaFillVisible = htmlSnap.ChartAreaFillVisible;
                oldSnap.ChartAreaFillRgb = htmlSnap.ChartAreaFillRgb;
            }

            if (htmlSnap.PlotFillVisible.HasValue)
            {
                oldSnap.PlotFillVisible = htmlSnap.PlotFillVisible;
                oldSnap.PlotFillRgb = htmlSnap.PlotFillRgb;
            }

            if (htmlSnap.GapWidth.HasValue)
            {
                oldSnap.GapWidth = htmlSnap.GapWidth;
            }

            if (htmlSnap.Overlap.HasValue)
            {
                oldSnap.Overlap = htmlSnap.Overlap;
            }

            OverlayBox(htmlSnap, oldSnap);
            oldSnap.Category = OverlayAxis(oldSnap.Category, htmlSnap.Category);
            oldSnap.Value = OverlayAxis(oldSnap.Value, htmlSnap.Value);
            oldSnap.ValueSecondary = OverlayAxis(oldSnap.ValueSecondary, htmlSnap.ValueSecondary);
            if (htmlSnap.Series != null && htmlSnap.Series.Count > 0)
            {
                if (oldSnap.Series == null)
                {
                    oldSnap.Series = new List<SeriesStyleSnap>();
                }

                for (int i = 0; i < htmlSnap.Series.Count; i++)
                {
                    if (i < oldSnap.Series.Count)
                    {
                        oldSnap.Series[i] = OverlaySeries(oldSnap.Series[i], htmlSnap.Series[i]);
                    }
                    else
                    {
                        oldSnap.Series.Add(htmlSnap.Series[i]);
                    }
                }
            }

            return oldSnap;
        }

        private static void OverlayBox(ChartStyleSnap src, ChartStyleSnap dest)
        {
            if (src.PlotLeft.HasValue)
            {
                dest.PlotLeft = src.PlotLeft;
            }

            if (src.PlotTop.HasValue)
            {
                dest.PlotTop = src.PlotTop;
            }

            if (src.PlotWidth.HasValue)
            {
                dest.PlotWidth = src.PlotWidth;
            }

            if (src.PlotHeight.HasValue)
            {
                dest.PlotHeight = src.PlotHeight;
            }

            if (src.PlotInsideLeft.HasValue)
            {
                dest.PlotInsideLeft = src.PlotInsideLeft;
            }

            if (src.PlotInsideTop.HasValue)
            {
                dest.PlotInsideTop = src.PlotInsideTop;
            }

            if (src.PlotInsideWidth.HasValue)
            {
                dest.PlotInsideWidth = src.PlotInsideWidth;
            }

            if (src.PlotInsideHeight.HasValue)
            {
                dest.PlotInsideHeight = src.PlotInsideHeight;
            }
        }

        private static AxisStyleSnap OverlayAxis(AxisStyleSnap dest, AxisStyleSnap src)
        {
            if (src == null)
            {
                return dest;
            }

            if (dest == null)
            {
                return null;
            }

            // 有无轴、有无标题、标题文案以旧图为准，不跟 HTML 增删
            if (dest.Deleted == true)
            {
                return dest;
            }

            if (src.Deleted == true)
            {
                return dest;
            }

            if (!string.IsNullOrEmpty(src.TickFontName))
            {
                dest.TickFontName = src.TickFontName;
            }

            if (src.TickFontColor.HasValue)
            {
                dest.TickFontColor = src.TickFontColor;
            }

            if (src.TickFontSize.HasValue)
            {
                dest.TickFontSize = src.TickFontSize;
            }

            if (src.TickLabelPosition.HasValue)
            {
                dest.TickLabelPosition = src.TickLabelPosition;
            }

            if (src.MajorTickMark.HasValue)
            {
                dest.MajorTickMark = src.MajorTickMark;
            }

            if (src.MinorTickMark.HasValue)
            {
                dest.MinorTickMark = src.MinorTickMark;
            }

            if (!string.IsNullOrEmpty(src.NumberFormat))
            {
                dest.NumberFormat = src.NumberFormat;
            }

            if (src.HasMajorGridlines.HasValue)
            {
                dest.HasMajorGridlines = src.HasMajorGridlines;
            }

            if (src.MajorGridlineRgb.HasValue)
            {
                dest.MajorGridlineRgb = src.MajorGridlineRgb;
            }

            if (src.GridlineVisible.HasValue)
            {
                dest.GridlineVisible = src.GridlineVisible;
            }

            if (src.GridlineWeight.HasValue)
            {
                dest.GridlineWeight = src.GridlineWeight;
            }

            if (src.GridlineTransparency.HasValue)
            {
                dest.GridlineTransparency = src.GridlineTransparency;
            }

            if (src.LineVisible.HasValue)
            {
                dest.LineVisible = src.LineVisible;
            }

            if (src.LineRgb.HasValue)
            {
                dest.LineRgb = src.LineRgb;
            }

            if (src.LineWeight.HasValue)
            {
                dest.LineWeight = src.LineWeight;
            }

            return dest;
        }

        private static SeriesStyleSnap OverlaySeries(SeriesStyleSnap dest, SeriesStyleSnap src)
        {
            if (src == null)
            {
                return dest;
            }

            if (dest == null)
            {
                return src;
            }

            if (src.ChartType.HasValue)
            {
                // th data-series-type="pie" 只表示饼族，不能把旧图 pie3d 压成 pie2d。
                // 2D/3D 只听节点 data-chart-type。
                if (!(dest.ChartType.HasValue
                    && IsPieXl(dest.ChartType.Value)
                    && IsPieXl(src.ChartType.Value)))
                {
                    dest.ChartType = src.ChartType;
                }
            }

            if (src.AxisGroup.HasValue)
            {
                dest.AxisGroup = src.AxisGroup;
            }

            dest.Fill = OverlayFill(dest.Fill, src.Fill);
            dest.Line = OverlayLine(dest.Line, src.Line);
            if (src.PointFills != null && src.PointFills.Count > 0)
            {
                dest.PointFills = src.PointFills;
            }
            if (src.MarkerStyle.HasValue)
            {
                dest.MarkerStyle = src.MarkerStyle;
            }

            if (src.MarkerSize.HasValue)
            {
                dest.MarkerSize = src.MarkerSize;
            }

            if (src.MarkerForeRgb.HasValue)
            {
                dest.MarkerForeRgb = src.MarkerForeRgb;
            }

            if (src.MarkerBackRgb.HasValue)
            {
                dest.MarkerBackRgb = src.MarkerBackRgb;
            }

            if (src.HasDataLabels.HasValue)
            {
                dest.HasDataLabels = src.HasDataLabels;
            }

            if (src.DataLabelPosition.HasValue)
            {
                dest.DataLabelPosition = src.DataLabelPosition;
            }

            if (!string.IsNullOrEmpty(src.DataLabelFontName))
            {
                dest.DataLabelFontName = src.DataLabelFontName;
            }

            if (src.DataLabelFontSize.HasValue)
            {
                dest.DataLabelFontSize = src.DataLabelFontSize;
            }

            if (src.DataLabelFontColor.HasValue)
            {
                dest.DataLabelFontColor = src.DataLabelFontColor;
            }

            if (!string.IsNullOrEmpty(src.DataLabelNumberFormat))
            {
                dest.DataLabelNumberFormat = src.DataLabelNumberFormat;
            }

            return dest;
        }

        private static FillSnap OverlayFill(FillSnap dest, FillSnap src)
        {
            if (src == null)
            {
                return dest;
            }

            if (dest == null)
            {
                return src;
            }

            if (src.Visible.HasValue)
            {
                dest.Visible = src.Visible;
            }

            if (src.FillType.HasValue)
            {
                dest.FillType = src.FillType;
            }

            if (src.SolidRgb.HasValue)
            {
                dest.SolidRgb = src.SolidRgb;
            }

            if (src.Angle.HasValue)
            {
                dest.Angle = src.Angle;
            }

            if (src.Stops != null && src.Stops.Count > 0)
            {
                dest.Stops = src.Stops;
            }

            return dest;
        }

        private static LineSnap OverlayLine(LineSnap dest, LineSnap src)
        {
            if (src == null)
            {
                return dest;
            }

            if (dest == null)
            {
                return src;
            }

            if (src.Visible.HasValue)
            {
                dest.Visible = src.Visible;
            }

            if (src.Rgb.HasValue)
            {
                dest.Rgb = src.Rgb;
            }

            if (src.Weight.HasValue)
            {
                dest.Weight = src.Weight;
            }

            if (src.Stops != null && src.Stops.Count > 0)
            {
                dest.Stops = src.Stops;
            }

            return dest;
        }

        private static SeriesStyleSnap SeriesSnapFromColumn(PptHtmlChartColumn col)
        {
            var one = new SeriesStyleSnap();
            if (TryParseSeriesXl(col.SeriesType, col.Marker, out int xl))
            {
                one.ChartType = xl;
            }

            if (string.Equals(col.AxisY, "secondary", StringComparison.OrdinalIgnoreCase))
            {
                one.AxisGroup = XlSecondary;
            }
            else if (string.Equals(col.AxisY, "primary", StringComparison.OrdinalIgnoreCase))
            {
                one.AxisGroup = XlPrimary;
            }

            one.Fill = FillFromColumn(col);
            one.Line = LineFromColumn(col);
            if (!string.IsNullOrWhiteSpace(col.Marker))
            {
                one.MarkerStyle = MarkerToXl(col.Marker);
            }

            if (int.TryParse(col.MarkerSize, NumberStyles.Integer, CultureInfo.InvariantCulture, out int ms))
            {
                one.MarkerSize = ms;
            }

            if (TryParseHexToOffice(col.MarkerColor, out int mFore))
            {
                one.MarkerForeRgb = mFore;
            }

            if (TryParseHexToOffice(col.MarkerFill, out int mBack))
            {
                one.MarkerBackRgb = mBack;
            }

            if (!string.IsNullOrWhiteSpace(col.ShowDataLabels))
            {
                one.HasDataLabels = IsTrue(col.ShowDataLabels);
            }

            if (!string.IsNullOrWhiteSpace(col.LabelPosition))
            {
                one.DataLabelPosition = LabelPosToXl(col.LabelPosition);
            }

            if (!string.IsNullOrWhiteSpace(col.LabelFont))
            {
                one.DataLabelFontName = col.LabelFont;
            }

            if (double.TryParse(col.LabelSize, NumberStyles.Float, CultureInfo.InvariantCulture, out double lsz))
            {
                one.DataLabelFontSize = lsz;
            }

            if (TryParseHexToOffice(col.LabelColor, out int lRgb))
            {
                one.DataLabelFontColor = lRgb;
            }

            if (!string.IsNullOrWhiteSpace(col.LabelFormat))
            {
                one.DataLabelNumberFormat = col.LabelFormat;
            }

            return one;
        }

        private static FillSnap FillFromColumn(PptHtmlChartColumn col)
        {
            List<GradientStopSnap> stops = DecodeGradient(col.FillGradient);
            if (stops != null && stops.Count > 0)
            {
                var fill = new FillSnap
                {
                    Visible = true,
                    FillType = MsoFillGradient,
                    Stops = stops
                };
                if (double.TryParse(col.FillAngle, NumberStyles.Float, CultureInfo.InvariantCulture, out double ang))
                {
                    fill.Angle = ang;
                }

                return fill;
            }

            if (string.Equals(col.Color, "none", StringComparison.OrdinalIgnoreCase))
            {
                return new FillSnap { Visible = false };
            }

            if (TryParseHexToOffice(col.Color, out int rgb))
            {
                return new FillSnap { Visible = true, SolidRgb = rgb };
            }

            return null;
        }

        private static LineSnap LineFromColumn(PptHtmlChartColumn col)
        {
            if (string.IsNullOrWhiteSpace(col.Line) && string.IsNullOrWhiteSpace(col.LineWeight))
            {
                return null;
            }

            var line = new LineSnap();
            if (string.Equals(col.Line, "none", StringComparison.OrdinalIgnoreCase))
            {
                line.Visible = false;
            }
            else if (TryParseHexToOffice(col.Line, out int rgb))
            {
                line.Visible = true;
                line.Rgb = rgb;
            }

            if (double.TryParse(col.LineWeight, NumberStyles.Float, CultureInfo.InvariantCulture, out double w))
            {
                line.Weight = w;
            }

            return line;
        }

        private static PptHtmlAxisExtras AxisExtrasFromSnap(AxisStyleSnap ax)
        {
            if (ax == null)
            {
                return null;
            }

            var extras = new PptHtmlAxisExtras();
            if (ax.Deleted == true)
            {
                extras.Visible = "false";
            }
            else if (ax.Deleted == false)
            {
                extras.Visible = "true";
            }

            extras.TickFont = ax.TickFontName;
            if (ax.TickFontColor.HasValue)
            {
                extras.TickColor = OfficeRgbToHex(ax.TickFontColor.Value);
            }

            if (ax.TickFontSize.HasValue)
            {
                extras.TickSize = ax.TickFontSize.Value.ToString("0.##", CultureInfo.InvariantCulture);
            }

            if (ax.TickLabelPosition.HasValue)
            {
                extras.TickPosition = TickPosFromXl(ax.TickLabelPosition.Value);
            }

            if (ax.MajorTickMark.HasValue)
            {
                extras.MajorTick = TickMarkFromXl(ax.MajorTickMark.Value);
            }

            if (ax.MinorTickMark.HasValue)
            {
                extras.MinorTick = TickMarkFromXl(ax.MinorTickMark.Value);
            }

            extras.Format = ax.NumberFormat;
            if (ax.HasMajorGridlines.HasValue)
            {
                extras.Grid = ax.HasMajorGridlines.Value ? "true" : "false";
            }

            if (ax.MajorGridlineRgb.HasValue)
            {
                extras.GridColor = OfficeRgbToHex(ax.MajorGridlineRgb.Value);
            }

            if (ax.LineVisible == false)
            {
                extras.Line = "none";
            }
            else if (ax.LineRgb.HasValue)
            {
                extras.Line = OfficeRgbToHex(ax.LineRgb.Value);
            }

            if (ax.LineWeight.HasValue && !IsPhantomWeight(ax.LineWeight.Value))
            {
                extras.LineWeight = ax.LineWeight.Value.ToString("0.##", CultureInfo.InvariantCulture);
            }

            return extras;
        }

        private static AxisStyleSnap AxisSnapFromExtras(PptHtmlAxisExtras extras, string title, string formatFallback)
        {
            if (extras == null && title == null && string.IsNullOrWhiteSpace(formatFallback))
            {
                return null;
            }

            extras = extras ?? new PptHtmlAxisExtras();
            var ax = new AxisStyleSnap();
            if (!string.IsNullOrWhiteSpace(extras.Visible))
            {
                ax.Deleted = !IsTrue(extras.Visible);
            }

            if (title != null)
            {
                ax.HasTitle = !string.IsNullOrEmpty(title);
                ax.Title = title;
            }

            ax.TickFontName = extras.TickFont;
            if (TryParseHexToOffice(extras.TickColor, out int tickRgb))
            {
                ax.TickFontColor = tickRgb;
            }

            if (double.TryParse(extras.TickSize, NumberStyles.Float, CultureInfo.InvariantCulture, out double tsz))
            {
                ax.TickFontSize = tsz;
            }

            if (!string.IsNullOrWhiteSpace(extras.TickPosition))
            {
                ax.TickLabelPosition = TickPosToXl(extras.TickPosition);
            }

            if (!string.IsNullOrWhiteSpace(extras.MajorTick))
            {
                ax.MajorTickMark = TickMarkToXl(extras.MajorTick);
            }

            if (!string.IsNullOrWhiteSpace(extras.MinorTick))
            {
                ax.MinorTickMark = TickMarkToXl(extras.MinorTick);
            }

            ax.NumberFormat = !string.IsNullOrWhiteSpace(extras.Format) ? extras.Format : formatFallback;
            if (!string.IsNullOrWhiteSpace(extras.Grid))
            {
                ax.HasMajorGridlines = IsTrue(extras.Grid);
            }

            if (TryParseHexToOffice(extras.GridColor, out int gRgb))
            {
                ax.MajorGridlineRgb = gRgb;
            }

            if (string.Equals(extras.Line, "none", StringComparison.OrdinalIgnoreCase))
            {
                ax.LineVisible = false;
            }
            else if (TryParseHexToOffice(extras.Line, out int lRgb))
            {
                ax.LineVisible = true;
                ax.LineRgb = lRgb;
            }

            if (double.TryParse(extras.LineWeight, NumberStyles.Float, CultureInfo.InvariantCulture, out double lw))
            {
                ax.LineWeight = lw;
            }

            return ax;
        }

        private static string AreaColorFromSnap(bool? visible, int? rgb)
        {
            if (visible == false)
            {
                return "none";
            }

            if (visible == true && rgb.HasValue)
            {
                return OfficeRgbToHex(rgb.Value);
            }

            return null;
        }

        private static void ApplyAreaColor(string raw, out bool? visible, out int? rgb)
        {
            visible = null;
            rgb = null;
            if (string.IsNullOrWhiteSpace(raw))
            {
                return;
            }

            if (string.Equals(raw, "none", StringComparison.OrdinalIgnoreCase)
                || string.Equals(raw, "transparent", StringComparison.OrdinalIgnoreCase))
            {
                visible = false;
                return;
            }

            if (TryParseHexToOffice(raw, out int parsed))
            {
                visible = true;
                rgb = parsed;
            }
        }

        private static string BoxFromSnap(float? left, float? top, float? width, float? height)
        {
            if (!left.HasValue || !top.HasValue || !width.HasValue || !height.HasValue)
            {
                return null;
            }

            return left.Value.ToString("0.##", CultureInfo.InvariantCulture)
                + "," + top.Value.ToString("0.##", CultureInfo.InvariantCulture)
                + "," + width.Value.ToString("0.##", CultureInfo.InvariantCulture)
                + "," + height.Value.ToString("0.##", CultureInfo.InvariantCulture);
        }

        private static void ParseBox(string raw, out float? left, out float? top, out float? width, out float? height)
        {
            left = top = width = height = null;
            if (string.IsNullOrWhiteSpace(raw))
            {
                return;
            }

            string[] parts = raw.Split(new[] { ',', ' ' }, StringSplitOptions.RemoveEmptyEntries);
            if (parts.Length < 4)
            {
                return;
            }

            if (float.TryParse(parts[0], NumberStyles.Float, CultureInfo.InvariantCulture, out float l)
                && float.TryParse(parts[1], NumberStyles.Float, CultureInfo.InvariantCulture, out float t)
                && float.TryParse(parts[2], NumberStyles.Float, CultureInfo.InvariantCulture, out float w)
                && float.TryParse(parts[3], NumberStyles.Float, CultureInfo.InvariantCulture, out float h))
            {
                left = l;
                top = t;
                width = w;
                height = h;
            }
        }

        private static string EncodeGradient(List<GradientStopSnap> stops)
        {
            if (stops == null || stops.Count == 0)
            {
                return null;
            }

            var sb = new StringBuilder();
            for (int i = 0; i < stops.Count; i++)
            {
                if (i > 0)
                {
                    sb.Append(';');
                }

                GradientStopSnap s = stops[i];
                sb.Append(s.Position.ToString("0.##", CultureInfo.InvariantCulture))
                    .Append(':')
                    .Append(OfficeRgbToHex(s.Rgb))
                    .Append('@')
                    .Append(s.Transparency.ToString("0.##", CultureInfo.InvariantCulture));
            }

            return sb.ToString();
        }

        private static List<GradientStopSnap> DecodeGradient(string raw)
        {
            if (string.IsNullOrWhiteSpace(raw))
            {
                return null;
            }

            var list = new List<GradientStopSnap>();
            foreach (string part in raw.Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries))
            {
                int colon = part.IndexOf(':');
                int at = part.LastIndexOf('@');
                if (colon <= 0)
                {
                    continue;
                }

                if (!double.TryParse(part.Substring(0, colon).Trim(), NumberStyles.Float, CultureInfo.InvariantCulture, out double pos))
                {
                    continue;
                }

                string hex = at > colon
                    ? part.Substring(colon + 1, at - colon - 1).Trim()
                    : part.Substring(colon + 1).Trim();
                if (!TryParseHexToOffice(hex, out int rgb))
                {
                    continue;
                }

                double trans = 0;
                if (at > colon)
                {
                    double.TryParse(part.Substring(at + 1).Trim(), NumberStyles.Float, CultureInfo.InvariantCulture, out trans);
                }

                list.Add(new GradientStopSnap { Position = pos, Rgb = rgb, Transparency = trans });
            }

            return list.Count == 0 ? null : list;
        }

        private static bool TryParseSeriesXl(string seriesType, string marker, out int xl)
        {
            xl = XlColumnClustered;
            if (!string.IsNullOrWhiteSpace(seriesType))
            {
                if (int.TryParse(seriesType.Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out int raw)
                    && raw != 0)
                {
                    xl = raw;
                    return true;
                }

                if (TryParseType(seriesType, out xl, out _, out _))
                {
                    if (xl == XlLine && !string.IsNullOrWhiteSpace(marker)
                        && !string.Equals(marker, "none", StringComparison.OrdinalIgnoreCase))
                    {
                        xl = XlLineMarkers;
                    }

                    return true;
                }

                return false;
            }

            if (!string.IsNullOrWhiteSpace(marker)
                && !string.Equals(marker, "none", StringComparison.OrdinalIgnoreCase))
            {
                xl = XlLineMarkers;
                return true;
            }

            return false;
        }

        private static string SeriesTypeFromXl(int xl)
        {
            if (xl == XlBarClustered)
            {
                return "bar";
            }

            if (xl == XlLine || xl == XlLineMarkers)
            {
                return "line";
            }

            if (xl == Xl3DPie)
            {
                return "pie3d";
            }

            if (xl == XlPie)
            {
                return "pie2d";
            }

            if (xl == XlColumnClustered)
            {
                return "column";
            }

            return xl.ToString(CultureInfo.InvariantCulture);
        }

        private static string MarkerFromXl(int style)
        {
            switch (style)
            {
                case -4142:
                    return "none";
                case 8:
                    return "circle";
                case 2:
                    return "diamond";
                case 1:
                    return "square";
                case 3:
                    return "triangle";
                case 5:
                    return "star";
                case 9:
                    return "plus";
                case -4168:
                    return "x";
                case -4118:
                    return "dot";
                default:
                    return style.ToString(CultureInfo.InvariantCulture);
            }
        }

        private static int MarkerToXl(string raw)
        {
            switch ((raw ?? "").Trim().ToLowerInvariant())
            {
                case "none":
                    return -4142;
                case "circle":
                    return 8;
                case "diamond":
                    return 2;
                case "square":
                    return 1;
                case "triangle":
                    return 3;
                case "star":
                    return 5;
                case "plus":
                    return 9;
                case "x":
                    return -4168;
                case "dot":
                    return -4118;
                default:
                    return int.TryParse(raw, NumberStyles.Integer, CultureInfo.InvariantCulture, out int n)
                        ? n
                        : 8;
            }
        }

        private static string TickPosFromXl(int pos)
        {
            if (pos == XlTickLabelPositionNone)
            {
                return "none";
            }

            if (pos == -4134)
            {
                return "low";
            }

            if (pos == -4127)
            {
                return "high";
            }

            return "next";
        }

        private static int TickPosToXl(string raw)
        {
            switch ((raw ?? "").Trim().ToLowerInvariant())
            {
                case "none":
                    return XlTickLabelPositionNone;
                case "low":
                    return -4134;
                case "high":
                    return -4127;
                default:
                    return XlTickLabelPositionNextToAxis;
            }
        }

        private static string TickMarkFromXl(int mark)
        {
            if (mark == XlTickMarkNone)
            {
                return "none";
            }

            if (mark == 2)
            {
                return "inside";
            }

            if (mark == 4)
            {
                return "cross";
            }

            return "outside";
        }

        private static int TickMarkToXl(string raw)
        {
            switch ((raw ?? "").Trim().ToLowerInvariant())
            {
                case "none":
                    return XlTickMarkNone;
                case "inside":
                    return 2;
                case "cross":
                    return 4;
                default:
                    return 3;
            }
        }

        private static string LabelPosFromXl(int pos)
        {
            switch (pos)
            {
                case -4108:
                    return "center";
                case 0:
                    return "above";
                case 1:
                    return "below";
                case -4131:
                    return "left";
                case -4152:
                    return "right";
                case 2:
                    return "outside";
                case 3:
                    return "inside-end";
                case 4:
                    return "inside-base";
                default:
                    return pos.ToString(CultureInfo.InvariantCulture);
            }
        }

        private static int LabelPosToXl(string raw)
        {
            switch ((raw ?? "").Trim().ToLowerInvariant())
            {
                case "center":
                    return -4108;
                case "above":
                    return 0;
                case "below":
                    return 1;
                case "left":
                    return -4131;
                case "right":
                    return -4152;
                case "outside":
                    return 2;
                case "inside-end":
                    return 3;
                case "inside-base":
                    return 4;
                default:
                    return int.TryParse(raw, NumberStyles.Integer, CultureInfo.InvariantCulture, out int n)
                        ? n
                        : 0;
            }
        }

        private static bool IsPhantomWeight(double weight)
        {
            return weight < -1000 || weight > 1000;
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
                case "pie2d":
                    xlType = XlPie;
                    canonical = "pie2d";
                    return true;
                case "pie3d":
                    xlType = Xl3DPie;
                    canonical = "pie3d";
                    return true;
                default:
                    error = "data-chart-type 仅支持 column/bar/line/pie2d/pie3d（pie 为 pie2d 别名）";
                    return false;
            }
        }

        public static string CanonicalTypeFromXl(int xlType)
        {
            if (xlType == XlBarClustered)
            {
                return "bar";
            }

            if (xlType == XlLine || xlType == XlLineMarkers)
            {
                return "line";
            }

            if (xlType == Xl3DPie)
            {
                return "pie3d";
            }

            if (xlType == XlPie)
            {
                return "pie2d";
            }

            return "column";
        }

        private static bool IsPieXl(int xlType)
        {
            return xlType == XlPie || xlType == Xl3DPie;
        }

        private static bool IsPieChart(object chart)
        {
            if (chart == null)
            {
                return false;
            }

            try
            {
                object t = WppCom.GetProperty(chart, "ChartType");
                return t != null && IsPieXl(Convert.ToInt32(t));
            }
            catch (Exception)
            {
                return false;
            }
        }

        private static bool TryReadChartXl(object chart, out int xlType)
        {
            xlType = XlColumnClustered;
            if (chart == null)
            {
                return false;
            }

            try
            {
                object t = WppCom.GetProperty(chart, "ChartType");
                if (t == null)
                {
                    return false;
                }

                xlType = Convert.ToInt32(t);
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        private static void PinChartSeriesType(ChartStyleSnap snap, int xlType)
        {
            if (snap == null || snap.Series == null)
            {
                return;
            }

            for (int i = 0; i < snap.Series.Count; i++)
            {
                if (snap.Series[i] != null)
                {
                    snap.Series[i].ChartType = xlType;
                }
            }
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
                            Color = ParseColorOrNone(GetAttr(cells[i], "data-color")),
                            SeriesType = GetAttr(cells[i], "data-series-type"),
                            AxisY = GetAttr(cells[i], "data-axis-y"),
                            ShowDataLabels = GetAttr(cells[i], "data-show-data-labels"),
                            FillGradient = GetAttr(cells[i], "data-fill-gradient"),
                            FillAngle = GetAttr(cells[i], "data-fill-angle"),
                            Line = GetAttr(cells[i], "data-line"),
                            LineWeight = GetAttr(cells[i], "data-line-weight"),
                            Marker = GetAttr(cells[i], "data-marker"),
                            MarkerSize = GetAttr(cells[i], "data-marker-size"),
                            MarkerColor = GetAttr(cells[i], "data-marker-color"),
                            MarkerFill = GetAttr(cells[i], "data-marker-fill"),
                            LabelPosition = GetAttr(cells[i], "data-label-position"),
                            LabelFont = GetAttr(cells[i], "data-label-font"),
                            LabelSize = GetAttr(cells[i], "data-label-size"),
                            LabelColor = GetAttr(cells[i], "data-label-color"),
                            LabelFormat = GetAttr(cells[i], "data-label-format")
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
                AxisYSecondary = GetAttr(el, "data-axis-y-secondary"),
                ChartStyle = GetAttr(el, "data-chart-style"),
                LegendFontColor = GetAttr(el, "data-legend-font-color"),
                ChartAreaColor = GetAttr(el, "data-chart-area-color"),
                Overlap = GetAttr(el, "data-overlap"),
                PlotBox = GetAttr(el, "data-plot-box"),
                PlotInside = GetAttr(el, "data-plot-inside"),
                AxisXStyle = ParseAxisExtras(el, "data-axis-x"),
                AxisYStyle = ParseAxisExtras(el, "data-axis-y"),
                AxisY2Style = ParseAxisExtras(el, "data-axis-y2")
            };
        }

        private static PptHtmlAxisExtras ParseAxisExtras(XElement el, string prefix)
        {
            return new PptHtmlAxisExtras
            {
                Visible = GetAttr(el, prefix + "-visible"),
                TickFont = GetAttr(el, prefix + "-tick-font"),
                TickColor = GetAttr(el, prefix + "-tick-color"),
                TickSize = GetAttr(el, prefix + "-tick-size"),
                TickPosition = GetAttr(el, prefix + "-tick-position"),
                MajorTick = GetAttr(el, prefix + "-major-tick"),
                MinorTick = GetAttr(el, prefix + "-minor-tick"),
                Format = prefix == "data-axis-x"
                    ? GetAttr(el, "data-axis-x-format")
                    : GetAttr(el, prefix + "-format"),
                Grid = GetAttr(el, prefix + "-grid"),
                GridColor = GetAttr(el, prefix + "-grid-color"),
                Line = GetAttr(el, prefix + "-line"),
                LineWeight = GetAttr(el, prefix + "-line-weight")
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

                WriteRawAttr(sb, "data-fill-gradient", col.FillGradient);
                WriteRawAttr(sb, "data-fill-angle", col.FillAngle);
                WriteRawAttr(sb, "data-line", col.Line);
                WriteRawAttr(sb, "data-line-weight", col.LineWeight);
                WriteRawAttr(sb, "data-marker", col.Marker);
                WriteRawAttr(sb, "data-marker-size", col.MarkerSize);
                WriteRawAttr(sb, "data-marker-color", col.MarkerColor);
                WriteRawAttr(sb, "data-marker-fill", col.MarkerFill);
                WriteRawAttr(sb, "data-label-position", col.LabelPosition);
                WriteRawAttr(sb, "data-label-font", col.LabelFont);
                WriteRawAttr(sb, "data-label-size", col.LabelSize);
                WriteRawAttr(sb, "data-label-color", col.LabelColor);
                WriteRawAttr(sb, "data-label-format", col.LabelFormat);

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
            WriteAttr(sb, "data-chart-style", fmt.ChartStyle);
            WriteAttr(sb, "data-legend-font-color", fmt.LegendFontColor);
            WriteAttr(sb, "data-chart-area-color", fmt.ChartAreaColor);
            WriteAttr(sb, "data-overlap", fmt.Overlap);
            WriteAttr(sb, "data-plot-box", fmt.PlotBox);
            WriteAttr(sb, "data-plot-inside", fmt.PlotInside);
            WriteAxisExtras(sb, "data-axis-x", fmt.AxisXStyle, skipFormat: true);
            WriteAxisExtras(sb, "data-axis-y", fmt.AxisYStyle, skipFormat: false);
            WriteAxisExtras(sb, "data-axis-y2", fmt.AxisY2Style, skipFormat: false);
        }

        private static void WriteAxisExtras(StringBuilder sb, string prefix, PptHtmlAxisExtras ax, bool skipFormat)
        {
            if (ax == null)
            {
                return;
            }

            WriteAttr(sb, prefix + "-visible", ax.Visible);
            WriteAttr(sb, prefix + "-tick-font", ax.TickFont);
            WriteAttr(sb, prefix + "-tick-color", ax.TickColor);
            WriteAttr(sb, prefix + "-tick-size", ax.TickSize);
            WriteAttr(sb, prefix + "-tick-position", ax.TickPosition);
            WriteAttr(sb, prefix + "-major-tick", ax.MajorTick);
            WriteAttr(sb, prefix + "-minor-tick", ax.MinorTick);
            if (!skipFormat)
            {
                WriteAttr(sb, prefix + "-format", ax.Format);
            }

            WriteAttr(sb, prefix + "-grid", ax.Grid);
            WriteAttr(sb, prefix + "-grid-color", ax.GridColor);
            WriteAttr(sb, prefix + "-line", ax.Line);
            WriteAttr(sb, prefix + "-line-weight", ax.LineWeight);
        }

        private static void WriteRawAttr(StringBuilder sb, string name, string value)
        {
            if (string.IsNullOrEmpty(value))
            {
                return;
            }

            sb.Append(" ").Append(name).Append("=\"").Append(EscapeAttr(value)).Append("\"");
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
            ChartStyleSnap snap = TryCaptureStyle(chart);

            if (!TryReadGrid(chart, out PptHtmlChartGrid grid, out error))
            {
                return false;
            }

            ProjectSnapToFormat(format, snap);
            ProjectSnapToColumns(grid, snap);
            model = new PptHtmlChartReadModel
            {
                Format = format,
                Grid = grid
            };
            return true;
        }

        public static bool TryPourGrid(
            object chart,
            PptHtmlChartGrid grid,
            out string error,
            List<string> warnings = null)
        {
            error = null;
            if (chart == null || grid == null || !grid.IsPourable)
            {
                error = "新建 chart 必须在节点内嵌 <table> 灌数，不能建空图";
                return false;
            }

            PourLog(warnings, "开始灌数（按 Word 建图写 Series） " + DescribeWantGrid(grid));
            object excelApp = null;
            try
            {
                // 改已有图也是新建：AddChart2 之后按 Word FillChartData 写 Values/XValues。
                // ChartData / SetSourceData 只在点数扩不上时兜底。
                int wantSeries = CountValueColumns(grid);
                if (!TryEnsureSeriesCount(chart, wantSeries, out error))
                {
                    return false;
                }

                TrimExtraSeries(chart, wantSeries, out _);
                if (!TryPourSeriesLikeWord(chart, grid, out error))
                {
                    PourLog(warnings, "写 Series 失败: " + (error ?? ""));
                    return false;
                }

                PourLog(warnings, "Word 灌数后 " + DescribeLiveSeries(chart));
                if (!SeriesRowCountMatches(chart, grid))
                {
                    PourLog(warnings, "Series 未扩到 " + grid.Rows.Count + " 行，兜底写内嵌表");
                    if (!TryExpandViaChartData(chart, grid, out excelApp, out error, warnings))
                    {
                        PourLog(warnings, "ChartData 兜底失败: " + (error ?? ""));
                        return false;
                    }

                    TrimExtraSeries(chart, wantSeries, out _);
                    if (!TryPourSeriesLikeWord(chart, grid, out error))
                    {
                        PourLog(warnings, "兜底后再写 Series 失败: " + (error ?? ""));
                        return false;
                    }

                    PourLog(warnings, "兜底后 " + DescribeLiveSeries(chart));
                }

                if (!TryVerifyPouredGrid(chart, grid, out error, warnings))
                {
                    PourLog(warnings, "灌后校验失败: " + (error ?? ""));
                    return false;
                }

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
                TryHideChartExcel(chart);
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
        /// 改已有图 = 新建图。先拍旧图属性，AddChart2 灌数（Word 同路），
        /// HTML 没写的属性用旧图快照补上，再删旧图。
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

            // 组内 chart 的 HTML style 是相对组的 %；瘦稿平铺到 section 后会当成整页坐标
            //（常见 top:0% → 飞到页顶）。换数沿用原图幻灯片盒，不听稿上的几何。
            float useLeft;
            float useTop;
            float useWidth;
            float useHeight;
            if (IsInsideGroup(oldShape))
            {
                if (left.HasValue || top.HasValue || width.HasValue || height.HasValue)
                {
                    StyleLog(warnings, "组内 chart 忽略稿上 style，沿用原图幻灯片坐标 "
                        + oldLeft.ToString("0.#", CultureInfo.InvariantCulture) + ","
                        + oldTop.ToString("0.#", CultureInfo.InvariantCulture));
                }

                useLeft = oldLeft;
                useTop = oldTop;
                useWidth = oldWidth;
                useHeight = oldHeight;
            }
            else
            {
                useLeft = left ?? oldLeft;
                useTop = top ?? oldTop;
                useWidth = width ?? oldWidth;
                useHeight = height ?? oldHeight;
            }

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
            bool htmlWroteType = format != null && !string.IsNullOrWhiteSpace(format.ChartType);

            ChartStyleSnap oldSnap = null;
            try
            {
                oldSnap = TryCaptureStyle(oldChart, warnings, "旧图");
            }
            catch (Exception ex)
            {
                StyleLog(warnings, "拍旧图样式异常: " + ex.Message);
            }

            ChartStyleSnap htmlSnap = SnapFromFormat(useFormat, useGrid);
            ChartStyleSnap snap = OverlaySnap(oldSnap, htmlSnap);
            if (snap == null)
            {
                snap = htmlSnap ?? oldSnap;
            }

            int xlType;
            if (htmlWroteType)
            {
                if (!TryParseType(useFormat.ChartType, out xlType, out string canon, out error))
                {
                    return false;
                }

                useFormat.ChartType = canon;
                PinChartSeriesType(snap, xlType);
            }
            else if (!TryReadChartXl(oldChart, out xlType))
            {
                if (!TryParseType(useFormat.ChartType, out xlType, out _, out error))
                {
                    return false;
                }
            }
            else
            {
                useFormat.ChartType = CanonicalTypeFromXl(xlType);
            }

            StyleLog(warnings, "重建开始 type=" + (useFormat.ChartType ?? "")
                + " xl=" + xlType
                + (htmlWroteType ? " from=html" : " from=old")
                + " box=" + useLeft.ToString("0.#", CultureInfo.InvariantCulture)
                + "," + useTop.ToString("0.#", CultureInfo.InvariantCulture)
                + " " + useWidth.ToString("0.#", CultureInfo.InvariantCulture)
                + "x" + useHeight.ToString("0.#", CultureInfo.InvariantCulture));

            SyncFormatToOldSnap(useFormat, xlType, warnings);
            MergeGridlinesFromFormat(snap, useFormat);

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
                -1,
                newLayout: false,
                applyHtmlChrome: false))
            {
                return false;
            }

            object newChart = TryGetChart(newShape);
            FinishLineChartLayout(newChart, xlType, useFormat, warnings);
            TryApplyStyleSnap(newChart, snap, warnings, useGrid, useFormat);
            PourLog(warnings, "套快照后 " + DescribeLiveSeries(newChart));
            // 套快照（含 3D）可能把数据打回字面量；必须再验，对不上再灌一次，仍不对就失败并删新图。
            if (!TryVerifyPouredGrid(newChart, useGrid, out error, warnings))
            {
                StyleLog(warnings, "套回后数据对不上，再按新建图灌数: " + (error ?? ""));
                if (!TryPourGrid(newChart, useGrid, out error, warnings)
                    || !TryVerifyPouredGrid(newChart, useGrid, out error, warnings))
                {
                    TryDelete(newShape);
                    newShape = null;
                    if (string.IsNullOrEmpty(error))
                    {
                        error = "图表换数后内嵌表与稿不一致";
                    }

                    return false;
                }
            }

            // 轴皮是最后一层：前面结构/开标签/再灌数都可能按 ChartStyle 掀字色。
            TryInheritAxisChrome(newChart, snap, warnings);
            try
            {
                TryCaptureStyle(newChart, warnings, "新图套回后");
            }
            catch (Exception ex)
            {
                StyleLog(warnings, "回读新图异常: " + ex.Message);
            }

            TryHideChartExcel(newChart);
            DismissChartExcelUi();
            TryDelete(oldShape);
            warnings?.Add("已按旧图属性新建图表（HTML 未写的属性用快照补上）");
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
            PptHtmlChartFormat format = null)
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
                TryInheritSeriesStructure(chart, snap);
                TryApplyChartTheme(chart, snap);
                TryInheritTitleAndLegend(chart, snap);
                TryWriteAreaFill(chart, "ChartArea", snap.ChartAreaFillVisible, snap.ChartAreaFillRgb);
                TryWriteAreaFill(chart, "PlotArea", snap.PlotFillVisible, snap.PlotFillRgb);
                TryApplyChartGroup(chart, snap);
                TryInheritAxes(chart, snap, grid, format, warnings);
                TryInheritColors(chart, snap, warnings);
                TryInheritSeriesLineAndMarker(chart, snap, warnings);
                TryRestorePieChartType(chart, snap);
                TryApplyPlotLayout(chart, snap);
                TryInheritDataLabels(chart, snap, warnings);
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

        private static void TryInheritTitleAndLegend(object chart, ChartStyleSnap snap)
        {
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

            if (!snap.HasLegend.HasValue)
            {
                return;
            }

            WppCom.TrySetProperty(chart, "HasLegend", snap.HasLegend.Value);
            if (!snap.HasLegend.Value)
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
            if (grid == null || grid.Rows == null || grid.Columns == null)
            {
                return "稿=null";
            }

            var cats = new List<string>();
            var vals = new List<string>();
            for (int r = 0; r < grid.Rows.Count; r++)
            {
                List<string> row = grid.Rows[r];
                cats.Add(row != null && row.Count > 0 ? row[0] ?? "" : "");
                if (row != null && row.Count > 1)
                {
                    vals.Add(row[1] ?? "");
                }
            }

            return "稿 " + grid.Rows.Count + "x" + grid.Columns.Count
                + " cats=[" + string.Join(",", cats) + "]"
                + " vals=[" + string.Join(",", vals) + "]";
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
                    FinishLineChartLayout(chart, xlType, format, warnings);
                    TryApplyStyleSnap(chart, htmlSnap, warnings, grid, format);
                    TryInheritAxisChrome(chart, htmlSnap, warnings);
                }
                else
                {
                    EnsureCategoryAxisLabels(chart, grid);
                }

                return true;
            }
            finally
            {
                TryHideChartExcel(chart);
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

                    if (!IsPieChart(chart)
                        && !string.IsNullOrEmpty(col.Color)
                        && TryParseHexToOffice(col.Color, out int rgb))
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

        private const int SwHide = 0;
        private const uint WmClose = 0x0010;
        private const uint WmSysCommand = 0x0112;
        private static readonly IntPtr ScClose = new IntPtr(0xF060);

        private delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

        [DllImport("user32.dll")]
        private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        [DllImport("user32.dll")]
        private static extern bool IsWindowVisible(IntPtr hWnd);

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

        [DllImport("user32.dll")]
        private static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);

        [DllImport("user32.dll")]
        private static extern bool PostMessage(IntPtr hWnd, uint msg, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll")]
        private static extern IntPtr SendMessage(IntPtr hWnd, uint msg, IntPtr wParam, IntPtr lParam);

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

            TryHideExcelAppHwnd(excelApp);
        }

        /// <summary>
        /// 藏图表拉起的内嵌 Excel，不 Quit，避免弄坏包内 embeddings。
        /// COM Visible=false 常只藏内容，PowerPoint 会留下空白编辑框，须再关 HWND。
        /// </summary>
        private static void TryHideChartExcel(object chart)
        {
            if (chart != null)
            {
                try
                {
                    object chartData = WppCom.GetProperty(chart, "ChartData");
                    object workbook = chartData == null ? null : WppCom.GetProperty(chartData, "Workbook");
                    object excelApp = workbook == null ? null : WppCom.GetProperty(workbook, "Application");
                    HideEmbeddedExcel(excelApp);
                }
                catch (Exception)
                {
                }
            }

            TryHidePowerPointChartExcelWindows();
            TryCloseChartExcelHwnds(forceClose: false);
        }

        /// <summary>
        /// apply 整页结束后再清一次：AddChart2 常在 COM 返回后才把编辑框画出来。
        /// </summary>
        public static void DismissChartExcelUi()
        {
            for (int i = 0; i < 6; i++)
            {
                TryHidePowerPointChartExcelWindows();
                TryCloseChartExcelHwnds(forceClose: i >= 2);
                if (!HasVisibleChartExcelWindow())
                {
                    return;
                }

                try
                {
                    Application.DoEvents();
                }
                catch (Exception)
                {
                }

                Thread.Sleep(80);
            }
        }

        private static void TryHidePowerPointChartExcelWindows()
        {
            foreach (string progId in new[] { "Excel.Application", "Ket.Application", "et.Application" })
            {
                object excelApp = null;
                try
                {
                    excelApp = Marshal.GetActiveObject(progId);
                }
                catch (Exception)
                {
                    continue;
                }

                if (excelApp == null)
                {
                    continue;
                }

                try
                {
                    object windows = WppCom.GetProperty(excelApp, "Windows");
                    int count = Convert.ToInt32(WppCom.GetProperty(windows, "Count"));
                    bool hidChartWindow = false;
                    int stillVisible = 0;
                    for (int i = 1; i <= count; i++)
                    {
                        object win = WppCom.GetIndexed(windows, i);
                        string caption = Convert.ToString(WppCom.GetProperty(win, "Caption") ?? "");
                        if (IsPowerPointChartExcelCaption(caption))
                        {
                            WppCom.TrySetProperty(win, "Visible", false);
                            hidChartWindow = true;
                        }
                        else if (IsTruthy(WppCom.GetProperty(win, "Visible")))
                        {
                            stillVisible++;
                        }
                    }

                    if (hidChartWindow && stillVisible == 0)
                    {
                        SuppressExcel(excelApp);
                    }

                    TryHideExcelAppHwnd(excelApp);
                }
                catch (Exception)
                {
                }
            }
        }

        private static void TryHideExcelAppHwnd(object excelApp)
        {
            if (excelApp == null)
            {
                return;
            }

            try
            {
                object raw = WppCom.GetProperty(excelApp, "Hwnd");
                if (raw == null)
                {
                    return;
                }

                IntPtr hwnd = new IntPtr(Convert.ToInt64(raw, CultureInfo.InvariantCulture));
                if (hwnd == IntPtr.Zero)
                {
                    return;
                }

                string title = GetWindowTitle(hwnd);
                if (!IsPowerPointChartExcelCaption(title))
                {
                    return;
                }

                ShowWindow(hwnd, SwHide);
                PostMessage(hwnd, WmClose, IntPtr.Zero, IntPtr.Zero);
            }
            catch (Exception)
            {
            }
        }

        private static bool TryCloseChartExcelHwnds(bool forceClose)
        {
            List<IntPtr> targets = FindChartExcelHwnds(visibleOnly: !forceClose);
            if (targets.Count == 0 && forceClose)
            {
                targets = FindChartExcelHwnds(visibleOnly: false);
            }

            foreach (IntPtr hwnd in targets)
            {
                try
                {
                    ShowWindow(hwnd, SwHide);
                    PostMessage(hwnd, WmClose, IntPtr.Zero, IntPtr.Zero);
                    if (forceClose && IsWindowVisible(hwnd))
                    {
                        SendMessage(hwnd, WmSysCommand, ScClose, IntPtr.Zero);
                    }
                }
                catch (Exception)
                {
                }
            }

            return targets.Count > 0;
        }

        private static bool HasVisibleChartExcelWindow()
        {
            return FindChartExcelHwnds(visibleOnly: true).Count > 0;
        }

        private static List<IntPtr> FindChartExcelHwnds(bool visibleOnly)
        {
            var found = new List<IntPtr>();
            try
            {
                EnumWindows((hWnd, _) =>
                {
                    if (visibleOnly && !IsWindowVisible(hWnd))
                    {
                        return true;
                    }

                    if (IsPowerPointChartExcelCaption(GetWindowTitle(hWnd)))
                    {
                        found.Add(hWnd);
                    }

                    return true;
                }, IntPtr.Zero);
            }
            catch (Exception)
            {
            }

            return found;
        }

        private static string GetWindowTitle(IntPtr hWnd)
        {
            try
            {
                var sb = new StringBuilder(512);
                if (GetWindowText(hWnd, sb, sb.Capacity) <= 0)
                {
                    return "";
                }

                return sb.ToString();
            }
            catch (Exception)
            {
                return "";
            }
        }

        private static bool IsPowerPointChartExcelCaption(string caption)
        {
            if (string.IsNullOrEmpty(caption))
            {
                return false;
            }

            if (caption.IndexOf("PowerPoint 中的图表", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return true;
            }

            if (caption.IndexOf("Chart in Microsoft PowerPoint", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return true;
            }

            if (caption.IndexOf("中的图表 - Excel", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return true;
            }

            if (caption.IndexOf("中的图表 - WPS", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return true;
            }

            if (caption.IndexOf("Chart in WPS", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return true;
            }

            if (caption.IndexOf("演示中的图表", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return true;
            }

            if (caption.IndexOf("演示文稿", StringComparison.OrdinalIgnoreCase) >= 0
                && caption.IndexOf("图表", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return true;
            }

            return false;
        }

        private static bool TryExpandViaChartData(
            object chart,
            PptHtmlChartGrid grid,
            out object excelApp,
            out string error,
            List<string> warnings = null)
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

                PourLog(warnings, "ChartData IsLinked=" + (TryPropString(chartData, "IsLinked") ?? "?"));
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

                PourLog(warnings, "打开内嵌簿 " + (TryPropString(workbook, "Name") ?? "?")
                    + " | " + DescribeSheetCells(ws, 5, 2));
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
                PourLog(warnings, "写入后 " + DescribeSheetCells(ws, rows + 1, cols)
                    + " range=" + TryRangeAddress(range));
                if (range != null)
                {
                    if (!TrySetSourceData(chart, range, warnings))
                    {
                        PourLog(warnings, "SetSourceData 未成功，继续（旧行为） | "
                            + DescribeLiveSeries(chart));
                    }
                }
                else
                {
                    PourLog(warnings, "无法定位写入区");
                }

                PourLog(warnings, "ChartData 结束 " + DescribeLiveSeries(chart));
                return true;
            }
            catch (Exception ex)
            {
                error = "灌入图表数据失败: " + ex.Message;
                PourLog(warnings, error);
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

        private static bool TrySetSourceData(object chart, object range, List<string> warnings = null)
        {
            try
            {
                WppCom.Invoke(chart, "SetSourceData", range, 2);
                PourLog(warnings, "SetSourceData(range,2) ok");
                return true;
            }
            catch (Exception ex)
            {
                PourLog(warnings, "SetSourceData(range,2) 失败: " + ex.Message);
            }

            try
            {
                WppCom.Invoke(chart, "SetSourceData", range);
                PourLog(warnings, "SetSourceData(range) ok");
                return true;
            }
            catch (Exception ex)
            {
                PourLog(warnings, "SetSourceData(range) 失败: " + ex.Message);
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

        private static int CountValueColumns(PptHtmlChartGrid grid)
        {
            int n = 0;
            if (grid == null || grid.Columns == null)
            {
                return 0;
            }

            for (int c = 0; c < grid.Columns.Count; c++)
            {
                if (grid.Columns[c] != null && grid.Columns[c].Role != "category")
                {
                    n++;
                }
            }

            return n;
        }

        private static bool TryVerifyPouredGrid(
            object chart,
            PptHtmlChartGrid want,
            out string error,
            List<string> warnings = null)
        {
            error = null;
            if (chart == null || want == null || !want.IsPourable)
            {
                error = "图表换数后无法校验内嵌表";
                return false;
            }

            PourLog(warnings, "校验 " + DescribeWantGrid(want) + " | " + DescribeLiveSeries(chart));
            if (!SeriesRowCountMatches(chart, want))
            {
                error = "图表系列点数与稿不一致（期望 "
                    + want.Rows.Count + " 行，图上 " + ReadSeriesRowCount(chart) + " 行）";
                PourLog(warnings, "校验失败点数 " + error);
                return false;
            }

            if (TryReadGridFromSeries(chart, out PptHtmlChartGrid live, out _)
                && live != null
                && !GridsMatch(want, live, compareNames: false))
            {
                error = "图表系列数值与稿不一致";
                PourLog(warnings, "校验失败数值 稿=" + DescribeWantGrid(want)
                    + " 系列=" + DescribeWantGrid(live));
                return false;
            }

            PourLog(warnings, "校验通过（系列点数与稿一致）");
            return true;
        }

        private static bool TryReadGridFromEmbeddedSheet(
            object chart,
            int cols,
            int rows,
            out PptHtmlChartGrid grid,
            out string error)
        {
            grid = null;
            error = null;
            object excelApp = null;
            try
            {
                object chartData = WppCom.GetProperty(chart, "ChartData");
                if (chartData == null)
                {
                    error = "无法访问 ChartData 以回读";
                    return false;
                }

                TryInvoke(chartData, "Activate");
                object workbook = WppCom.GetProperty(chartData, "Workbook");
                if (workbook == null)
                {
                    error = "无法打开图表内嵌工作簿以回读";
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

                var columns = new List<PptHtmlChartColumn>();
                for (int c = 0; c < cols; c++)
                {
                    string name = Convert.ToString(GetCell(ws, 1, c + 1) ?? "");
                    columns.Add(new PptHtmlChartColumn
                    {
                        Role = c == 0 ? "category" : "value",
                        Name = name
                    });
                }

                var dataRows = new List<List<string>>();
                for (int r = 0; r < rows; r++)
                {
                    var row = new List<string>();
                    for (int c = 0; c < cols; c++)
                    {
                        object raw = GetCell(ws, r + 2, c + 1);
                        row.Add(Convert.ToString(raw ?? "", CultureInfo.InvariantCulture));
                    }

                    dataRows.Add(row);
                }

                grid = new PptHtmlChartGrid
                {
                    Columns = columns,
                    Rows = dataRows
                };
                return grid.IsPourable;
            }
            catch (Exception ex)
            {
                error = "回读图表内嵌表失败: " + ex.Message;
                return false;
            }
            finally
            {
                HideEmbeddedExcel(excelApp);
                TryHideChartExcel(chart);
            }
        }

        private static bool GridAlreadyMatches(object chart, PptHtmlChartGrid grid)
        {
            if (chart == null || grid == null || !grid.IsPourable)
            {
                return false;
            }

            return TryReadGridFromSeries(chart, out PptHtmlChartGrid cur, out _)
                && GridsMatch(grid, cur, compareNames: true);
        }

        private static bool SeriesRowCountMatches(object chart, PptHtmlChartGrid grid)
        {
            return grid != null && grid.Rows != null && ReadSeriesRowCount(chart) == grid.Rows.Count;
        }

        private static int ReadSeriesRowCount(object chart)
        {
            try
            {
                object s1 = GetSeries(chart, 1);
                if (s1 == null)
                {
                    return 0;
                }

                List<string> cats = ToStringList(WppCom.GetProperty(s1, "XValues"));
                if (cats.Count > 0)
                {
                    return cats.Count;
                }

                return ToStringList(WppCom.GetProperty(s1, "Values")).Count;
            }
            catch (Exception)
            {
                return 0;
            }
        }

        private static bool GridsMatch(PptHtmlChartGrid want, PptHtmlChartGrid got, bool compareNames)
        {
            if (want == null || got == null
                || want.Columns == null || got.Columns == null
                || want.Rows == null || got.Rows == null
                || want.Columns.Count != got.Columns.Count
                || want.Rows.Count != got.Rows.Count)
            {
                return false;
            }

            for (int c = 0; c < want.Columns.Count; c++)
            {
                string role = want.Columns[c].Role ?? "";
                if (!string.Equals(got.Columns[c].Role ?? "", role, StringComparison.OrdinalIgnoreCase))
                {
                    return false;
                }

                if (compareNames
                    && role != "category"
                    && !string.Equals(got.Columns[c].Name ?? "", want.Columns[c].Name ?? "", StringComparison.Ordinal))
                {
                    return false;
                }
            }

            for (int r = 0; r < want.Rows.Count; r++)
            {
                List<string> a = want.Rows[r];
                List<string> b = got.Rows[r];
                if (a == null || b == null || a.Count < want.Columns.Count || b.Count < want.Columns.Count)
                {
                    return false;
                }

                for (int c = 0; c < want.Columns.Count; c++)
                {
                    if (want.Columns[c].Role == "value")
                    {
                        if (!TryParseLooseDouble(a[c], out double x)
                            || !TryParseLooseDouble(b[c], out double y)
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

        private static bool TryParseLooseDouble(string raw, out double n)
        {
            return double.TryParse(raw, NumberStyles.Float, CultureInfo.InvariantCulture, out n)
                || double.TryParse(raw, NumberStyles.Float, CultureInfo.CurrentCulture, out n);
        }

        private static int TryGetPointCount(object series)
        {
            object pts = TryGetPoints(series);
            if (pts == null)
            {
                return 0;
            }

            try
            {
                return Convert.ToInt32(WppCom.GetProperty(pts, "Count"));
            }
            catch (Exception)
            {
                return 0;
            }
        }

        /// <summary>
        /// 主题盘已在套回里先套。这里只盖更细的色：饼扇区或自动分色；柱/线系列填+点填。
        /// 饼图不套系列填。
        /// </summary>
        private static void TryInheritColors(object chart, ChartStyleSnap snap, List<string> warnings)
        {
            if (chart == null || snap == null)
            {
                return;
            }

            if (IsPieChart(chart))
            {
                object series = GetSeries(chart, 1);
                SeriesStyleSnap one = snap.Series != null && snap.Series.Count > 0
                    ? snap.Series[0]
                    : null;
                int live = TryGetPointCount(series);
                int oldPts = one == null || one.PointFills == null ? 0 : one.PointFills.Count;
                if (one != null && one.PointFills != null && live > 0 && oldPts == live)
                {
                    StyleLog(warnings, "饼图在主题上覆盖扇区色 " + live + " 个");
                    TryApplyPointFills(series, one.PointFills, warnings, "S1");
                    return;
                }

                StyleLog(warnings, "饼图已继承主题，扇区色 " + oldPts + " 个对 " + live + " 瓣，按旧色带扩/收");
                TrySetVaryByCategories(chart, true, warnings);
                List<FillSnap> expanded = ExpandPieSliceFills(one == null ? null : one.PointFills, live);
                if (expanded != null && expanded.Count == live)
                {
                    TryApplyPointFills(series, expanded, warnings, "S1");
                }

                return;
            }

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

                TryApplyFill(series, snap.Series[i].Fill, warnings, "S" + (i + 1));
                TryApplyPointFills(series, snap.Series[i].PointFills, warnings, "S" + (i + 1));
            }
        }

        private static void TryApplyChartTheme(object chart, ChartStyleSnap snap)
        {
            if (snap.ChartStyle.HasValue)
            {
                WppCom.TrySetProperty(chart, "ChartStyle", snap.ChartStyle.Value);
            }

            if (snap.ChartColor.HasValue)
            {
                WppCom.TrySetProperty(chart, "ChartColor", snap.ChartColor.Value);
            }
        }

        private static List<FillSnap> ExpandPieSliceFills(List<FillSnap> oldFills, int want)
        {
            if (want <= 0)
            {
                return null;
            }

            var seeds = new List<int>();
            if (oldFills != null)
            {
                for (int i = 0; i < oldFills.Count; i++)
                {
                    if (oldFills[i] != null && oldFills[i].SolidRgb.HasValue)
                    {
                        seeds.Add(oldFills[i].SolidRgb.Value);
                    }
                }
            }

            var fills = new List<FillSnap>();
            if (seeds.Count == 0)
            {
                return null;
            }

            int[] rgb = ExpandOfficeRgbRamp(seeds, want);
            for (int i = 0; i < rgb.Length; i++)
            {
                fills.Add(new FillSnap { Visible = true, SolidRgb = rgb[i] });
            }

            return fills;
        }

        /// <summary>
        /// 旧扇区色当种子：少了沿色带 HSL 插值拉长，多了均匀取样（含首尾）。
        /// 只有 1 个种子时绕色相、拉明度生成可区分的瓣。
        /// </summary>
        private static int[] ExpandOfficeRgbRamp(List<int> seeds, int want)
        {
            var dest = new int[want];
            if (seeds.Count == 1)
            {
                RgbToHsl(OfficeRgbToChannels(seeds[0]), out double h, out double s, out double l);
                for (int i = 0; i < want; i++)
                {
                    double t = want == 1 ? 0 : i / (double)(want - 1);
                    double nh = (h + (t - 0.5) * 0.22 + 1.0) % 1.0;
                    double nl = Clamp01(l + (t - 0.5) * 0.28);
                    dest[i] = ChannelsToOfficeRgb(HslToRgb(nh, Math.Max(0.25, s), nl));
                }

                return dest;
            }

            if (seeds.Count >= want)
            {
                for (int i = 0; i < want; i++)
                {
                    double t = want == 1 ? 0 : i / (double)(want - 1);
                    dest[i] = seeds[(int)Math.Round(t * (seeds.Count - 1))];
                }

                return dest;
            }

            int segments = seeds.Count - 1;
            for (int i = 0; i < want; i++)
            {
                double t = want == 1 ? 0 : i / (double)(want - 1);
                double pos = t * segments;
                int a = (int)Math.Floor(pos);
                if (a >= segments)
                {
                    dest[i] = seeds[seeds.Count - 1];
                    continue;
                }

                double local = pos - a;
                dest[i] = LerpOfficeRgbHsl(seeds[a], seeds[a + 1], local);
            }

            return dest;
        }

        private static int LerpOfficeRgbHsl(int from, int to, double t)
        {
            RgbToHsl(OfficeRgbToChannels(from), out double h1, out double s1, out double l1);
            RgbToHsl(OfficeRgbToChannels(to), out double h2, out double s2, out double l2);
            double dh = h2 - h1;
            if (dh > 0.5)
            {
                dh -= 1;
            }
            else if (dh < -0.5)
            {
                dh += 1;
            }

            double h = (h1 + dh * t + 1.0) % 1.0;
            return ChannelsToOfficeRgb(HslToRgb(h, s1 + (s2 - s1) * t, l1 + (l2 - l1) * t));
        }

        private static int[] OfficeRgbToChannels(int office)
        {
            return new[] { office & 0xFF, (office >> 8) & 0xFF, (office >> 16) & 0xFF };
        }

        private static int ChannelsToOfficeRgb(int[] rgb)
        {
            return (rgb[0] & 0xFF) | ((rgb[1] & 0xFF) << 8) | ((rgb[2] & 0xFF) << 16);
        }

        private static void RgbToHsl(int[] rgb, out double h, out double s, out double l)
        {
            double r = rgb[0] / 255.0;
            double g = rgb[1] / 255.0;
            double b = rgb[2] / 255.0;
            double max = Math.Max(r, Math.Max(g, b));
            double min = Math.Min(r, Math.Min(g, b));
            l = (max + min) / 2.0;
            if (Math.Abs(max - min) < 0.0001)
            {
                h = 0;
                s = 0;
                return;
            }

            double d = max - min;
            s = l > 0.5 ? d / (2.0 - max - min) : d / (max + min);
            if (max == r)
            {
                h = ((g - b) / d + (g < b ? 6 : 0)) / 6.0;
            }
            else if (max == g)
            {
                h = ((b - r) / d + 2) / 6.0;
            }
            else
            {
                h = ((r - g) / d + 4) / 6.0;
            }
        }

        private static int[] HslToRgb(double h, double s, double l)
        {
            s = Clamp01(s);
            l = Clamp01(l);
            if (s <= 0)
            {
                int g = (int)Math.Round(l * 255);
                return new[] { g, g, g };
            }

            double q = l < 0.5 ? l * (1 + s) : l + s - l * s;
            double p = 2 * l - q;
            return new[]
            {
                (int)Math.Round(HueToRgb(p, q, h + 1.0 / 3.0) * 255),
                (int)Math.Round(HueToRgb(p, q, h) * 255),
                (int)Math.Round(HueToRgb(p, q, h - 1.0 / 3.0) * 255)
            };
        }

        private static double HueToRgb(double p, double q, double t)
        {
            if (t < 0)
            {
                t += 1;
            }

            if (t > 1)
            {
                t -= 1;
            }

            if (t < 1.0 / 6.0)
            {
                return p + (q - p) * 6 * t;
            }

            if (t < 0.5)
            {
                return q;
            }

            if (t < 2.0 / 3.0)
            {
                return p + (q - p) * (2.0 / 3.0 - t) * 6;
            }

            return p;
        }

        private static double Clamp01(double x)
        {
            if (x < 0)
            {
                return 0;
            }

            return x > 1 ? 1 : x;
        }

        private static void TrySetVaryByCategories(object chart, bool on, List<string> warnings = null)
        {
            try
            {
                object g = TryInvoke(chart, "ChartGroups", 1);
                WppCom.TrySetProperty(g, "VaryByCategories", on);
                object live = g == null ? null : WppCom.GetProperty(g, "VaryByCategories");
                StyleLog(warnings, "VaryByCategories=" + on + " live=" + live);
            }
            catch (Exception ex)
            {
                StyleLog(warnings, "VaryByCategories 异常: " + ex.Message);
            }
        }

        /// <summary>
        /// 饼图已是饼则不改 2D↔3D。未写类型时建图已跟旧图；这里只兜底「建成了非饼、快照却是饼」。
        /// </summary>
        private static void TryRestorePieChartType(object chart, ChartStyleSnap snap)
        {
            if (chart == null || snap == null || snap.Series == null)
            {
                return;
            }

            if (IsPieChart(chart))
            {
                return;
            }

            int? want = null;
            for (int i = 0; i < snap.Series.Count; i++)
            {
                if (snap.Series[i].ChartType.HasValue && IsPieXl(snap.Series[i].ChartType.Value))
                {
                    want = snap.Series[i].ChartType.Value;
                    break;
                }
            }

            if (!want.HasValue)
            {
                return;
            }

            try
            {
                WppCom.TrySetProperty(chart, "ChartType", want.Value);
            }
            catch (Exception)
            {
            }
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

        private static string ParseColorOrNone(string raw)
        {
            if (string.Equals(raw, "none", StringComparison.OrdinalIgnoreCase)
                || string.Equals(raw, "transparent", StringComparison.OrdinalIgnoreCase))
            {
                return "none";
            }

            return NormalizeHexOrNull(raw);
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
