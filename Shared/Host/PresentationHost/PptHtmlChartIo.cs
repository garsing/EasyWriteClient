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
        public const int MaxRows = 30;

        public const int MaxCols = 8;

        private const int XlColumnClustered = 51;
        private const int XlBarClustered = 57;
        private const int XlLine = 4;
        private const int XlLineMarkers = 65;
        private const int XlDataLabelsShowValue = 2;
        private const int XlPie = 5;
        private const int Xl3DPie = -4102;

        /// <summary>xlCombination：混合图整图结果态，不可传给 AddChart2。</summary>
        private const int XlCombination = -4111;

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

            /// <summary>图标题正文。换数要拍旧图并写回新图；仅开关/字体会留下 AddChart2 默认系列名。</summary>
            public string Title { get; set; }

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

            public bool? ShowValue { get; set; }

            public bool? ShowPercentage { get; set; }

            public int? DataLabelPosition { get; set; }

            public string DataLabelFontName { get; set; }

            public double? DataLabelFontSize { get; set; }

            public int? DataLabelFontColor { get; set; }

            public string DataLabelNumberFormat { get; set; }

            /// <summary>解析层：该系列列是否写过与数据标签开关相关的支持属性。</summary>
            public bool DataLabelsMentioned { get; set; }

            /// <summary>解析层：列上显式 data-show-data-labels=false。</summary>
            public bool DataLabelsExplicitOff { get; set; }

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

    }
}
