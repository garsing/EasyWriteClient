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

        /// <summary>对齐 COM DataLabels.ShowValue。</summary>
        public string ShowValue { get; set; }

        /// <summary>对齐 COM DataLabels.ShowPercentage。</summary>
        public string ShowPercentage { get; set; }

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

        /// <summary>
        /// 解析层钉死：该列是否写过与「数据标签开关」相关的支持属性。
        /// </summary>
        public bool DataLabelsMentioned { get; set; }
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

        /// <summary>图级：对齐 COM ShowValue；会落到各系列（列上未写时）。</summary>
        public string ShowValue { get; set; }

        /// <summary>图级：对齐 COM ShowPercentage。</summary>
        public string ShowPercentage { get; set; }

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

        /// <summary>
        /// 解析层钉死：稿是否写过与「横轴显隐」相关的支持属性。
        /// 合并时只读此标志，禁止从 htmlSnap 反推（易被其它开关误伤）。
        /// </summary>
        public bool AxisXVisibilityMentioned { get; set; }

        /// <summary>
        /// 解析层钉死：稿是否写过与「纵轴显隐」相关的支持属性。
        /// </summary>
        public bool AxisYVisibilityMentioned { get; set; }

        /// <summary>解析层：是否写过与图例开关相关的支持属性。</summary>
        public bool LegendMentioned { get; set; }

        /// <summary>解析层：是否写过与图标题开关相关的支持属性。</summary>
        public bool TitleMentioned { get; set; }

        /// <summary>解析层：图级是否写过与数据标签开关相关的支持属性。</summary>
        public bool DataLabelsMentioned { get; set; }

        /// <summary>解析层：是否写过与网格开闭相关的支持属性。</summary>
        public bool GridlinesMentioned { get; set; }
    }

    internal sealed class PptHtmlChartReadModel
    {
        public PptHtmlChartFormat Format { get; set; }

        public PptHtmlChartGrid Grid { get; set; }
    }
}
