using WordAddIn1.PresentationHost;

namespace PptChartRoundtripTest
{
    internal static class ChartAssert
    {
        public static void FormatFull(TestRun run, string prefix, PptHtmlChartFormat fmt, string chartType)
        {
            if (fmt == null)
            {
                run.Fail(prefix + " format", "ChartFormat 为空");
                return;
            }

            run.ExpectEqual(prefix + " data-chart-type", chartType, fmt.ChartType);
            run.ExpectEqual(prefix + " data-title", "市场规模", fmt.Title);
            run.ExpectEqual(prefix + " data-title-font-size", "14", fmt.TitleFontSize);
            run.ExpectEqual(prefix + " data-title-font-bold", "true", fmt.TitleFontBold);
            run.ExpectEqual(prefix + " data-title-font-color", "#FFFFFF", fmt.TitleFontColor);
            run.ExpectEqual(prefix + " data-theme", "office", fmt.Theme);
            run.ExpectEqual(prefix + " data-chart-style", "286", fmt.ChartStyle);
            run.ExpectEqual(prefix + " data-legend", "bottom", fmt.Legend);
            run.ExpectEqual(prefix + " data-legend-font-color", "#EEEEEE", fmt.LegendFontColor);
            run.ExpectEqual(prefix + " data-show-data-labels", "true", fmt.ShowDataLabels);
            run.ExpectEqual(prefix + " data-show-value", "true", fmt.ShowValue);
            run.ExpectEqual(prefix + " data-show-percentage", "false", fmt.ShowPercentage);
            run.ExpectEqual(prefix + " data-gridlines", "false", fmt.Gridlines);
            run.ExpectEqual(prefix + " data-gap-width", "110", fmt.GapWidth);
            run.ExpectEqual(prefix + " data-overlap", "-20", fmt.Overlap);
            run.ExpectEqual(prefix + " data-plot-color", "none", fmt.PlotColor);
            run.ExpectEqual(prefix + " data-chart-area-color", "none", fmt.ChartAreaColor);
            run.ExpectEqual(prefix + " data-plot-box", "8.6,15.29,349.1,153.31", fmt.PlotBox);
            run.ExpectEqual(prefix + " data-plot-inside", "21.68,19.79,336.02,133.98", fmt.PlotInside);
            run.ExpectEqual(prefix + " data-data-markers", "false", fmt.DataMarkers);
            run.ExpectEqual(prefix + " data-marker-size", "5", fmt.MarkerSize);
            run.ExpectEqual(prefix + " data-chart-line-weight", "1.5", fmt.ChartLineWeight);
            run.ExpectEqual(prefix + " data-axis-y-min", "0", fmt.AxisYMin);
            run.ExpectEqual(prefix + " data-axis-y-max", "10", fmt.AxisYMax);
            run.ExpectEqual(prefix + " data-axis-y-major-unit", "2", fmt.AxisYMajorUnit);
            run.ExpectEqual(prefix + " data-explosion", "8", fmt.Explosion);
            run.ExpectEqual(prefix + " data-fill-missing", "gap", fmt.FillMissing);
            run.ExpectEqual(prefix + " data-axis-x 标题", "年份", fmt.AxisX);
            run.ExpectEqual(prefix + " data-axis-x-type", "category", fmt.AxisXType);
            run.ExpectEqual(prefix + " data-axis-x-format", "General", fmt.AxisXFormat);
            run.ExpectEqual(prefix + " data-axis-x-tick-count", "4", fmt.AxisXTickCount);
            run.ExpectEqual(prefix + " data-axis-x-tick-spacing", "1", fmt.AxisXTickSpacing);
            run.ExpectEqual(prefix + " data-axis-x-between", "true", fmt.AxisXBetween);
            run.ExpectEqual(prefix + " data-axis-y 标题", "规模", fmt.AxisY);
            run.ExpectEqual(prefix + " data-axis-y-secondary", "增速", fmt.AxisYSecondary);
            run.Expect(prefix + " LegendMentioned", fmt.LegendMentioned, "应钉死图例提及");
            run.Expect(prefix + " TitleMentioned", fmt.TitleMentioned, "应钉死标题提及");
            run.Expect(prefix + " DataLabelsMentioned", fmt.DataLabelsMentioned, "应钉死标签提及");
            run.Expect(prefix + " GridlinesMentioned", fmt.GridlinesMentioned, "应钉死网格提及");
            run.Expect(prefix + " AxisXVisibilityMentioned", fmt.AxisXVisibilityMentioned, "应钉死横轴提及");
            run.Expect(prefix + " AxisYVisibilityMentioned", fmt.AxisYVisibilityMentioned, "应钉死纵轴提及");
        }

        public static void AxisFull(TestRun run, string prefix, PptHtmlAxisExtras ax, string format)
        {
            if (ax == null)
            {
                run.Fail(prefix, "轴 extras 为空");
                return;
            }

            run.ExpectEqual(prefix + "-visible", "true", ax.Visible);
            run.ExpectEqual(prefix + "-tick-font", "微软雅黑", ax.TickFont);
            run.ExpectEqual(prefix + "-tick-color", "#FFFFFF", ax.TickColor);
            run.ExpectEqual(prefix + "-tick-size", "12", ax.TickSize);
            run.ExpectEqual(prefix + "-tick-position", "next", ax.TickPosition);
            run.ExpectEqual(prefix + "-major-tick", "outside", ax.MajorTick);
            run.ExpectEqual(prefix + "-minor-tick", "none", ax.MinorTick);
            if (format != null)
            {
                run.ExpectEqual(prefix + "-format", format, ax.Format);
            }

            run.ExpectEqual(prefix + "-grid", "false", ax.Grid);
            // 关网格时 COM/读回不拍线色；稿也不应依赖 GridColor
            run.ExpectEqual(prefix + "-line", "#000000", ax.Line);
            run.ExpectEqual(prefix + "-line-weight", "0.75", ax.LineWeight);
        }

        public static void AxisNamed(
            TestRun run,
            string prefix,
            PptHtmlAxisExtras ax,
            string visible,
            string tickColor,
            string majorTick,
            string format,
            string line)
        {
            if (ax == null)
            {
                run.Fail(prefix, "轴 extras 为空");
                return;
            }

            run.ExpectEqual(prefix + "-visible", visible, ax.Visible);
            run.ExpectEqual(prefix + "-tick-color", tickColor, ax.TickColor);
            run.ExpectEqual(prefix + "-major-tick", majorTick, ax.MajorTick);
            if (format != null)
            {
                run.ExpectEqual(prefix + "-format", format, ax.Format);
            }

            run.ExpectEqual(prefix + "-line", line, ax.Line);
        }

        public static void SeriesFull(TestRun run, string prefix, PptHtmlChartColumn col, string seriesType, string axis)
        {
            if (col == null)
            {
                run.Fail(prefix, "系列列为空");
                return;
            }

            run.ExpectEqual(prefix + " data-series-type", seriesType, col.SeriesType);
            run.ExpectEqual(prefix + " data-axis", axis, col.AxisY);
            run.ExpectEqual(prefix + " data-color", "#2497EA", col.Color);
            run.ExpectEqual(prefix + " data-show-data-labels", "true", col.ShowDataLabels);
            run.ExpectEqual(prefix + " data-show-value", "true", col.ShowValue);
            run.ExpectEqual(prefix + " data-show-percentage", "false", col.ShowPercentage);
            run.ExpectEqual(prefix + " data-fill-gradient", "0:#FFFFFF@1;1:#2497EA@0", col.FillGradient);
            run.ExpectEqual(prefix + " data-fill-angle", "270", col.FillAngle);
            run.ExpectEqual(prefix + " data-line", "none", col.Line);
            run.ExpectEqual(prefix + " data-line-weight", "2", col.LineWeight);
            run.ExpectEqual(prefix + " data-marker", "none", col.Marker);
            run.ExpectEqual(prefix + " data-marker-size", "5", col.MarkerSize);
            run.ExpectEqual(prefix + " data-marker-color", "#FFFFFF", col.MarkerColor);
            run.ExpectEqual(prefix + " data-marker-fill", "#FFFFFF", col.MarkerFill);
            run.ExpectEqual(prefix + " data-label-position", "outside", col.LabelPosition);
            run.ExpectEqual(prefix + " data-label-font", "微软雅黑", col.LabelFont);
            run.ExpectEqual(prefix + " data-label-size", "10", col.LabelSize);
            run.ExpectEqual(prefix + " data-label-color", "#FFFFFF", col.LabelColor);
            run.ExpectEqual(prefix + " data-label-format", "0.00", col.LabelFormat);
            run.Expect(prefix + " DataLabelsMentioned", col.DataLabelsMentioned, "列标签应被提及");
        }

        public static void LineSeriesFull(TestRun run, string prefix, PptHtmlChartColumn col)
        {
            if (col == null)
            {
                run.Fail(prefix, "折线列为空");
                return;
            }

            run.ExpectEqual(prefix + " data-series-type", "line", col.SeriesType);
            run.ExpectEqual(prefix + " data-axis", "y2", col.AxisY);
            run.ExpectEqual(prefix + " data-line", "#FFFFFF", col.Line);
            run.ExpectEqual(prefix + " data-line-weight", "2", col.LineWeight);
            run.ExpectEqual(prefix + " data-marker", "circle", col.Marker);
            run.ExpectEqual(prefix + " data-marker-size", "6", col.MarkerSize);
            run.ExpectEqual(prefix + " data-marker-color", "#FFFFFF", col.MarkerColor);
            run.ExpectEqual(prefix + " data-marker-fill", "#FFFFFF", col.MarkerFill);
            run.ExpectEqual(prefix + " data-label-position", "above", col.LabelPosition);
            run.ExpectEqual(prefix + " data-label-format", "0.0%", col.LabelFormat);
            run.ExpectEqual(prefix + " data-show-percentage", "false", col.ShowPercentage);
        }

        public static bool TryParse(TestRun run, string name, string html, out PptHtmlApplyNode node)
        {
            if (ChartHtml.TryParseFirstChart(html, out node, out string error))
            {
                return true;
            }

            run.Fail(name, error);
            node = null;
            return false;
        }
    }
}
