using System.Collections.Generic;

namespace PptChartRoundtripTest
{
    /// <summary>第二波属性交叉：50 组 × 新建/换数 = 100 条（前缀 mix-）。</summary>
    internal static partial class ChartAttrCatalog
    {
        private static void AddMixes(List<ChartCase> list, string p, bool replace)
        {
            list.Add(Typed(p + "mix-title-legend-top", replace, "column", "top",
                Merge(Node("data-title", "图例top"), Node("data-title-font-color", "#C00000"),
                    Node("data-show-data-labels", "true"), Node("data-show-value", "true")),
                Th(Pair("data-color", "#5B9BD5"), Pair("data-show-data-labels", "true"),
                    Pair("data-show-value", "true"), Pair("data-label-position", "outside")),
                AttrCheck.Exact("Title", "图例top"),
                AttrCheck.Exact("Legend", "top"),
                AttrCheck.Hex("TitleFontColor", "#C00000"),
                AttrCheck.Hex("S0.Color", "#5B9BD5"),
                AttrCheck.Exact("S0.ShowDataLabels", "true")));

            list.Add(Typed(p + "mix-title-legend-left", replace, "column", "left",
                Merge(Node("data-title", "图例left"), Node("data-title-font-color", "#C00000"),
                    Node("data-show-data-labels", "true"), Node("data-show-value", "true")),
                Th(Pair("data-color", "#5B9BD5"), Pair("data-show-data-labels", "true"),
                    Pair("data-show-value", "true"), Pair("data-label-position", "outside")),
                AttrCheck.Exact("Title", "图例left"),
                AttrCheck.Exact("Legend", "left"),
                AttrCheck.Hex("TitleFontColor", "#C00000"),
                AttrCheck.Hex("S0.Color", "#5B9BD5"),
                AttrCheck.Exact("S0.ShowDataLabels", "true")));

            list.Add(Typed(p + "mix-title-legend-right", replace, "column", "right",
                Merge(Node("data-title", "图例right"), Node("data-title-font-color", "#C00000"),
                    Node("data-show-data-labels", "true"), Node("data-show-value", "true")),
                Th(Pair("data-color", "#5B9BD5"), Pair("data-show-data-labels", "true"),
                    Pair("data-show-value", "true"), Pair("data-label-position", "outside")),
                AttrCheck.Exact("Title", "图例right"),
                AttrCheck.Exact("Legend", "right"),
                AttrCheck.Hex("TitleFontColor", "#C00000"),
                AttrCheck.Hex("S0.Color", "#5B9BD5"),
                AttrCheck.Exact("S0.ShowDataLabels", "true")));

            list.Add(Typed(p + "mix-title-legend-bot", replace, "column", "bottom",
                Merge(Node("data-title", "图例bot"), Node("data-title-font-color", "#C00000"),
                    Node("data-show-data-labels", "true"), Node("data-show-value", "true")),
                Th(Pair("data-color", "#5B9BD5"), Pair("data-show-data-labels", "true"),
                    Pair("data-show-value", "true"), Pair("data-label-position", "outside")),
                AttrCheck.Exact("Title", "图例bot"),
                AttrCheck.Exact("Legend", "bottom"),
                AttrCheck.Hex("TitleFontColor", "#C00000"),
                AttrCheck.Hex("S0.Color", "#5B9BD5"),
                AttrCheck.Exact("S0.ShowDataLabels", "true")));

            list.Add(Typed(p + "mix-title-legend-none", replace, "column", "none",
                Merge(Node("data-title", "图例none"), Node("data-title-font-color", "#C00000"),
                    Node("data-show-data-labels", "true"), Node("data-show-value", "true")),
                Th(Pair("data-color", "#5B9BD5"), Pair("data-show-data-labels", "true"),
                    Pair("data-show-value", "true"), Pair("data-label-position", "outside")),
                AttrCheck.Exact("Title", "图例none"),
                AttrCheck.Exact("Legend", "none"),
                AttrCheck.Hex("TitleFontColor", "#C00000"),
                AttrCheck.Hex("S0.Color", "#5B9BD5"),
                AttrCheck.Exact("S0.ShowDataLabels", "true")));

            list.Add(Column(p + "mix-gap-ov-1", replace,
                Merge(Node("data-gap-width", "50"), Node("data-overlap", "-40")),
                Th(Pair("data-color", "#ED7D31")),
                AttrCheck.Num("GapWidth", "50"),
                AttrCheck.Num("Overlap", "-40"),
                AttrCheck.Hex("S0.Color", "#ED7D31")));

            list.Add(Column(p + "mix-gap-ov-2", replace,
                Merge(Node("data-gap-width", "70"), Node("data-overlap", "-20")),
                Th(Pair("data-color", "#70AD47")),
                AttrCheck.Num("GapWidth", "70"),
                AttrCheck.Num("Overlap", "-20"),
                AttrCheck.Hex("S0.Color", "#70AD47")));

            list.Add(Column(p + "mix-gap-ov-3", replace,
                Merge(Node("data-gap-width", "100"), Node("data-overlap", "0")),
                Th(Pair("data-color", "#FFC000")),
                AttrCheck.Num("GapWidth", "100"),
                AttrCheck.Num("Overlap", "0"),
                AttrCheck.Hex("S0.Color", "#FFC000")));

            list.Add(Column(p + "mix-gap-ov-4", replace,
                Merge(Node("data-gap-width", "120"), Node("data-overlap", "-10")),
                Th(Pair("data-color", "#4472C4")),
                AttrCheck.Num("GapWidth", "120"),
                AttrCheck.Num("Overlap", "-10"),
                AttrCheck.Hex("S0.Color", "#4472C4")));

            list.Add(Column(p + "mix-gap-ov-5", replace,
                Merge(Node("data-gap-width", "150"), Node("data-overlap", "-35")),
                Th(Pair("data-color", "#7030A0")),
                AttrCheck.Num("GapWidth", "150"),
                AttrCheck.Num("Overlap", "-35"),
                AttrCheck.Hex("S0.Color", "#7030A0")));

            list.Add(Column(p + "mix-scale-1", replace,
                Merge(Node("data-axis-y-min", "0"), Node("data-axis-y-max", "8"),
                    Node("data-axis-y-major-unit", "2"), Node("data-axis-y", "刻度1"),
                    Node("data-gridlines", "true")),
                AttrCheck.Num("AxisYMin", "0"),
                AttrCheck.Num("AxisYMax", "8"),
                AttrCheck.Num("AxisYMajorUnit", "2"),
                AttrCheck.Exact("AxisY", "刻度1"),
                AttrCheck.Exact("Gridlines", "true")));

            list.Add(Column(p + "mix-scale-2", replace,
                Merge(Node("data-axis-y-min", "0"), Node("data-axis-y-max", "20"),
                    Node("data-axis-y-major-unit", "4"), Node("data-axis-y", "刻度2"),
                    Node("data-gridlines", "true")),
                AttrCheck.Num("AxisYMin", "0"),
                AttrCheck.Num("AxisYMax", "20"),
                AttrCheck.Num("AxisYMajorUnit", "4"),
                AttrCheck.Exact("AxisY", "刻度2"),
                AttrCheck.Exact("Gridlines", "true")));

            list.Add(Column(p + "mix-scale-3", replace,
                Merge(Node("data-axis-y-min", "1"), Node("data-axis-y-max", "11"),
                    Node("data-axis-y-major-unit", "2"), Node("data-axis-y", "刻度3"),
                    Node("data-gridlines", "true")),
                AttrCheck.Num("AxisYMin", "1"),
                AttrCheck.Num("AxisYMax", "11"),
                AttrCheck.Num("AxisYMajorUnit", "2"),
                AttrCheck.Exact("AxisY", "刻度3"),
                AttrCheck.Exact("Gridlines", "true")));

            list.Add(Column(p + "mix-scale-4", replace,
                Merge(Node("data-axis-y-min", "0"), Node("data-axis-y-max", "100"),
                    Node("data-axis-y-major-unit", "25"), Node("data-axis-y", "刻度4"),
                    Node("data-gridlines", "true")),
                AttrCheck.Num("AxisYMin", "0"),
                AttrCheck.Num("AxisYMax", "100"),
                AttrCheck.Num("AxisYMajorUnit", "25"),
                AttrCheck.Exact("AxisY", "刻度4"),
                AttrCheck.Exact("Gridlines", "true")));

            list.Add(Column(p + "mix-scale-5", replace,
                Merge(Node("data-axis-y-min", "-2"), Node("data-axis-y-max", "10"),
                    Node("data-axis-y-major-unit", "2"), Node("data-axis-y", "刻度5"),
                    Node("data-gridlines", "true")),
                AttrCheck.Num("AxisYMin", "-2"),
                AttrCheck.Num("AxisYMax", "10"),
                AttrCheck.Num("AxisYMajorUnit", "2"),
                AttrCheck.Exact("AxisY", "刻度5"),
                AttrCheck.Exact("Gridlines", "true")));

            list.Add(Typed(p + "mix-line-skin-1", replace, "line", "bottom",
                Merge(Node("data-title", "折线皮1"), Node("data-data-markers", "true")),
                Th(Pair("data-line", "#00B0F0"), Pair("data-line-weight", "1.5"),
                    Pair("data-marker", "circle"), Pair("data-marker-size", "7"),
                    Pair("data-marker-fill", "#FFFFFF")),
                AttrCheck.Exact("Title", "折线皮1"),
                AttrCheck.Exact("Legend", "bottom"),
                AttrCheck.ApplyOnly("DataMarkers", "true"),
                AttrCheck.Hex("S0.Line", "#00B0F0"),
                AttrCheck.Num("S0.LineWeight", "1.5"),
                AttrCheck.Exact("S0.Marker", "circle")));

            list.Add(Typed(p + "mix-line-skin-2", replace, "line", "bottom",
                Merge(Node("data-title", "折线皮2"), Node("data-data-markers", "true")),
                Th(Pair("data-line", "#ED7D31"), Pair("data-line-weight", "2"),
                    Pair("data-marker", "square"), Pair("data-marker-size", "7"),
                    Pair("data-marker-fill", "#FFFFFF")),
                AttrCheck.Exact("Title", "折线皮2"),
                AttrCheck.Exact("Legend", "bottom"),
                AttrCheck.ApplyOnly("DataMarkers", "true"),
                AttrCheck.Hex("S0.Line", "#ED7D31"),
                AttrCheck.Num("S0.LineWeight", "2"),
                AttrCheck.Exact("S0.Marker", "square")));

            list.Add(Typed(p + "mix-line-skin-3", replace, "line", "bottom",
                Merge(Node("data-title", "折线皮3"), Node("data-data-markers", "true")),
                Th(Pair("data-line", "#70AD47"), Pair("data-line-weight", "2.5"),
                    Pair("data-marker", "triangle"), Pair("data-marker-size", "7"),
                    Pair("data-marker-fill", "#FFFFFF")),
                AttrCheck.Exact("Title", "折线皮3"),
                AttrCheck.Exact("Legend", "bottom"),
                AttrCheck.ApplyOnly("DataMarkers", "true"),
                AttrCheck.Hex("S0.Line", "#70AD47"),
                AttrCheck.Num("S0.LineWeight", "2.5"),
                AttrCheck.Exact("S0.Marker", "triangle")));

            list.Add(Typed(p + "mix-line-skin-4", replace, "line", "bottom",
                Merge(Node("data-title", "折线皮4"), Node("data-data-markers", "true")),
                Th(Pair("data-line", "#7030A0"), Pair("data-line-weight", "1.75"),
                    Pair("data-marker", "diamond"), Pair("data-marker-size", "7"),
                    Pair("data-marker-fill", "#FFFFFF")),
                AttrCheck.Exact("Title", "折线皮4"),
                AttrCheck.Exact("Legend", "bottom"),
                AttrCheck.ApplyOnly("DataMarkers", "true"),
                AttrCheck.Hex("S0.Line", "#7030A0"),
                AttrCheck.Num("S0.LineWeight", "1.75"),
                AttrCheck.Exact("S0.Marker", "diamond")));

            list.Add(Typed(p + "mix-line-skin-5", replace, "line", "bottom",
                Merge(Node("data-title", "折线皮5"), Node("data-data-markers", "true")),
                Th(Pair("data-line", "#C00000"), Pair("data-line-weight", "2"),
                    Pair("data-marker", "star"), Pair("data-marker-size", "7"),
                    Pair("data-marker-fill", "#FFFFFF")),
                AttrCheck.Exact("Title", "折线皮5"),
                AttrCheck.Exact("Legend", "bottom"),
                AttrCheck.ApplyOnly("DataMarkers", "true"),
                AttrCheck.Hex("S0.Line", "#C00000"),
                AttrCheck.Num("S0.LineWeight", "2"),
                AttrCheck.Exact("S0.Marker", "star")));

            list.Add(Pie(p + "mix-pie-1", replace, "right",
                Merge(Node("data-title", "饼交叉1"), Node("data-explosion", "5"),
                    Node("data-show-data-labels", "true"), Node("data-show-value", "false"),
                    Node("data-show-percentage", "true")),
                Th(Pair("data-show-data-labels", "true"), Pair("data-show-value", "false"),
                    Pair("data-show-percentage", "true"), Pair("data-label-position", "outside")),
                AttrCheck.Exact("Title", "饼交叉1"),
                AttrCheck.Exact("Explosion", "5"),
                AttrCheck.Exact("S0.ShowPercentage", "true"),
                AttrCheck.Exact("S0.LabelPosition", "outside")));

            list.Add(Pie(p + "mix-pie-2", replace, "right",
                Merge(Node("data-title", "饼交叉2"), Node("data-explosion", "8"),
                    Node("data-show-data-labels", "true"), Node("data-show-value", "false"),
                    Node("data-show-percentage", "true")),
                Th(Pair("data-show-data-labels", "true"), Pair("data-show-value", "false"),
                    Pair("data-show-percentage", "true"), Pair("data-label-position", "outside")),
                AttrCheck.Exact("Title", "饼交叉2"),
                AttrCheck.Exact("Explosion", "8"),
                AttrCheck.Exact("S0.ShowPercentage", "true"),
                AttrCheck.Exact("S0.LabelPosition", "outside")));

            list.Add(Pie(p + "mix-pie-3", replace, "right",
                Merge(Node("data-title", "饼交叉3"), Node("data-explosion", "12"),
                    Node("data-show-data-labels", "true"), Node("data-show-value", "false"),
                    Node("data-show-percentage", "true")),
                Th(Pair("data-show-data-labels", "true"), Pair("data-show-value", "false"),
                    Pair("data-show-percentage", "true"), Pair("data-label-position", "outside")),
                AttrCheck.Exact("Title", "饼交叉3"),
                AttrCheck.Exact("Explosion", "12"),
                AttrCheck.Exact("S0.ShowPercentage", "true"),
                AttrCheck.Exact("S0.LabelPosition", "outside")));

            list.Add(Pie(p + "mix-pie-4", replace, "right",
                Merge(Node("data-title", "饼交叉4"), Node("data-explosion", "0"),
                    Node("data-show-data-labels", "true"), Node("data-show-value", "false"),
                    Node("data-show-percentage", "true")),
                Th(Pair("data-show-data-labels", "true"), Pair("data-show-value", "false"),
                    Pair("data-show-percentage", "true"), Pair("data-label-position", "outside")),
                AttrCheck.Exact("Title", "饼交叉4"),
                AttrCheck.Exact("Explosion", "0"),
                AttrCheck.Exact("S0.ShowPercentage", "true"),
                AttrCheck.Exact("S0.LabelPosition", "outside")));

            list.Add(Pie(p + "mix-pie-5", replace, "right",
                Merge(Node("data-title", "饼交叉5"), Node("data-explosion", "15"),
                    Node("data-show-data-labels", "true"), Node("data-show-value", "false"),
                    Node("data-show-percentage", "true")),
                Th(Pair("data-show-data-labels", "true"), Pair("data-show-value", "false"),
                    Pair("data-show-percentage", "true"), Pair("data-label-position", "outside")),
                AttrCheck.Exact("Title", "饼交叉5"),
                AttrCheck.Exact("Explosion", "15"),
                AttrCheck.Exact("S0.ShowPercentage", "true"),
                AttrCheck.Exact("S0.LabelPosition", "outside")));

            list.Add(Typed(p + "mix-plot-area-1", replace, "column", "top",
                Merge(Node("data-title", "区色1"), Node("data-plot-color", "#FFF2CC"),
                    Node("data-chart-area-color", "#F2F2F2")),
                Th(Pair("data-color", "#5B9BD5")),
                AttrCheck.Exact("Title", "区色1"),
                AttrCheck.Exact("Legend", "top"),
                AttrCheck.Hex("PlotColor", "#FFF2CC"),
                AttrCheck.Hex("ChartAreaColor", "#F2F2F2"),
                AttrCheck.Hex("S0.Color", "#5B9BD5")));

            list.Add(Typed(p + "mix-plot-area-2", replace, "column", "top",
                Merge(Node("data-title", "区色2"), Node("data-plot-color", "#E2EFDA"),
                    Node("data-chart-area-color", "#DDEBF7")),
                Th(Pair("data-color", "#70AD47")),
                AttrCheck.Exact("Title", "区色2"),
                AttrCheck.Exact("Legend", "top"),
                AttrCheck.Hex("PlotColor", "#E2EFDA"),
                AttrCheck.Hex("ChartAreaColor", "#DDEBF7"),
                AttrCheck.Hex("S0.Color", "#70AD47")));

            list.Add(Typed(p + "mix-plot-area-3", replace, "column", "top",
                Merge(Node("data-title", "区色3"), Node("data-plot-color", "#FCE4D6"),
                    Node("data-chart-area-color", "#FFF2CC")),
                Th(Pair("data-color", "#ED7D31")),
                AttrCheck.Exact("Title", "区色3"),
                AttrCheck.Exact("Legend", "top"),
                AttrCheck.Hex("PlotColor", "#FCE4D6"),
                AttrCheck.Hex("ChartAreaColor", "#FFF2CC"),
                AttrCheck.Hex("S0.Color", "#ED7D31")));

            list.Add(Typed(p + "mix-plot-area-4", replace, "column", "top",
                Merge(Node("data-title", "区色4"), Node("data-plot-color", "none"),
                    Node("data-chart-area-color", "none")),
                Th(Pair("data-color", "#4472C4")),
                AttrCheck.Exact("Title", "区色4"),
                AttrCheck.Exact("Legend", "top"),
                AttrCheck.Exact("PlotColor", "none"),
                AttrCheck.Exact("ChartAreaColor", "none"),
                AttrCheck.Hex("S0.Color", "#4472C4")));

            list.Add(Typed(p + "mix-plot-area-5", replace, "column", "top",
                Merge(Node("data-title", "区色5"), Node("data-plot-color", "#DDEBF7"),
                    Node("data-chart-area-color", "#FFFFFF")),
                Th(Pair("data-color", "#C00000")),
                AttrCheck.Exact("Title", "区色5"),
                AttrCheck.Exact("Legend", "top"),
                AttrCheck.Hex("PlotColor", "#DDEBF7"),
                AttrCheck.Hex("ChartAreaColor", "#FFFFFF"),
                AttrCheck.Hex("S0.Color", "#C00000")));

            list.Add(Column(p + "mix-label-chrome-1", replace,
                Merge(Node("data-show-data-labels", "true"), Node("data-show-value", "true"),
                    Node("data-gap-width", "85")),
                Th(Pair("data-show-data-labels", "true"), Pair("data-show-value", "true"),
                    Pair("data-label-position", "outside"), Pair("data-label-font", "微软雅黑"),
                    Pair("data-label-size", "8"), Pair("data-label-color", "#333333"),
                    Pair("data-label-format", "0.0")),
                AttrCheck.Num("GapWidth", "85"),
                AttrCheck.Exact("S0.ShowDataLabels", "true"),
                AttrCheck.Exact("S0.LabelPosition", "outside"),
                AttrCheck.Exact("S0.LabelFont", "微软雅黑"),
                AttrCheck.Num("S0.LabelSize", "8"),
                AttrCheck.Hex("S0.LabelColor", "#333333"),
                AttrCheck.Exact("S0.LabelFormat", "0.0")));

            list.Add(Column(p + "mix-label-chrome-2", replace,
                Merge(Node("data-show-data-labels", "true"), Node("data-show-value", "true"),
                    Node("data-gap-width", "90")),
                Th(Pair("data-show-data-labels", "true"), Pair("data-show-value", "true"),
                    Pair("data-label-position", "outside"), Pair("data-label-font", "微软雅黑"),
                    Pair("data-label-size", "10"), Pair("data-label-color", "#C00000"),
                    Pair("data-label-format", "0.0")),
                AttrCheck.Num("GapWidth", "90"),
                AttrCheck.Exact("S0.ShowDataLabels", "true"),
                AttrCheck.Exact("S0.LabelPosition", "outside"),
                AttrCheck.Exact("S0.LabelFont", "微软雅黑"),
                AttrCheck.Num("S0.LabelSize", "10"),
                AttrCheck.Hex("S0.LabelColor", "#C00000"),
                AttrCheck.Exact("S0.LabelFormat", "0.0")));

            list.Add(Column(p + "mix-label-chrome-3", replace,
                Merge(Node("data-show-data-labels", "true"), Node("data-show-value", "true"),
                    Node("data-gap-width", "95")),
                Th(Pair("data-show-data-labels", "true"), Pair("data-show-value", "true"),
                    Pair("data-label-position", "outside"), Pair("data-label-font", "微软雅黑"),
                    Pair("data-label-size", "11"), Pair("data-label-color", "#0070C0"),
                    Pair("data-label-format", "0.0")),
                AttrCheck.Num("GapWidth", "95"),
                AttrCheck.Exact("S0.ShowDataLabels", "true"),
                AttrCheck.Exact("S0.LabelPosition", "outside"),
                AttrCheck.Exact("S0.LabelFont", "微软雅黑"),
                AttrCheck.Num("S0.LabelSize", "11"),
                AttrCheck.Hex("S0.LabelColor", "#0070C0"),
                AttrCheck.Exact("S0.LabelFormat", "0.0")));

            list.Add(Column(p + "mix-label-chrome-4", replace,
                Merge(Node("data-show-data-labels", "true"), Node("data-show-value", "true"),
                    Node("data-gap-width", "100")),
                Th(Pair("data-show-data-labels", "true"), Pair("data-show-value", "true"),
                    Pair("data-label-position", "outside"), Pair("data-label-font", "微软雅黑"),
                    Pair("data-label-size", "9"), Pair("data-label-color", "#548235"),
                    Pair("data-label-format", "0.0")),
                AttrCheck.Num("GapWidth", "100"),
                AttrCheck.Exact("S0.ShowDataLabels", "true"),
                AttrCheck.Exact("S0.LabelPosition", "outside"),
                AttrCheck.Exact("S0.LabelFont", "微软雅黑"),
                AttrCheck.Num("S0.LabelSize", "9"),
                AttrCheck.Hex("S0.LabelColor", "#548235"),
                AttrCheck.Exact("S0.LabelFormat", "0.0")));

            list.Add(Column(p + "mix-label-chrome-5", replace,
                Merge(Node("data-show-data-labels", "true"), Node("data-show-value", "true"),
                    Node("data-gap-width", "105")),
                Th(Pair("data-show-data-labels", "true"), Pair("data-show-value", "true"),
                    Pair("data-label-position", "outside"), Pair("data-label-font", "微软雅黑"),
                    Pair("data-label-size", "12"), Pair("data-label-color", "#7030A0"),
                    Pair("data-label-format", "0.0")),
                AttrCheck.Num("GapWidth", "105"),
                AttrCheck.Exact("S0.ShowDataLabels", "true"),
                AttrCheck.Exact("S0.LabelPosition", "outside"),
                AttrCheck.Exact("S0.LabelFont", "微软雅黑"),
                AttrCheck.Num("S0.LabelSize", "12"),
                AttrCheck.Hex("S0.LabelColor", "#7030A0"),
                AttrCheck.Exact("S0.LabelFormat", "0.0")));

            list.Add(Typed(p + "mix-style-chrome-1", replace, "column", "bottom",
                Merge(Node("data-title", "样式A"), Node("data-chart-style", "286"),
                    Node("data-axis-y-min", "0"), Node("data-axis-y-max", "11"),
                    Node("data-gridlines", "false")),
                Th(Pair("data-color", "#2497EA")),
                AttrCheck.Exact("Title", "样式A"),
                AttrCheck.Num("ChartStyle", "286"),
                AttrCheck.Num("AxisYMax", "11"),
                AttrCheck.Exact("Gridlines", "false"),
                AttrCheck.Hex("S0.Color", "#2497EA")));

            list.Add(Typed(p + "mix-style-chrome-2", replace, "column", "bottom",
                Merge(Node("data-title", "样式B"), Node("data-chart-style", "286"),
                    Node("data-axis-y-min", "0"), Node("data-axis-y-max", "12"),
                    Node("data-gridlines", "false")),
                Th(Pair("data-color", "#ED7D31")),
                AttrCheck.Exact("Title", "样式B"),
                AttrCheck.Num("ChartStyle", "286"),
                AttrCheck.Num("AxisYMax", "12"),
                AttrCheck.Exact("Gridlines", "false"),
                AttrCheck.Hex("S0.Color", "#ED7D31")));

            list.Add(Typed(p + "mix-style-chrome-3", replace, "column", "bottom",
                Merge(Node("data-title", "样式C"), Node("data-chart-style", "286"),
                    Node("data-axis-y-min", "0"), Node("data-axis-y-max", "13"),
                    Node("data-gridlines", "false")),
                Th(Pair("data-color", "#70AD47")),
                AttrCheck.Exact("Title", "样式C"),
                AttrCheck.Num("ChartStyle", "286"),
                AttrCheck.Num("AxisYMax", "13"),
                AttrCheck.Exact("Gridlines", "false"),
                AttrCheck.Hex("S0.Color", "#70AD47")));

            list.Add(Typed(p + "mix-style-chrome-4", replace, "column", "bottom",
                Merge(Node("data-title", "样式D"), Node("data-chart-style", "286"),
                    Node("data-axis-y-min", "0"), Node("data-axis-y-max", "14"),
                    Node("data-gridlines", "false")),
                Th(Pair("data-color", "#FFC000")),
                AttrCheck.Exact("Title", "样式D"),
                AttrCheck.Num("ChartStyle", "286"),
                AttrCheck.Num("AxisYMax", "14"),
                AttrCheck.Exact("Gridlines", "false"),
                AttrCheck.Hex("S0.Color", "#FFC000")));

            list.Add(Typed(p + "mix-style-chrome-5", replace, "column", "bottom",
                Merge(Node("data-title", "样式E"), Node("data-chart-style", "286"),
                    Node("data-axis-y-min", "0"), Node("data-axis-y-max", "15"),
                    Node("data-gridlines", "false")),
                Th(Pair("data-color", "#7030A0")),
                AttrCheck.Exact("Title", "样式E"),
                AttrCheck.Num("ChartStyle", "286"),
                AttrCheck.Num("AxisYMax", "15"),
                AttrCheck.Exact("Gridlines", "false"),
                AttrCheck.Hex("S0.Color", "#7030A0")));

            list.Add(Column(p + "mix-y-extras-scale-1", replace,
                Merge(Node("data-axis-y", "轴皮1"),
                    Node("data-axis-y-min", "0"), Node("data-axis-y-max", "10"),
                    Node("data-axis-y-major-unit", "2"),
                    AxisPack("data-axis-y", "#FF0000")),
                ConcatChecks(
                    new[]
                    {
                        AttrCheck.Exact("AxisY", "轴皮1"),
                        AttrCheck.Num("AxisYMin", "0"),
                        AttrCheck.Num("AxisYMax", "10")
                    },
                    AxisExpects("AxisYStyle", "#FF0000"))));

            list.Add(Column(p + "mix-y-extras-scale-2", replace,
                Merge(Node("data-axis-y", "轴皮2"),
                    Node("data-axis-y-min", "0"), Node("data-axis-y-max", "12"),
                    Node("data-axis-y-major-unit", "2"),
                    AxisPack("data-axis-y", "#0070C0")),
                ConcatChecks(
                    new[]
                    {
                        AttrCheck.Exact("AxisY", "轴皮2"),
                        AttrCheck.Num("AxisYMin", "0"),
                        AttrCheck.Num("AxisYMax", "12")
                    },
                    AxisExpects("AxisYStyle", "#0070C0"))));

            list.Add(Column(p + "mix-y-extras-scale-3", replace,
                Merge(Node("data-axis-y", "轴皮3"),
                    Node("data-axis-y-min", "0"), Node("data-axis-y-max", "14"),
                    Node("data-axis-y-major-unit", "2"),
                    AxisPack("data-axis-y", "#00B050")),
                ConcatChecks(
                    new[]
                    {
                        AttrCheck.Exact("AxisY", "轴皮3"),
                        AttrCheck.Num("AxisYMin", "0"),
                        AttrCheck.Num("AxisYMax", "14")
                    },
                    AxisExpects("AxisYStyle", "#00B050"))));

            list.Add(Column(p + "mix-y-extras-scale-4", replace,
                Merge(Node("data-axis-y", "轴皮4"),
                    Node("data-axis-y-min", "0"), Node("data-axis-y-max", "16"),
                    Node("data-axis-y-major-unit", "5"),
                    AxisPack("data-axis-y", "#7030A0")),
                ConcatChecks(
                    new[]
                    {
                        AttrCheck.Exact("AxisY", "轴皮4"),
                        AttrCheck.Num("AxisYMin", "0"),
                        AttrCheck.Num("AxisYMax", "16")
                    },
                    AxisExpects("AxisYStyle", "#7030A0"))));

            list.Add(Column(p + "mix-y-extras-scale-5", replace,
                Merge(Node("data-axis-y", "轴皮5"),
                    Node("data-axis-y-min", "0"), Node("data-axis-y-max", "18"),
                    Node("data-axis-y-major-unit", "5"),
                    AxisPack("data-axis-y", "#C65911")),
                ConcatChecks(
                    new[]
                    {
                        AttrCheck.Exact("AxisY", "轴皮5"),
                        AttrCheck.Num("AxisYMin", "0"),
                        AttrCheck.Num("AxisYMax", "18")
                    },
                    AxisExpects("AxisYStyle", "#C65911"))));

            list.Add(Combo(p + "mix-y2-skin-1", replace,
                Merge(Node("data-title", "双轴1"), Node("data-axis-y", "主1"),
                    Node("data-axis-y-secondary", "次1"),
                    Node("data-axis-y-min", "0"), Node("data-axis-y-max", "9")),
                Th(Pair("data-color", "#5B9BD5"), Pair("data-show-data-labels", "false")),
                Th(Pair("data-line", "#ED7D31"), Pair("data-line-weight", "2"),
                    Pair("data-marker", "circle"), Pair("data-marker-size", "6")),
                AttrCheck.Exact("Title", "双轴1"),
                AttrCheck.Exact("Legend", "bottom"),
                AttrCheck.Exact("AxisY", "主1"),
                AttrCheck.Exact("AxisYSecondary", "次1"),
                AttrCheck.Num("AxisYMax", "9"),
                AttrCheck.Exact("S1.AxisY", "y2"),
                AttrCheck.Hex("S0.Color", "#5B9BD5"),
                AttrCheck.Hex("S1.Line", "#ED7D31"),
                AttrCheck.Exact("S1.Marker", "circle")));

            list.Add(Combo(p + "mix-y2-skin-2", replace,
                Merge(Node("data-title", "双轴2"), Node("data-axis-y", "主2"),
                    Node("data-axis-y-secondary", "次2"),
                    Node("data-axis-y-min", "0"), Node("data-axis-y-max", "10")),
                Th(Pair("data-color", "#70AD47"), Pair("data-show-data-labels", "false")),
                Th(Pair("data-line", "#C00000"), Pair("data-line-weight", "2"),
                    Pair("data-marker", "circle"), Pair("data-marker-size", "6")),
                AttrCheck.Exact("Title", "双轴2"),
                AttrCheck.Exact("Legend", "bottom"),
                AttrCheck.Exact("AxisY", "主2"),
                AttrCheck.Exact("AxisYSecondary", "次2"),
                AttrCheck.Num("AxisYMax", "10"),
                AttrCheck.Exact("S1.AxisY", "y2"),
                AttrCheck.Hex("S0.Color", "#70AD47"),
                AttrCheck.Hex("S1.Line", "#C00000"),
                AttrCheck.Exact("S1.Marker", "circle")));

            list.Add(Combo(p + "mix-y2-skin-3", replace,
                Merge(Node("data-title", "双轴3"), Node("data-axis-y", "主3"),
                    Node("data-axis-y-secondary", "次3"),
                    Node("data-axis-y-min", "0"), Node("data-axis-y-max", "11")),
                Th(Pair("data-color", "#4472C4"), Pair("data-show-data-labels", "false")),
                Th(Pair("data-line", "#FFC000"), Pair("data-line-weight", "2"),
                    Pair("data-marker", "circle"), Pair("data-marker-size", "6")),
                AttrCheck.Exact("Title", "双轴3"),
                AttrCheck.Exact("Legend", "bottom"),
                AttrCheck.Exact("AxisY", "主3"),
                AttrCheck.Exact("AxisYSecondary", "次3"),
                AttrCheck.Num("AxisYMax", "11"),
                AttrCheck.Exact("S1.AxisY", "y2"),
                AttrCheck.Hex("S0.Color", "#4472C4"),
                AttrCheck.Hex("S1.Line", "#FFC000"),
                AttrCheck.Exact("S1.Marker", "circle")));

            list.Add(Combo(p + "mix-y2-skin-4", replace,
                Merge(Node("data-title", "双轴4"), Node("data-axis-y", "主4"),
                    Node("data-axis-y-secondary", "次4"),
                    Node("data-axis-y-min", "0"), Node("data-axis-y-max", "12")),
                Th(Pair("data-color", "#7030A0"), Pair("data-show-data-labels", "false")),
                Th(Pair("data-line", "#00B0F0"), Pair("data-line-weight", "2"),
                    Pair("data-marker", "circle"), Pair("data-marker-size", "6")),
                AttrCheck.Exact("Title", "双轴4"),
                AttrCheck.Exact("Legend", "bottom"),
                AttrCheck.Exact("AxisY", "主4"),
                AttrCheck.Exact("AxisYSecondary", "次4"),
                AttrCheck.Num("AxisYMax", "12"),
                AttrCheck.Exact("S1.AxisY", "y2"),
                AttrCheck.Hex("S0.Color", "#7030A0"),
                AttrCheck.Hex("S1.Line", "#00B0F0"),
                AttrCheck.Exact("S1.Marker", "circle")));

            list.Add(Combo(p + "mix-y2-skin-5", replace,
                Merge(Node("data-title", "双轴5"), Node("data-axis-y", "主5"),
                    Node("data-axis-y-secondary", "次5"),
                    Node("data-axis-y-min", "0"), Node("data-axis-y-max", "13")),
                Th(Pair("data-color", "#ED7D31"), Pair("data-show-data-labels", "false")),
                Th(Pair("data-line", "#5B9BD5"), Pair("data-line-weight", "2"),
                    Pair("data-marker", "circle"), Pair("data-marker-size", "6")),
                AttrCheck.Exact("Title", "双轴5"),
                AttrCheck.Exact("Legend", "bottom"),
                AttrCheck.Exact("AxisY", "主5"),
                AttrCheck.Exact("AxisYSecondary", "次5"),
                AttrCheck.Num("AxisYMax", "13"),
                AttrCheck.Exact("S1.AxisY", "y2"),
                AttrCheck.Hex("S0.Color", "#ED7D31"),
                AttrCheck.Hex("S1.Line", "#5B9BD5"),
                AttrCheck.Exact("S1.Marker", "circle")));
        }
    }
}
