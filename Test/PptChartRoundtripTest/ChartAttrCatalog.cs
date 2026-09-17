using System;
using System.Collections.Generic;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PptChartRoundtripTest
{
    /// <summary>可写属性正式用例：50 新建 + 50 改已有，两份 PPT。</summary>
    internal static class ChartAttrCatalog
    {
        public const int BatchSize = 63;

        public const int BatchCount = 2;

        public const int Total = BatchSize * BatchCount;

        private static readonly string[] Years = { "2026", "2027", "2028", "2029" };

        private static readonly double[] ColVals = { 2.15, 2.8, 3.65, 4.8 };

        private static readonly double[] LineVals = { 0.265, 0.302, 0.304, 0.315 };

        private static readonly string[] PieCats = { "A", "B", "C", "D" };

        private static readonly double[] PieVals = { 40, 30, 20, 10 };

        private static readonly string[] PieCatsMore = { "A", "B", "C", "D", "E" };

        private static readonly double[] PieValsMore = { 40, 30, 20, 10, 5 };

        public static List<ChartCase> Build()
        {
            var creates = new List<ChartCase>();
            AddAll(creates, replace: false);
            var replaces = new List<ChartCase>();
            AddAll(replaces, replace: true);

            var list = new List<ChartCase>(Total);
            list.AddRange(creates);
            list.AddRange(replaces);
            for (int i = 0; i < list.Count; i++)
            {
                list[i].Index = i + 1;
                list[i].Batch = i / BatchSize + 1;
                list[i].Page = i % BatchSize + 1;
            }

            return list;
        }

        private static void AddAll(List<ChartCase> list, bool replace)
        {
            string p = replace ? "replace-" : "create-";
            list.Add(SinkColumn(p, replace));
            list.Add(SinkLine(p, replace));
            list.Add(SinkPie(p, replace));
            list.Add(SinkCombo(p, replace));
            list.Add(Column(p + "title-text", replace, Node("data-title", "属性测标题"),
                AttrCheck.Exact("Title", "属性测标题")));
            list.Add(Column(p + "title-font-size", replace,
                Merge(Node("data-title", "属性测标题"), Node("data-title-font-size", "16")),
                AttrCheck.Exact("Title", "属性测标题"), AttrCheck.Num("TitleFontSize", "16")));
            list.Add(Column(p + "title-font-bold", replace,
                Merge(Node("data-title", "属性测标题"), Node("data-title-font-bold", "false")),
                AttrCheck.Exact("Title", "属性测标题"), AttrCheck.Exact("TitleFontBold", "false")));
            list.Add(Column(p + "title-font-color", replace,
                Merge(Node("data-title", "属性测标题"), Node("data-title-font-color", "#C00000")),
                AttrCheck.Exact("Title", "属性测标题"), AttrCheck.Hex("TitleFontColor", "#C00000")));
            list.Add(Column(p + "title-off", replace, Node("data-title", ""),
                AttrCheck.Exact("Title", null)));
            list.Add(Column(p + "title-full", replace,
                Merge(
                    Node("data-title", "完整标题样式"),
                    Node("data-title-font-size", "18"),
                    Node("data-title-font-bold", "true"),
                    Node("data-title-font-color", "#0070C0")),
                AttrCheck.Exact("Title", "完整标题样式"),
                AttrCheck.Num("TitleFontSize", "18"),
                AttrCheck.Exact("TitleFontBold", "true"),
                AttrCheck.Hex("TitleFontColor", "#0070C0")));
            list.Add(Column(p + "title-long", replace,
                Node("data-title", "一二三四五六七八九十年增长率对比图"),
                AttrCheck.Exact("Title", "一二三四五六七八九十年增长率对比图")));
            list.Add(Line(p + "title-off-line", replace, Node("data-title", ""), null,
                AttrCheck.Exact("Title", null)));
            list.Add(Pie(p + "title-off-pie", replace, Node("data-title", ""), null,
                AttrCheck.Exact("Title", null)));
            list.Add(Combo(p + "title-off-combo", replace, Node("data-title", ""),
                AttrCheck.Exact("Title", null)));
            list.Add(TitleOffFromTitled(p + "title-off-from-on", replace));
            list.Add(TitleKeepFromOld(p + "title-keep-from-old", replace));
            list.Add(TitleChangeText(p + "title-change-text", replace));
            list.Add(Typed(p + "legend-top", replace, "column", "top", null, null,
                AttrCheck.Exact("Legend", "top")));
            list.Add(Typed(p + "legend-left", replace, "column", "left", null, null,
                AttrCheck.Exact("Legend", "left")));
            list.Add(Typed(p + "legend-right", replace, "column", "right", null, null,
                AttrCheck.Exact("Legend", "right")));
            list.Add(Typed(p + "legend-bottom", replace, "column", "bottom", null, null,
                AttrCheck.Exact("Legend", "bottom")));
            list.Add(Typed(p + "legend-none", replace, "column", "none", null, null,
                AttrCheck.Exact("Legend", "none")));
            list.Add(Typed(p + "legend-font-color", replace, "column", "bottom",
                Node("data-legend-font-color", "#00B050"), null,
                AttrCheck.Exact("Legend", "bottom"), AttrCheck.Hex("LegendFontColor", "#00B050")));
            list.Add(Column(p + "labels-on-value", replace,
                Merge(Node("data-show-data-labels", "true"), Node("data-show-value", "true")),
                AttrCheck.Exact("S0.ShowDataLabels", "true"), AttrCheck.Exact("S0.ShowValue", "true")));
            list.Add(Column(p + "labels-off", replace, Node("data-show-data-labels", "false"),
                AttrCheck.Exact("S0.ShowDataLabels", "false")));
            list.Add(Pie(p + "labels-pie-pct", replace,
                Merge(Node("data-show-data-labels", "true"),
                    Node("data-show-value", "false"),
                    Node("data-show-percentage", "true")),
                Th(Pair("data-show-data-labels", "true"),
                    Pair("data-show-value", "false"),
                    Pair("data-show-percentage", "true")),
                AttrCheck.Exact("S0.ShowDataLabels", "true"),
                AttrCheck.Exact("S0.ShowValue", "false"),
                AttrCheck.Exact("S0.ShowPercentage", "true")));
            list.Add(Column(p + "gridlines-on", replace, Node("data-gridlines", "true"),
                AttrCheck.Exact("Gridlines", "true")));
            list.Add(Column(p + "gridlines-off", replace, Node("data-gridlines", "false"),
                AttrCheck.Exact("Gridlines", "false")));
            list.Add(Column(p + "gap-width", replace, Node("data-gap-width", "80"),
                AttrCheck.Num("GapWidth", "80")));
            list.Add(Column(p + "overlap", replace, Node("data-overlap", "-30"),
                AttrCheck.Num("Overlap", "-30")));
            list.Add(Column(p + "plot-color", replace, Node("data-plot-color", "#FFF2CC"),
                AttrCheck.Hex("PlotColor", "#FFF2CC")));
            list.Add(Column(p + "chart-area-color", replace, Node("data-chart-area-color", "#F2F2F2"),
                AttrCheck.Hex("ChartAreaColor", "#F2F2F2")));
            list.Add(Column(p + "plot-none", replace,
                Merge(Node("data-plot-color", "none"), Node("data-chart-area-color", "none")),
                AttrCheck.Exact("PlotColor", "none"), AttrCheck.Exact("ChartAreaColor", "none")));
            list.Add(Column(p + "plot-box", replace, Node("data-plot-box", "20,25,400,200"),
                AttrCheck.ApplyOnly("PlotBox", "20,25,400,200")));
            list.Add(Column(p + "plot-inside", replace, Node("data-plot-inside", "30,35,360,170"),
                AttrCheck.ApplyOnly("PlotInside", "30,35,360,170")));
            list.Add(Column(p + "axis-y-scale", replace,
                Merge(Node("data-axis-y-min", "0"), Node("data-axis-y-max", "20"),
                    Node("data-axis-y-major-unit", "5")),
                AttrCheck.Num("AxisYMin", "0"), AttrCheck.Num("AxisYMax", "20"),
                AttrCheck.Num("AxisYMajorUnit", "5")));
            list.Add(Column(p + "axis-x-title", replace, Node("data-axis-x", "类目轴"),
                AttrCheck.Exact("AxisX", "类目轴")));
            list.Add(Column(p + "axis-y-title", replace, Node("data-axis-y", "数值轴"),
                AttrCheck.Exact("AxisY", "数值轴")));
            list.Add(Combo(p + "axis-y-secondary", replace,
                Node("data-axis-y-secondary", "次轴"),
                AttrCheck.Exact("AxisYSecondary", "次轴"),
                AttrCheck.Exact("S1.AxisY", "y2")));
            list.Add(Column(p + "axis-x-extras", replace, AxisPack("data-axis-x", "#FF0000"),
                AxisExpects("AxisXStyle", "#FF0000")));
            list.Add(Column(p + "axis-y-extras", replace, AxisPack("data-axis-y", "#0070C0"),
                AxisExpects("AxisYStyle", "#0070C0")));
            list.Add(Combo(p + "axis-y2-extras", replace, AxisPack("data-axis-y2", "#7030A0"),
                AxisExpects("AxisY2Style", "#7030A0")));
            list.Add(Column(p + "axis-x-hidden", replace, Node("data-axis-x-visible", "false"),
                AttrCheck.Exact("AxisXStyle.Visible", "false")));
            list.Add(Column(p + "axis-y-hidden", replace, Node("data-axis-y-visible", "false"),
                AttrCheck.Exact("AxisYStyle.Visible", "false")));
            list.Add(Column(p + "series-color", replace, null,
                Th(Pair("data-color", "#2497EA")),
                AttrCheck.Hex("S0.Color", "#2497EA")));
            list.Add(Column(p + "series-color-none", replace, null,
                Th(Pair("data-color", "none")),
                AttrCheck.Exact("S0.Color", "none")));
            list.Add(SeriesColorKeepNone(p + "series-color-keep-none", replace));
            list.Add(Column(p + "series-gradient", replace, null,
                Th(Pair("data-fill-gradient", "0:#FFFFFF@1;1:#2497EA@0"), Pair("data-fill-angle", "270")),
                AttrCheck.Exact("S0.FillGradient", "0:#FFFFFF@1;1:#2497EA@0"),
                AttrCheck.Num("S0.FillAngle", "270")));
            list.Add(Line(p + "series-line", replace, null,
                Th(Pair("data-line", "#ED7D31"), Pair("data-line-weight", "2.5")),
                AttrCheck.Hex("S0.Line", "#ED7D31"), AttrCheck.Num("S0.LineWeight", "2.5")));
            list.Add(Line(p + "series-marker", replace, null,
                Th(Pair("data-marker", "diamond"), Pair("data-marker-size", "8"),
                    Pair("data-marker-color", "#00B0F0"), Pair("data-marker-fill", "#FFFFFF")),
                AttrCheck.Exact("S0.Marker", "diamond"), AttrCheck.Num("S0.MarkerSize", "8"),
                AttrCheck.Hex("S0.MarkerColor", "#00B0F0"), AttrCheck.Hex("S0.MarkerFill", "#FFFFFF")));
            list.Add(Column(p + "series-label-chrome", replace, null,
                Th(Pair("data-show-data-labels", "true"), Pair("data-show-value", "true"),
                    Pair("data-label-position", "outside"), Pair("data-label-font", "微软雅黑"),
                    Pair("data-label-size", "9"), Pair("data-label-color", "#333333"),
                    Pair("data-label-format", "0.00")),
                AttrCheck.Exact("S0.ShowDataLabels", "true"),
                AttrCheck.Exact("S0.LabelPosition", "outside"),
                AttrCheck.Exact("S0.LabelFont", "微软雅黑"),
                AttrCheck.Num("S0.LabelSize", "9"),
                AttrCheck.Hex("S0.LabelColor", "#333333"),
                AttrCheck.Exact("S0.LabelFormat", "0.00")));
            // 柱图稿写 above：写路径归一成 outside（COM 无 Above）
            list.Add(Column(p + "series-label-above-coerced", replace, null,
                Th(Pair("data-show-data-labels", "true"), Pair("data-show-value", "true"),
                    Pair("data-label-position", "above")),
                AttrCheck.Exact("S0.ShowDataLabels", "true"),
                AttrCheck.Exact("S0.LabelPosition", "outside")));
            // 折线仍支持 above
            list.Add(Line(p + "series-label-above-line", replace, null,
                Th(Pair("data-show-data-labels", "true"), Pair("data-show-value", "true"),
                    Pair("data-label-position", "above")),
                AttrCheck.Exact("S0.ShowDataLabels", "true"),
                AttrCheck.Exact("S0.LabelPosition", "above")));
            list.Add(Pie(p + "series-label-outside", replace, Node("data-show-data-labels", "true"),
                Th(Pair("data-show-data-labels", "true"), Pair("data-show-percentage", "true"),
                    Pair("data-show-value", "false"), Pair("data-label-position", "outside")),
                AttrCheck.Exact("S0.LabelPosition", "outside")));
            list.Add(Pie(p + "explosion", replace, Node("data-explosion", "12"), null,
                AttrCheck.Exact("Explosion", "12")));
            list.Add(ExplosionKeepAll(p + "explosion-keep-all", replace));
            list.Add(ExplosionSkipPartial(p + "explosion-skip-partial", replace));
            list.Add(Column(p + "chart-style", replace, Node("data-chart-style", "286"),
                AttrCheck.Num("ChartStyle", "286")));
            list.Add(Column(p + "theme", replace, Node("data-theme", "office"),
                AttrCheck.ApplyOnly("Theme", "office")));
            list.Add(Column(p + "fill-missing", replace, Node("data-fill-missing", "gap"),
                AttrCheck.ApplyOnly("FillMissing", "gap")));
            list.Add(Column(p + "axis-x-ticks", replace,
                Merge(Node("data-axis-x-type", "category"), Node("data-axis-x-format", "General"),
                    Node("data-axis-x-tick-count", "4"), Node("data-axis-x-tick-spacing", "1"),
                    Node("data-axis-x-between", "true")),
                AttrCheck.ApplyOnly("AxisXType", "category"),
                AttrCheck.ApplyOnly("AxisXFormat", "General"),
                AttrCheck.ApplyOnly("AxisXTickCount", "4"),
                AttrCheck.ApplyOnly("AxisXTickSpacing", "1"),
                AttrCheck.ApplyOnly("AxisXBetween", "true")));
            list.Add(Line(p + "data-markers", replace, Node("data-data-markers", "true"),
                Th(Pair("data-marker", "circle"), Pair("data-marker-size", "7")),
                AttrCheck.ApplyOnly("DataMarkers", "true"),
                AttrCheck.Exact("S0.Marker", "circle"),
                AttrCheck.Num("S0.MarkerSize", "7")));
            list.Add(Line(p + "chart-line-weight", replace, Node("data-chart-line-weight", "2.5"),
                Th(Pair("data-line-weight", "2.5")),
                AttrCheck.ApplyOnly("ChartLineWeight", "2.5"),
                AttrCheck.Num("S0.LineWeight", "2.5")));
        }

        private static ChartCase SinkColumn(string prefix, bool replace)
        {
            var node = Merge(
                Node("data-title", "市场规模"),
                Node("data-title-font-size", "14"),
                Node("data-title-font-bold", "true"),
                Node("data-title-font-color", "#C00000"),
                Node("data-legend-font-color", "#00B050"),
                Node("data-show-data-labels", "true"),
                Node("data-show-value", "true"),
                Node("data-gridlines", "false"),
                Node("data-gap-width", "110"),
                Node("data-overlap", "-20"),
                Node("data-plot-color", "none"),
                Node("data-chart-area-color", "none"),
                Node("data-axis-y-min", "0"),
                Node("data-axis-y-max", "10"),
                Node("data-axis-y-major-unit", "2"),
                Node("data-axis-x", "年份"),
                Node("data-axis-y", "规模"),
                AxisPack("data-axis-x", "#333333"),
                AxisPack("data-axis-y", "#333333"));
            var th = Th(
                Pair("data-color", "#2497EA"),
                Pair("data-show-data-labels", "true"),
                Pair("data-show-value", "true"),
                Pair("data-label-position", "outside"),
                Pair("data-label-font", "微软雅黑"),
                Pair("data-label-size", "10"),
                Pair("data-label-color", "#333333"),
                Pair("data-label-format", "0.00"));
            return Typed(prefix + "sink-column", replace, "column", "bottom", node, th,
                AttrCheck.Exact("Title", "市场规模"),
                AttrCheck.Num("TitleFontSize", "14"),
                AttrCheck.Exact("TitleFontBold", "true"),
                AttrCheck.Hex("TitleFontColor", "#C00000"),
                AttrCheck.Exact("Legend", "bottom"),
                AttrCheck.Hex("LegendFontColor", "#00B050"),
                AttrCheck.Num("GapWidth", "110"),
                AttrCheck.Num("Overlap", "-20"),
                AttrCheck.Exact("Gridlines", "false"),
                AttrCheck.Exact("AxisX", "年份"),
                AttrCheck.Exact("AxisY", "规模"),
                AttrCheck.Num("AxisYMin", "0"),
                AttrCheck.Num("AxisYMax", "10"),
                AttrCheck.Hex("S0.Color", "#2497EA"),
                AttrCheck.Exact("S0.LabelPosition", "outside"),
                AttrCheck.Exact("S0.LabelFont", "微软雅黑"));
        }

        private static ChartCase SinkLine(string prefix, bool replace)
        {
            var th = Th(
                Pair("data-line", "#00B0F0"),
                Pair("data-line-weight", "2"),
                Pair("data-marker", "circle"),
                Pair("data-marker-size", "6"),
                Pair("data-marker-color", "#00B0F0"),
                Pair("data-marker-fill", "#FFFFFF"),
                Pair("data-show-data-labels", "true"),
                Pair("data-show-value", "true"),
                Pair("data-label-position", "above"));
            return Typed(prefix + "sink-line", replace, "line", "right",
                Node("data-data-markers", "true"), th,
                AttrCheck.Exact("Legend", "right"),
                AttrCheck.Hex("S0.Line", "#00B0F0"),
                AttrCheck.Num("S0.LineWeight", "2"),
                AttrCheck.Exact("S0.Marker", "circle"),
                AttrCheck.Num("S0.MarkerSize", "6"),
                AttrCheck.Hex("S0.MarkerColor", "#00B0F0"),
                AttrCheck.Hex("S0.MarkerFill", "#FFFFFF"));
        }

        private static ChartCase SinkPie(string prefix, bool replace)
        {
            return Pie(prefix + "sink-pie", replace,
                Merge(Node("data-explosion", "8"),
                    Node("data-show-data-labels", "true"),
                    Node("data-show-value", "false"),
                    Node("data-show-percentage", "true")),
                Th(Pair("data-show-data-labels", "true"),
                    Pair("data-show-value", "false"),
                    Pair("data-show-percentage", "true"),
                    Pair("data-label-position", "outside")),
                AttrCheck.Exact("Explosion", "8"),
                AttrCheck.Exact("S0.ShowPercentage", "true"),
                AttrCheck.Exact("S0.ShowValue", "false"),
                AttrCheck.Exact("S0.LabelPosition", "outside"));
        }

        private static ChartCase SinkCombo(string prefix, bool replace)
        {
            var node = Merge(
                Node("data-axis-y", "规模"),
                Node("data-axis-y-secondary", "增速"),
                AxisPack("data-axis-y2", "#7030A0"));
            var colTh = Th(Pair("data-color", "#2497EA"), Pair("data-show-data-labels", "false"));
            var lineTh = Th(
                Pair("data-line", "#ED7D31"),
                Pair("data-line-weight", "2"),
                Pair("data-marker", "circle"),
                Pair("data-marker-size", "6"),
                Pair("data-show-data-labels", "true"),
                Pair("data-show-value", "true"),
                Pair("data-label-format", "0.0%"));
            return Combo(prefix + "sink-combo", replace, node, colTh, lineTh,
                AttrCheck.Exact("AxisY", "规模"),
                AttrCheck.Exact("AxisYSecondary", "增速"),
                AttrCheck.Exact("S1.AxisY", "y2"),
                AttrCheck.Hex("S0.Color", "#2497EA"),
                AttrCheck.Hex("S1.Line", "#ED7D31"),
                AttrCheck.Exact("S1.Marker", "circle"),
                AttrCheck.Exact("AxisY2Style.Visible", "true"),
                AttrCheck.Hex("AxisY2Style.TickColor", "#7030A0"));
        }

        /// <summary>底图系列无色，换数稿不提 data-color → 应继承 none。</summary>
        private static ChartCase SeriesColorKeepNone(string name, bool replace)
        {
            if (!replace)
            {
                return Column(name, false, null,
                    Th(Pair("data-color", "none")),
                    AttrCheck.Exact("S0.Color", "none"));
            }

            IList<double>[] series = { ColVals };
            string baseline = ChartHtml.BuildAttrChart(
                true,
                "column",
                "none",
                Years,
                series,
                new[] { "column" },
                new[] { "y" },
                new[] { "底图" },
                null,
                new IDictionary<string, string>[] { Th(Pair("data-color", "none")) });
            string html = ChartHtml.BuildAttrChart(
                false,
                "column",
                "none",
                Years,
                series,
                new[] { "column" },
                new[] { "y" },
                new[] { "系列A" },
                null,
                null);
            return Finish(name, true, "column", 4, 1, null, Years[0], ColVals[0], ColVals[3],
                baseline, html, new[] { AttrCheck.Exact("S0.Color", "none") });
        }

        /// <summary>底图全瓣统一爆炸，换数增类别且稿不提 → 新全部瓣继承该值。</summary>
        private static ChartCase ExplosionKeepAll(string name, bool replace)
        {
            if (!replace)
            {
                return Pie(name, false, Node("data-explosion", "15"), null,
                    AttrCheck.Exact("Explosion", "15"));
            }

            IList<double>[] baseSeries = { PieVals };
            IList<double>[] moreSeries = { PieValsMore };
            string baseline = ChartHtml.BuildAttrChart(
                true,
                "pie2d",
                "none",
                PieCats,
                baseSeries,
                new[] { "pie2d" },
                new[] { "y" },
                new[] { "底图" },
                Node("data-explosion", "15"),
                null);
            string html = ChartHtml.BuildAttrChart(
                false,
                "pie2d",
                "right",
                PieCatsMore,
                moreSeries,
                new[] { "pie2d" },
                new[] { "y" },
                new[] { "份额" },
                null,
                null);
            return Finish(name, true, "pie2d", 5, 1, null, PieCatsMore[0], PieValsMore[0], PieValsMore[4],
                baseline, html, new[] { AttrCheck.Exact("Explosion", "15") });
        }

        /// <summary>底图仅一瓣爆炸，换数稿不提 → 不继承不对称爆炸。</summary>
        private static ChartCase ExplosionSkipPartial(string name, bool replace)
        {
            if (!replace)
            {
                return Pie(name, false, null, null,
                    AttrCheck.Exact("Explosion", "0"));
            }

            IList<double>[] series = { PieVals };
            string baseline = ChartHtml.BuildAttrChart(
                true,
                "pie2d",
                "none",
                PieCats,
                series,
                new[] { "pie2d" },
                new[] { "y" },
                new[] { "底图" },
                null,
                null);
            string html = ChartHtml.BuildAttrChart(
                false,
                "pie2d",
                "right",
                PieCats,
                series,
                new[] { "pie2d" },
                new[] { "y" },
                new[] { "份额" },
                null,
                null);
            ChartCase c = Finish(name, true, "pie2d", 4, 1, null, PieCats[0], PieVals[0], PieVals[3],
                baseline, html, new[] { AttrCheck.Exact("Explosion", "0") });
            c.AfterCreateMutate = ExplodeFirstPieSliceOnly;
            return c;
        }

        private static void ExplodeFirstPieSliceOnly(object shape)
        {
            if (shape == null)
            {
                throw new InvalidOperationException("底图 shape 为空");
            }

            PowerPoint.Shape shp = shape as PowerPoint.Shape;
            if (shp == null)
            {
                shp = (PowerPoint.Shape)shape;
            }

            PowerPoint.Series ser = (PowerPoint.Series)shp.Chart.SeriesCollection(1);
            PowerPoint.Point pt = (PowerPoint.Point)ser.Points(1);
            pt.Explosion = 25;
        }

        /// <summary>底图有标题，换数稿显式关掉（新建路径等价于直接关）。</summary>
        private static ChartCase TitleOffFromTitled(string name, bool replace)
        {
            if (!replace)
            {
                return Column(name, false, Node("data-title", ""),
                    AttrCheck.Exact("Title", null));
            }

            return ColumnReplaceWithBaseline(
                name,
                Node("data-title", "旧图标题"),
                Node("data-title", ""),
                AttrCheck.Exact("Title", null));
        }

        /// <summary>底图有标题，换数稿不提标题 → 应继承旧正文。</summary>
        private static ChartCase TitleKeepFromOld(string name, bool replace)
        {
            if (!replace)
            {
                return Column(name, false, Node("data-title", "新建保留标题"),
                    AttrCheck.Exact("Title", "新建保留标题"));
            }

            return ColumnReplaceWithBaseline(
                name,
                Node("data-title", "旧图应保留"),
                null,
                AttrCheck.Exact("Title", "旧图应保留"));
        }

        /// <summary>底图有标题，换数稿改成另一段正文。</summary>
        private static ChartCase TitleChangeText(string name, bool replace)
        {
            if (!replace)
            {
                return Column(name, false, Node("data-title", "新建改写标题"),
                    AttrCheck.Exact("Title", "新建改写标题"));
            }

            return ColumnReplaceWithBaseline(
                name,
                Node("data-title", "旧标题"),
                Node("data-title", "新标题"),
                AttrCheck.Exact("Title", "新标题"));
        }

        private static ChartCase ColumnReplaceWithBaseline(
            string name,
            Dictionary<string, string> baselineNode,
            Dictionary<string, string> replaceNode,
            params AttrCheck[] attrs)
        {
            IList<double>[] series = { ColVals };
            string baseline = ChartHtml.BuildAttrChart(
                true,
                "column",
                "none",
                Years,
                series,
                new[] { "column" },
                new[] { "y" },
                new[] { "底图" },
                baselineNode,
                null);
            string html = ChartHtml.BuildAttrChart(
                false,
                "column",
                "none",
                Years,
                series,
                new[] { "column" },
                new[] { "y" },
                new[] { "系列A" },
                replaceNode,
                null);
            return Finish(name, true, "column", 4, 1, null, Years[0], ColVals[0], ColVals[3],
                baseline, html, attrs);
        }

        private static ChartCase Column(string name, bool replace, Dictionary<string, string> node, params AttrCheck[] attrs)
        {
            return Typed(name, replace, "column", "none", node, null, attrs);
        }

        private static ChartCase Column(
            string name,
            bool replace,
            Dictionary<string, string> node,
            Dictionary<string, string> th,
            params AttrCheck[] attrs)
        {
            return Typed(name, replace, "column", "none", node, th, attrs);
        }

        private static ChartCase Line(
            string name,
            bool replace,
            Dictionary<string, string> node,
            Dictionary<string, string> th,
            params AttrCheck[] attrs)
        {
            return Typed(name, replace, "line", "right", node, th, attrs);
        }

        private static ChartCase Pie(
            string name,
            bool replace,
            Dictionary<string, string> node,
            Dictionary<string, string> th,
            params AttrCheck[] attrs)
        {
            IList<double>[] series = { PieVals };
            string html = ChartHtml.BuildAttrChart(
                !replace,
                "pie2d",
                "right",
                PieCats,
                series,
                new[] { "pie2d" },
                new[] { "y" },
                new[] { "份额" },
                node,
                th == null ? null : new IDictionary<string, string>[] { th });
            string baseline = ChartHtml.BuildAttrChart(
                true,
                "pie2d",
                "none",
                PieCats,
                series,
                new[] { "pie2d" },
                new[] { "y" },
                new[] { "底图" },
                null,
                null);
            return Finish(name, replace, "pie2d", 4, 1, null, PieCats[0], PieVals[0], PieVals[3],
                baseline, html, attrs);
        }

        private static ChartCase Combo(string name, bool replace, Dictionary<string, string> node, params AttrCheck[] attrs)
        {
            return Combo(name, replace, node, null, null, attrs);
        }

        private static ChartCase Combo(
            string name,
            bool replace,
            Dictionary<string, string> node,
            Dictionary<string, string> colTh,
            Dictionary<string, string> lineTh,
            params AttrCheck[] attrs)
        {
            IList<double>[] series = { ColVals, LineVals };
            var extraTh = new IDictionary<string, string>[] { colTh, lineTh };
            string html = ChartHtml.BuildAttrChart(
                !replace,
                "column",
                "bottom",
                Years,
                series,
                new[] { "column", "line" },
                new[] { "y", "y2" },
                new[] { "规模", "增速" },
                node,
                extraTh);
            string baseline = ChartHtml.BuildAttrChart(
                true,
                "column",
                "none",
                Years,
                new IList<double>[] { ColVals },
                new[] { "column" },
                new[] { "y" },
                new[] { "底图" },
                null,
                null);
            return Finish(name, replace, "column", 4, 2, "y2", Years[0], ColVals[0], ColVals[3],
                baseline, html, attrs);
        }

        private static ChartCase Typed(
            string name,
            bool replace,
            string type,
            string legend,
            Dictionary<string, string> node,
            Dictionary<string, string> th,
            params AttrCheck[] attrs)
        {
            IList<double>[] series = { ColVals };
            string html = ChartHtml.BuildAttrChart(
                !replace,
                type,
                legend,
                Years,
                series,
                new[] { type },
                new[] { "y" },
                new[] { "系列A" },
                node,
                th == null ? null : new IDictionary<string, string>[] { th });
            string baseline = ChartHtml.BuildAttrChart(
                true,
                type == "line" ? "line" : "column",
                "none",
                Years,
                series,
                new[] { type == "line" ? "line" : "column" },
                new[] { "y" },
                new[] { "底图" },
                null,
                null);
            return Finish(name, replace, type, 4, 1, null, Years[0], ColVals[0], ColVals[3],
                baseline, html, attrs);
        }

        private static ChartCase Finish(
            string name,
            bool replace,
            string type,
            int rows,
            int series,
            string lineAxis,
            string firstCat,
            double firstVal,
            double lastVal,
            string baseline,
            string html,
            AttrCheck[] attrs)
        {
            return new ChartCase
            {
                Name = name,
                IsReplace = replace,
                CreateHtml = replace ? baseline : html,
                ReplaceHtml = replace ? html : null,
                ExpectType = type,
                ExpectRows = rows,
                ExpectSeries = series,
                ExpectLineAxis = lineAxis,
                FirstCategory = firstCat,
                FirstValue = firstVal,
                LastValue = lastVal,
                Attrs = new List<AttrCheck>(attrs)
            };
        }

        private static Dictionary<string, string> Node(string key, string value)
        {
            return new Dictionary<string, string> { { key, value } };
        }

        private static KeyValuePair<string, string> Pair(string key, string value)
        {
            return new KeyValuePair<string, string>(key, value);
        }

        private static Dictionary<string, string> Th(params KeyValuePair<string, string>[] pairs)
        {
            var d = new Dictionary<string, string>();
            for (int i = 0; i < pairs.Length; i++)
            {
                d[pairs[i].Key] = pairs[i].Value;
            }

            return d;
        }

        private static Dictionary<string, string> Merge(params Dictionary<string, string>[] parts)
        {
            var d = new Dictionary<string, string>();
            for (int i = 0; i < parts.Length; i++)
            {
                if (parts[i] == null)
                {
                    continue;
                }

                foreach (KeyValuePair<string, string> kv in parts[i])
                {
                    d[kv.Key] = kv.Value;
                }
            }

            return d;
        }

        private static Dictionary<string, string> AxisPack(string prefix, string tickColor)
        {
            return Merge(
                Node(prefix + "-visible", "true"),
                Node(prefix + "-tick-font", "微软雅黑"),
                Node(prefix + "-tick-color", tickColor),
                Node(prefix + "-tick-size", "11"),
                Node(prefix + "-tick-position", "next"),
                Node(prefix + "-major-tick", "outside"),
                Node(prefix + "-minor-tick", "none"),
                prefix == "data-axis-x" ? null : Node(prefix + "-format", "0.00"),
                Node(prefix + "-grid", "true"),
                Node(prefix + "-grid-color", "#333333"),
                Node(prefix + "-line", "#000000"),
                Node(prefix + "-line-weight", "0.75"));
        }

        private static AttrCheck[] AxisExpects(string style, string tickColor)
        {
            var list = new List<AttrCheck>
            {
                AttrCheck.Exact(style + ".Visible", "true"),
                AttrCheck.Exact(style + ".TickFont", "微软雅黑"),
                AttrCheck.Hex(style + ".TickColor", tickColor),
                AttrCheck.Num(style + ".TickSize", "11"),
                AttrCheck.Exact(style + ".TickPosition", "next"),
                AttrCheck.Exact(style + ".MajorTick", "outside"),
                AttrCheck.Exact(style + ".MinorTick", "none"),
                AttrCheck.Exact(style + ".Grid", "true"),
                AttrCheck.Hex(style + ".GridColor", "#333333"),
                AttrCheck.Hex(style + ".Line", "#000000"),
                AttrCheck.Num(style + ".LineWeight", "0.75")
            };
            if (style != "AxisXStyle")
            {
                list.Add(AttrCheck.Exact(style + ".Format", "0.00"));
            }

            return list.ToArray();
        }
    }
}
