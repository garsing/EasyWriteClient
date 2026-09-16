using WordAddIn1.PresentationHost;

namespace PptChartRoundtripTest
{
    /// <summary>不启 PPT：约定 HTML → ApplyParser / ParseFormat / TryParseGrid。</summary>
    internal static class ChartParseTests
    {
        public static void Run(TestRun run)
        {
            ParseColumn(run);
            ParsePie3d(run);
            ParseComboY2(run);
            ParseComboBothY(run);
            ParseRejectsNonNumeric(run);
            ParseTypes(run);
        }

        private static void ParseColumn(TestRun run)
        {
            if (!ChartHtml.TryParseFirstChart(ChartHtml.ColumnCreate(), out PptHtmlApplyNode node, out string error))
            {
                run.Fail("parse-column", error);
                return;
            }

            run.ExpectEqual("parse-column 类型", "column", node.ChartType ?? node.ChartFormat?.ChartType);
            run.Expect(runName("parse-column 行数"), node.ChartGrid.Rows.Count == 4,
                "行数=" + node.ChartGrid.Rows.Count);
            run.ExpectClose("parse-column 2026", 2.15, node.ChartGrid.Rows[0][1]);
            PptHtmlChartColumn col = ChartHtml.ValueCol(node.ChartGrid, 0);
            run.ExpectEqual("parse-column 系列", "column", col?.SeriesType);
            run.ExpectEqual("parse-column 挂轴", "y", col?.AxisY);
        }

        private static void ParsePie3d(TestRun run)
        {
            if (!ChartHtml.TryParseFirstChart(ChartHtml.Pie3dCreate(), out PptHtmlApplyNode node, out string error))
            {
                run.Fail("parse-pie3d", error);
                return;
            }

            run.ExpectEqual("parse-pie3d 类型", "pie3d", node.ChartType ?? node.ChartFormat?.ChartType);
            run.Expect(runName("parse-pie3d 5 扇区"), node.ChartGrid.Rows.Count == 5,
                "行数=" + node.ChartGrid.Rows.Count);
            run.ExpectClose("parse-pie3d Bloomberg", 33, node.ChartGrid.Rows[0][1]);
            PptHtmlChartColumn col = ChartHtml.ValueCol(node.ChartGrid, 0);
            run.ExpectEqual("parse-pie3d 系列", "pie3d", col?.SeriesType);
            run.ExpectEqual("parse-pie3d 百分比标签", "true", col?.ShowPercentage);
        }

        private static void ParseComboY2(TestRun run)
        {
            if (!ChartHtml.TryParseFirstChart(ChartHtml.ComboY2Create(), out PptHtmlApplyNode node, out string error))
            {
                run.Fail("parse-combo-y2", error);
                return;
            }

            PptHtmlChartColumn col = ChartHtml.ValueCol(node.ChartGrid, 0);
            PptHtmlChartColumn line = ChartHtml.ValueCol(node.ChartGrid, 1);
            run.ExpectEqual("parse-combo-y2 柱系列", "column", col?.SeriesType);
            run.ExpectEqual("parse-combo-y2 柱轴", "y", col?.AxisY);
            run.ExpectEqual("parse-combo-y2 折线系列", "line", line?.SeriesType);
            run.ExpectEqual("parse-combo-y2 折线轴", "y2", line?.AxisY);
        }

        private static void ParseComboBothY(TestRun run)
        {
            // 会话 887：两列都写 data-axis=y，解析层不得偷偷改成 y2
            if (!ChartHtml.TryParseFirstChart(ChartHtml.ComboBothYCreate(), out PptHtmlApplyNode node, out string error))
            {
                run.Fail("parse-combo-both-y", error);
                return;
            }

            PptHtmlChartColumn line = ChartHtml.ValueCol(node.ChartGrid, 1);
            run.ExpectEqual("parse-combo-both-y 折线仍挂 y", "y", line?.AxisY);
        }

        private static void ParseRejectsNonNumeric(TestRun run)
        {
            string html = ChartHtml.Section(
                "  <div ShapeId=\"new-bad\" data-shape-type=\"chart\" data-chart-type=\"column\" style=\"position:absolute;left:10%;top:20%;width:40%;height:40%\">\n"
                + "    <table>\n"
                + "      <tr><th data-col=\"category\">类别</th><th data-col=\"value\">值</th></tr>\n"
                + "      <tr><td>Q1</td><td>不是数字</td></tr>\n"
                + "    </table>\n"
                + "  </div>");
            bool ok = ChartHtml.TryParseFirstChart(html, out _, out string error);
            run.Expect("parse-reject-non-numeric", !ok && !string.IsNullOrEmpty(error),
                ok ? "应失败却解析成功" : "error 为空");
        }

        private static void ParseTypes(TestRun run)
        {
            run.Expect("parse-type column",
                PptHtmlChartIo.TryParseType("column", out _, out string c1, out _) && c1 == "column",
                "column");
            run.Expect("parse-type pie",
                PptHtmlChartIo.TryParseType("pie", out _, out string c2, out _) && c2 == "pie2d",
                "pie→pie2d");
            run.Expect("parse-type pie3d",
                PptHtmlChartIo.TryParseType("pie3d", out _, out string c3, out _) && c3 == "pie3d",
                "pie3d");
            run.Expect("parse-type line",
                PptHtmlChartIo.TryParseType("line", out _, out string c4, out _) && c4 == "line",
                "line");
            run.Expect("parse-type 拒绝 combo",
                !PptHtmlChartIo.TryParseType("combo", out _, out _, out _),
                "combo 不应是可建底型");
        }

        private static string runName(string name)
        {
            return name;
        }
    }
}
