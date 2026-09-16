using System.Text;
using WordAddIn1.PresentationHost;

namespace PptChartRoundtripTest
{
    /// <summary>不启 PPT：全属性、交叉、新建/换数。</summary>
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
            KitchenSinkCreate(run);
            KitchenSinkReplace(run);
            KitchenSinkRoundtrip(run);
            CrossBar(run);
            CrossLine(run);
            CrossPie2d(run);
            CrossTwoColumn(run);
            CrossBarLineY2CreateAndReplace(run);
            LeanReplaceOmitsChrome(run);
            LegacyAxisSlots(run);
            CreateVsReplaceGeometry(run);
            Rejects(run);
        }

        private static void ParseColumn(TestRun run)
        {
            if (!ChartAssert.TryParse(run, "parse-column", ChartHtml.ColumnCreate(), out PptHtmlApplyNode node))
            {
                return;
            }

            run.ExpectEqual("parse-column 类型", "column", node.ChartType ?? node.ChartFormat?.ChartType);
            run.Expect("parse-column 行数", node.ChartGrid.Rows.Count == 4, "行数=" + node.ChartGrid.Rows.Count);
            run.ExpectClose("parse-column 2026", 2.15, node.ChartGrid.Rows[0][1]);
            PptHtmlChartColumn col = ChartHtml.ValueCol(node.ChartGrid, 0);
            run.ExpectEqual("parse-column 系列", "column", col?.SeriesType);
            run.ExpectEqual("parse-column 挂轴", "y", col?.AxisY);
            run.Expect("parse-column 新建", node.IsCreate, "应走新建");
        }

        private static void ParsePie3d(TestRun run)
        {
            if (!ChartAssert.TryParse(run, "parse-pie3d", ChartHtml.Pie3dCreate(), out PptHtmlApplyNode node))
            {
                return;
            }

            run.ExpectEqual("parse-pie3d 类型", "pie3d", node.ChartType ?? node.ChartFormat?.ChartType);
            run.Expect("parse-pie3d 5 扇区", node.ChartGrid.Rows.Count == 5, "行数=" + node.ChartGrid.Rows.Count);
            run.ExpectClose("parse-pie3d Bloomberg", 33, node.ChartGrid.Rows[0][1]);
            PptHtmlChartColumn col = ChartHtml.ValueCol(node.ChartGrid, 0);
            run.ExpectEqual("parse-pie3d 系列", "pie3d", col?.SeriesType);
            run.ExpectEqual("parse-pie3d 百分比标签", "true", col?.ShowPercentage);
        }

        private static void ParseComboY2(TestRun run)
        {
            if (!ChartAssert.TryParse(run, "parse-combo-y2", ChartHtml.ComboY2Create(), out PptHtmlApplyNode node))
            {
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
            if (!ChartAssert.TryParse(run, "parse-combo-both-y", ChartHtml.ComboBothYCreate(), out PptHtmlApplyNode node))
            {
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
            run.Expect("parse-type column_clustered",
                PptHtmlChartIo.TryParseType("column_clustered", out _, out string c1b, out _) && c1b == "column",
                "column_clustered→column");
            run.Expect("parse-type bar",
                PptHtmlChartIo.TryParseType("bar", out _, out string c5, out _) && c5 == "bar",
                "bar");
            run.Expect("parse-type pie",
                PptHtmlChartIo.TryParseType("pie", out _, out string c2, out _) && c2 == "pie2d",
                "pie→pie2d");
            run.Expect("parse-type pie2d",
                PptHtmlChartIo.TryParseType("pie2d", out _, out string c2b, out _) && c2b == "pie2d",
                "pie2d");
            run.Expect("parse-type pie3d",
                PptHtmlChartIo.TryParseType("pie3d", out _, out string c3, out _) && c3 == "pie3d",
                "pie3d");
            run.Expect("parse-type line",
                PptHtmlChartIo.TryParseType("line", out _, out string c4, out _) && c4 == "line",
                "line");
            run.Expect("parse-type 拒绝 combo",
                !PptHtmlChartIo.TryParseType("combo", out _, out _, out _),
                "combo 不应是可建底型");
            run.Expect("parse-type 拒绝 xlCombination",
                !PptHtmlChartIo.TryParseType("xlCombination", out _, out _, out _),
                "组合码不可建");
        }

        private static void KitchenSinkCreate(TestRun run)
        {
            if (!ChartAssert.TryParse(run, "sink-create", ChartHtml.KitchenSinkCreate(), out PptHtmlApplyNode node))
            {
                return;
            }

            run.Expect("sink-create 新建", node.IsCreate, "应走新建");
            AssertKitchenSink(run, "sink-create", node);
        }

        private static void KitchenSinkReplace(TestRun run)
        {
            if (!ChartAssert.TryParse(run, "sink-replace", ChartHtml.KitchenSinkReplace(), out PptHtmlApplyNode node))
            {
                return;
            }

            run.Expect("sink-replace 改已有", !node.IsCreate, "应走换数");
            run.ExpectEqual("sink-replace ShapeId", "sid256-s12", node.ShapeId);
            AssertKitchenSink(run, "sink-replace", node);
        }

        private static void AssertKitchenSink(TestRun run, string prefix, PptHtmlApplyNode node)
        {
            ChartAssert.FormatFull(run, prefix, node.ChartFormat, "column");
            ChartAssert.AxisFull(run, prefix + " x", node.ChartFormat.AxisXStyle, "General");
            ChartAssert.AxisFull(run, prefix + " y", node.ChartFormat.AxisYStyle, "0.00");
            ChartAssert.AxisFull(run, prefix + " y2", node.ChartFormat.AxisY2Style, "0.0%");
            ChartAssert.SeriesFull(run, prefix + " 柱", ChartHtml.ValueCol(node.ChartGrid, 0), "column", "y");
            ChartAssert.LineSeriesFull(run, prefix + " 折", ChartHtml.ValueCol(node.ChartGrid, 1));
            run.Expect("sink 4×3 表", node.ChartGrid.Rows.Count == 4 && node.ChartGrid.Rows[0].Count == 3,
                "rows=" + node.ChartGrid.Rows.Count);
            run.ExpectClose(prefix + " 线 0.315", 0.315, node.ChartGrid.Rows[3][2]);
        }

        private static void KitchenSinkRoundtrip(TestRun run)
        {
            if (!ChartAssert.TryParse(run, "sink-roundtrip", ChartHtml.KitchenSinkCreate(), out PptHtmlApplyNode node))
            {
                return;
            }

            var sb = new StringBuilder();
            sb.Append("<section style=\"position:relative;width:100%;height:100%\">");
            sb.Append("<div ShapeId=\"new-rt\" data-shape-type=\"chart\" style=\"position:absolute;left:10%;top:20%;width:40%;height:40%\"");
            PptHtmlChartIo.AppendFormatAttrs(sb, node.ChartFormat);
            sb.Append("><table>");
            sb.Append(PptHtmlChartIo.BuildInnerHtml(node.ChartGrid));
            sb.Append("</table></div></section>");

            if (!ChartAssert.TryParse(run, "sink-roundtrip 再解析", sb.ToString(), out PptHtmlApplyNode again))
            {
                return;
            }

            ChartAssert.FormatFull(run, "sink-roundtrip", again.ChartFormat, "column");
            ChartAssert.SeriesFull(run, "sink-roundtrip 柱", ChartHtml.ValueCol(again.ChartGrid, 0), "column", "y");
            ChartAssert.LineSeriesFull(run, "sink-roundtrip 折", ChartHtml.ValueCol(again.ChartGrid, 1));
        }

        private static void CrossBar(TestRun run)
        {
            if (!ChartAssert.TryParse(run, "cross-bar", ChartHtml.BarCreate(), out PptHtmlApplyNode node))
            {
                return;
            }

            run.ExpectEqual("cross-bar 类型", "bar", node.ChartType);
            run.ExpectEqual("cross-bar 系列", "bar", ChartHtml.ValueCol(node.ChartGrid, 0)?.SeriesType);
            run.Expect("cross-bar 新建", node.IsCreate, "应新建");
        }

        private static void CrossLine(TestRun run)
        {
            if (!ChartAssert.TryParse(run, "cross-line", ChartHtml.LineCreate(), out PptHtmlApplyNode node))
            {
                return;
            }

            run.ExpectEqual("cross-line 类型", "line", node.ChartType);
            run.ExpectEqual("cross-line 图例", "right", node.ChartFormat.Legend);
            run.ExpectEqual("cross-line 图级标记", "true", node.ChartFormat.DataMarkers);
            run.ExpectEqual("cross-line 图级标记大小", "7", node.ChartFormat.MarkerSize);
            run.ExpectEqual("cross-line 图级线粗", "2.5", node.ChartFormat.ChartLineWeight);
            PptHtmlChartColumn col = ChartHtml.ValueCol(node.ChartGrid, 0);
            run.ExpectEqual("cross-line 系列", "line", col?.SeriesType);
            run.ExpectEqual("cross-line 线色", "#00B0F0", col?.Line);
            run.ExpectEqual("cross-line 标记", "diamond", col?.Marker);
        }

        private static void CrossPie2d(TestRun run)
        {
            if (!ChartAssert.TryParse(run, "cross-pie2d", ChartHtml.Pie2dCreate(), out PptHtmlApplyNode node))
            {
                return;
            }

            run.ExpectEqual("cross-pie2d 类型", "pie2d", node.ChartType);
            run.ExpectEqual("cross-pie2d 爆炸", "12", node.ChartFormat.Explosion);
            run.ExpectEqual("cross-pie2d 图级百分比", "true", node.ChartFormat.ShowPercentage);
            run.ExpectEqual("cross-pie2d 图级数值关", "false", node.ChartFormat.ShowValue);
            PptHtmlChartColumn col = ChartHtml.ValueCol(node.ChartGrid, 0);
            run.ExpectEqual("cross-pie2d 系列", "pie2d", col?.SeriesType);
            run.ExpectEqual("cross-pie2d 标签位置", "outside", col?.LabelPosition);
            run.Expect("cross-pie2d 不写系列色", string.IsNullOrEmpty(col?.Color), "饼图扇区不应写 data-color");
        }

        private static void CrossTwoColumn(TestRun run)
        {
            if (!ChartAssert.TryParse(run, "cross-2col", ChartHtml.TwoColumnClusteredCreate(), out PptHtmlApplyNode node))
            {
                return;
            }

            run.ExpectEqual("cross-2col gap", "80", node.ChartFormat.GapWidth);
            run.ExpectEqual("cross-2col overlap", "-30", node.ChartFormat.Overlap);
            run.ExpectEqual("cross-2col 图例", "top", node.ChartFormat.Legend);
            PptHtmlChartColumn a = ChartHtml.ValueCol(node.ChartGrid, 0);
            PptHtmlChartColumn b = ChartHtml.ValueCol(node.ChartGrid, 1);
            run.ExpectEqual("cross-2col A 色", "#2497EA", a?.Color);
            run.ExpectEqual("cross-2col B 色", "#EE822F", b?.Color);
            run.ExpectEqual("cross-2col 两列都挂 y", "y", b?.AxisY);
            run.ExpectClose("cross-2col 进口 0.7", 0.7, node.ChartGrid.Rows[3][2]);
        }

        private static void CrossBarLineY2CreateAndReplace(TestRun run)
        {
            if (!ChartAssert.TryParse(run, "cross-barline-create", ChartHtml.BarLineY2Create(), out PptHtmlApplyNode created))
            {
                return;
            }

            run.Expect("cross-barline-create 新建", created.IsCreate, "应新建");
            run.ExpectEqual("cross-barline-create 底型", "bar", created.ChartType);
            run.ExpectEqual("cross-barline-create 条系列", "bar", ChartHtml.ValueCol(created.ChartGrid, 0)?.SeriesType);
            run.ExpectEqual("cross-barline-create 折线 y2", "y2", ChartHtml.ValueCol(created.ChartGrid, 1)?.AxisY);
            run.ExpectEqual("cross-barline-create y2-visible", "true", created.ChartFormat.AxisY2Style?.Visible);

            if (!ChartAssert.TryParse(run, "cross-barline-replace", ChartHtml.BarLineY2Replace(), out PptHtmlApplyNode replaced))
            {
                return;
            }

            run.Expect("cross-barline-replace 改已有", !replaced.IsCreate, "应换数");
            run.ExpectEqual("cross-barline-replace 底型", "bar", replaced.ChartType ?? replaced.ChartFormat.ChartType);
            run.ExpectEqual("cross-barline-replace 折线 y2", "y2", ChartHtml.ValueCol(replaced.ChartGrid, 1)?.AxisY);
        }

        private static void LeanReplaceOmitsChrome(TestRun run)
        {
            if (!ChartAssert.TryParse(run, "lean-replace", ChartHtml.LeanReplace(), out PptHtmlApplyNode node))
            {
                return;
            }

            run.Expect("lean-replace 改已有", !node.IsCreate, "应换数");
            run.Expect("lean-replace 不写类型", string.IsNullOrEmpty(node.ChartFormat.ChartType),
                "瘦稿未写 type=" + (node.ChartFormat.ChartType ?? ""));
            run.Expect("lean-replace 不写图例", string.IsNullOrEmpty(node.ChartFormat.Legend),
                "Legend=" + (node.ChartFormat.Legend ?? ""));
            run.Expect("lean-replace 标题未提及", !node.ChartFormat.TitleMentioned, "瘦稿不应钉标题");
            run.Expect("lean-replace 图例未提及", !node.ChartFormat.LegendMentioned, "瘦稿不应钉图例");
            run.ExpectClose("lean-replace 2029", 4.8, node.ChartGrid.Rows[3][1]);
        }

        private static void LegacyAxisSlots(TestRun run)
        {
            if (!ChartAssert.TryParse(run, "legacy-axis", ChartHtml.LegacyAxisCreate(), out PptHtmlApplyNode node))
            {
                return;
            }

            run.ExpectEqual("legacy data-axis-y=primary", "y", ChartHtml.ValueCol(node.ChartGrid, 0)?.AxisY);
            run.ExpectEqual("legacy data-axis-y2=true", "y2", ChartHtml.ValueCol(node.ChartGrid, 1)?.AxisY);

            if (!ChartAssert.TryParse(run, "legacy-axis-false", ChartHtml.LegacyAxisFalseCreate(), out PptHtmlApplyNode n2))
            {
                return;
            }

            run.ExpectEqual("legacy data-axis-y2=false → y", "y", ChartHtml.ValueCol(n2.ChartGrid, 0)?.AxisY);

            if (!ChartAssert.TryParse(run, "legacy-secondary", ChartHtml.ComboSecondaryCreate(), out PptHtmlApplyNode n3))
            {
                return;
            }

            run.ExpectEqual("legacy data-axis=secondary → y2", "y2", ChartHtml.ValueCol(n3.ChartGrid, 1)?.AxisY);
        }

        private static void CreateVsReplaceGeometry(TestRun run)
        {
            bool createNoStyle = ChartHtml.TryParseFirstChart(ChartHtml.CreateMissingStyle(), out _, out string createErr);
            run.Expect("create 无 style 失败", !createNoStyle && !string.IsNullOrEmpty(createErr),
                createNoStyle ? "应失败" : createErr);

            bool replaceOk = ChartHtml.TryParseFirstChart(ChartHtml.LeanReplace(), out PptHtmlApplyNode node, out string replaceErr);
            run.Expect("replace 无 style 成功", replaceOk && node != null && !node.IsCreate,
                replaceErr ?? "解析失败");
        }

        private static void Rejects(TestRun run)
        {
            string pctCell = ChartHtml.Section(
                "  <div ShapeId=\"new-pct\" data-shape-type=\"chart\" data-chart-type=\"pie2d\" style=\"position:absolute;left:10%;top:20%;width:40%;height:40%\">\n"
                + "    <table>\n"
                + "      <tr><th data-col=\"category\"> </th><th data-col=\"value\">份额</th></tr>\n"
                + "      <tr><td>A</td><td>20%</td></tr>\n"
                + "    </table>\n"
                + "  </div>");
            run.Expect("reject 单元格写 20%",
                !ChartHtml.TryParseFirstChart(pctCell, out _, out _),
                "百分比应走属性，不应写在格子里");

            string badCol = ChartHtml.Section(
                "  <div ShapeId=\"new-role\" data-shape-type=\"chart\" data-chart-type=\"column\" style=\"position:absolute;left:10%;top:20%;width:40%;height:40%\">\n"
                + "    <table>\n"
                + "      <tr><th data-col=\"label\">类别</th><th data-col=\"value\">值</th></tr>\n"
                + "      <tr><td>A</td><td>1</td></tr>\n"
                + "    </table>\n"
                + "  </div>");
            run.Expect("reject data-col=label",
                !ChartHtml.TryParseFirstChart(badCol, out _, out _),
                "data-col 只能 category/value");

            string twoTables = ChartHtml.Section(
                "  <div ShapeId=\"new-2t\" data-shape-type=\"chart\" data-chart-type=\"column\" style=\"position:absolute;left:10%;top:20%;width:40%;height:40%\">\n"
                + "    <table><tr><th data-col=\"category\">c</th><th data-col=\"value\">v</th></tr><tr><td>A</td><td>1</td></tr></table>\n"
                + "    <table><tr><th data-col=\"category\">c</th><th data-col=\"value\">v</th></tr><tr><td>B</td><td>2</td></tr></table>\n"
                + "  </div>");
            run.Expect("reject 两张表",
                !ChartHtml.TryParseFirstChart(twoTables, out _, out _),
                "chart 只能一张内嵌表");
        }

    }
}
