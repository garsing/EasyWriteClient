using System.IO;
using WordAddIn1.PresentationHost;

namespace PptHtmlContentRoundtripTest
{
    /// <summary>不启 PPT：解析契约与负例。</summary>
    internal static class ContentParseTests
    {
        public static void Run(TestRun run)
        {
            ParseTextbox(run);
            ParseTable(run);
            ParseTableMerge(run);
            ParseTableRejectBadMerge(run);
            ParseChartRejectTableEnhance(run);
            ParsePictureAbsolute(run);
            ParseRejectHttp(run);
            ParseLeanTextbox(run);
        }

        private static void ParseTextbox(TestRun run)
        {
            string html = ContentHtml.TextboxCreate(
                "标题",
                extraAttrs: "data-font-size=\"20pt\" data-font-color=\"#112233\" "
                    + "data-font-bold=\"true\" data-align=\"right\" data-valign=\"bottom\"");
            if (!ContentHtml.TryParse(html, "439", out PptHtmlApplyPlan plan, out string error))
            {
                run.Fail("parse-tb", error);
                return;
            }

            run.Expect("parse-tb 1 节点", plan.Nodes != null && plan.Nodes.Count == 1, "节点数");
            PptHtmlApplyNode n = plan.Nodes[0];
            run.Expect("parse-tb 新建", n.IsCreate, "应新建");
            run.ExpectEqual("parse-tb 类型", "textbox", n.ShapeType);
            run.ExpectEqual("parse-tb 正文", "标题", n.Text);
            run.Expect("parse-tb HasText", n.HasText, "HasText");
            run.Expect("parse-tb 字号", n.FontSizePt.HasValue && System.Math.Abs(n.FontSizePt.Value - 20) < 0.01,
                "字号=" + n.FontSizePt);
            run.ExpectEqual("parse-tb 字色", "#112233", n.FontColor);
            run.Expect("parse-tb 粗体", n.FontBold == true, "bold=" + n.FontBold);
            run.ExpectEqual("parse-tb align", "right", n.Align);
            run.ExpectEqual("parse-tb valign", "bottom", n.Valign);
            run.Expect("parse-tb 几何", n.HasGeometry && n.LeftPct.HasValue, "无几何");
        }

        private static void ParseTable(TestRun run)
        {
            string[][] cells =
            {
                new[] { "a", "b" },
                new[] { "c", "d" }
            };
            string html = ContentHtml.TableCreate(cells);
            if (!ContentHtml.TryParse(html, "1", out PptHtmlApplyPlan plan, out string error))
            {
                run.Fail("parse-tbl", error);
                return;
            }

            PptHtmlApplyNode n = plan.Nodes[0];
            run.ExpectEqual("parse-tbl 类型", "table", n.ShapeType);
            run.Expect("parse-tbl 2×2",
                n.TableCells != null && n.TableCells.Count == 2 && n.TableCells[0].Count == 2,
                "网格尺寸不对");
            run.ExpectEqual("parse-tbl a1", "a", n.TableCells[0][0]);
            run.ExpectEqual("parse-tbl d", "d", n.TableCells[1][1]);
            run.Expect("parse-tbl TableGrid", n.TableGrid != null && n.TableGrid.RowCount == 2, "无 TableGrid");
        }

        private static void ParseTableMerge(TestRun run)
        {
            string html = ContentHtml.Section(
                "  <table data-shape-type=\"table\" style=\""
                + ContentHtml.Geo(8, 35, 45, 30)
                + "\" data-table-style=\"three-line\" data-col-widths=\"50%,50%\">\n"
                + "    <tr><th colspan=\"2\">头</th></tr>\n"
                + "    <tr><td></td><td data-fill=\"#F5F5F5\"><span style=\"color:#C00000;font-weight:bold\">重点</span></td></tr>\n"
                + "  </table>");
            if (!ContentHtml.TryParse(html, "1", out PptHtmlApplyPlan plan, out string error))
            {
                run.Fail("parse-merge", error);
                return;
            }

            PptHtmlApplyNode n = plan.Nodes[0];
            run.Expect("parse-merge grid", n.TableGrid != null, "无网格");
            if (n.TableGrid == null)
            {
                return;
            }

            run.Expect("parse-merge rows", n.TableGrid.RowCount == 2, "rows=" + n.TableGrid.RowCount);
            run.Expect("parse-merge cols", n.TableGrid.ColCount == 2, "cols=" + n.TableGrid.ColCount);
            PptHtmlTableCell a0 = n.TableGrid.CellAt(0, 0);
            run.Expect("parse-merge colspan", a0 != null && !a0.IsCovered && a0.ColSpan == 2, "表头未合并");
            run.ExpectEqual("parse-merge 表头字", "头", a0.Text);
            run.Expect("parse-merge 占位", n.TableGrid.CellAt(0, 1) != null && n.TableGrid.CellAt(0, 1).IsCovered, "缺占位");
            PptHtmlTableCell b1 = n.TableGrid.CellAt(1, 1);
            run.Expect("parse-merge fill", b1 != null && b1.Fill == "#F5F5F5", "底色");
            run.Expect("parse-merge bold", b1 != null && b1.FontBold == true, "加粗");
            run.ExpectEqual("parse-merge 字色", "#C00000", b1.FontColor);
            run.Expect("parse-merge three-line",
                n.TableStyle != null && n.TableStyle.TableStyle == "three-line",
                "无三线");
            run.Expect("parse-merge col-widths",
                n.TableStyle != null && n.TableStyle.ColWidthPcts != null && n.TableStyle.ColWidthPcts.Length == 2,
                "无列宽");
        }

        private static void ParseTableRejectBadMerge(TestRun run)
        {
            // 第二行缺格 → 占位不合法
            string html = ContentHtml.Section(
                "  <table data-shape-type=\"table\" style=\""
                + ContentHtml.Geo(8, 35, 45, 30)
                + "\">\n"
                + "    <tr><td>a</td><td>b</td></tr>\n"
                + "    <tr><td>c</td></tr>\n"
                + "  </table>");
            bool ok = ContentHtml.TryParse(html, "1", out _, out string error);
            run.Expect("parse-bad-merge 应失败", !ok, ok ? "应失败" : ("error=" + error));
        }

        private static void ParseChartRejectTableEnhance(TestRun run)
        {
            string html = ContentHtml.Section(
                "  <div data-shape-type=\"chart\" data-chart-type=\"column\" style=\""
                + ContentHtml.Geo(10, 10, 40, 40)
                + "\">\n"
                + "    <table><tr><th colspan=\"2\">x</th></tr><tr><td></td><td>1</td></tr></table>\n"
                + "  </div>");
            bool ok = ContentHtml.TryParse(html, "1", out _, out string error);
            run.Expect(
                "parse-chart-enhance 应失败",
                !ok && error != null && error.IndexOf("chart", System.StringComparison.OrdinalIgnoreCase) >= 0,
                ok ? "应失败" : ("error=" + error));
        }

        private static void ParsePictureAbsolute(TestRun run)
        {
            string path = Path.Combine(Path.GetTempPath(), "ew-content-test.png");
            File.WriteAllBytes(path, ContentAssetsIo.MinimalPngRed);
            try
            {
                string html = ContentHtml.PictureCreate(path);
                if (!ContentHtml.TryParse(html, "2", out PptHtmlApplyPlan plan, out string error))
                {
                    run.Fail("parse-pic", error);
                    return;
                }

                PptHtmlApplyNode n = plan.Nodes[0];
                run.ExpectEqual("parse-pic 类型", "picture", n.ShapeType);
                run.Expect("parse-pic data-src", !string.IsNullOrEmpty(n.DataSrc), "data-src 空");
                ContentHtml.ResolvePicturePaths(plan);
                run.Expect("parse-pic ResolvedLocalPath",
                    !string.IsNullOrEmpty(n.ResolvedLocalPath) && File.Exists(n.ResolvedLocalPath),
                    "未解析到本地文件");
            }
            finally
            {
                try
                {
                    File.Delete(path);
                }
                catch
                {
                }
            }
        }

        private static void ParseRejectHttp(TestRun run)
        {
            // 解析本身不拦 http；分类阶段拒绝
            bool ok = PptHtmlFileSource.TryClassify(
                "https://example.com/x.png",
                out _,
                out _,
                out _,
                out string error);
            run.Expect(
                "parse-reject-http",
                !ok && error != null && error.IndexOf("http", System.StringComparison.OrdinalIgnoreCase) >= 0,
                ok ? "应失败却成功" : ("error=" + error));
        }

        private static void ParseLeanTextbox(TestRun run)
        {
            string html = ContentHtml.TextboxUpdate("sid439-s12", "瘦稿");
            if (!ContentHtml.TryParse(html, "439", out PptHtmlApplyPlan plan, out string error))
            {
                run.Fail("parse-lean-tb", error);
                return;
            }

            PptHtmlApplyNode n = plan.Nodes[0];
            run.Expect("parse-lean-tb 更新", !n.IsCreate, "应更新");
            run.ExpectEqual("parse-lean-tb ShapeId", "sid439-s12", n.ShapeId);
            run.Expect("parse-lean-tb 无几何", !n.HasGeometry, "瘦稿不应带几何");
            run.ExpectEqual("parse-lean-tb 正文", "瘦稿", n.Text);
        }
    }
}
