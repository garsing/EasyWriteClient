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
