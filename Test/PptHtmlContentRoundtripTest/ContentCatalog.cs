using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using WordAddIn1.PresentationHost;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PptHtmlContentRoundtripTest
{
    internal static class ContentCatalog
    {
        public const int BatchCount = 1;

        public static int Total => Build().Count;

        public static List<ContentCase> Build()
        {
            var list = new List<ContentCase>();
            int page = 0;

            Add(list, ref page, "tb-create-text",
                assets => ContentHtml.TextboxCreate("你好 EasyWrite"),
                null,
                ctx =>
                {
                    PptHtmlShapeNode n = PreferCreated(ctx, "textbox");
                    if (n == null)
                    {
                        return "读回无 textbox";
                    }

                    if (!string.Equals(
                            ContentAssert.NormText(n.Text),
                            "你好 EasyWrite",
                            StringComparison.Ordinal))
                    {
                        return "正文期望 你好 EasyWrite 实际 " + n.Text;
                    }

                    return null;
                });

            Add(list, ref page, "tb-replace-text",
                assets => ContentHtml.TextboxCreate(
                    "原稿",
                    extraAttrs: "data-fill=\"#FFEECC\" data-font-size=\"18pt\""),
                (assets, ids) => ContentHtml.TextboxUpdate(ids[0], "改后正文"),
                ctx =>
                {
                    PptHtmlShapeNode n = ContentAssert.FindById(ctx.AfterReplace, ctx.CreatedShapeIds[0])
                        ?? ContentAssert.FindByType(ctx.AfterReplace, "textbox");
                    if (n == null)
                    {
                        return "改后无 textbox";
                    }

                    if (!string.Equals(ContentAssert.NormText(n.Text), "改后正文", StringComparison.Ordinal))
                    {
                        return "改后正文不一致: " + n.Text;
                    }

                    if (!string.Equals(
                            ContentAssert.HexOrNull(n.Fill),
                            "#FFEECC",
                            StringComparison.OrdinalIgnoreCase))
                    {
                        return "瘦稿改字不应丢 fill，实际 " + (n.Fill ?? "(null)");
                    }

                    if (!ContentAssert.FontSizeClose(n.FontSizePt, 18))
                    {
                        return "瘦稿改字不应丢字号，实际 " + (n.FontSizePt.HasValue
                            ? n.FontSizePt.Value.ToString(CultureInfo.InvariantCulture)
                            : "(null)");
                    }

                    return null;
                });

            Add(list, ref page, "tb-create-font",
                assets => ContentHtml.TextboxCreate(
                    "样式字",
                    extraAttrs: "data-font-size=\"24pt\" data-font-color=\"#C00000\" "
                        + "data-font-bold=\"true\" data-align=\"center\" data-valign=\"middle\""),
                null,
                ctx =>
                {
                    PptHtmlShapeNode n = PreferCreated(ctx, "textbox");
                    if (n == null)
                    {
                        return "无 textbox";
                    }

                    if (!ContentAssert.FontSizeClose(n.FontSizePt, 24))
                    {
                        return "字号期望 24 实际 " + n.FontSizePt;
                    }

                    if (!string.Equals(
                            ContentAssert.HexOrNull(n.FontColor),
                            "#C00000",
                            StringComparison.OrdinalIgnoreCase))
                    {
                        return "字色期望 #C00000 实际 " + (n.FontColor ?? "(null)");
                    }

                    if (n.FontBold != true)
                    {
                        return "加粗期望 true 实际 " + n.FontBold;
                    }

                    if (!string.Equals(n.Align, "center", StringComparison.OrdinalIgnoreCase))
                    {
                        return "align 期望 center 实际 " + (n.Align ?? "(null)");
                    }

                    if (!string.Equals(n.Valign, "middle", StringComparison.OrdinalIgnoreCase))
                    {
                        return "valign 期望 middle 实际 " + (n.Valign ?? "(null)");
                    }

                    return null;
                });

            string[][] gridA =
            {
                new[] { "A1", "B1" },
                new[] { "A2", "B2" }
            };
            string[][] gridB =
            {
                new[] { "X1", "Y1" },
                new[] { "X2", "Y2" }
            };

            Add(list, ref page, "tbl-create-grid",
                assets => ContentHtml.TableCreate(gridA),
                null,
                ctx => ContentAssert.TableCellsMismatch(
                    PreferCreated(ctx, "table"),
                    gridA));

            Add(list, ref page, "tbl-replace-grid",
                assets => ContentHtml.TableCreate(gridA),
                (assets, ids) => ContentHtml.TableUpdate(ids[0], gridB),
                ctx => ContentAssert.TableCellsMismatch(
                    ContentAssert.FindById(ctx.AfterReplace, ctx.CreatedShapeIds[0])
                        ?? ContentAssert.FindByType(ctx.AfterReplace, "table"),
                    gridB));

            Add(list, ref page, "pic-create",
                assets => ContentHtml.PictureCreate(assets.RedPng),
                null,
                ctx =>
                {
                    PptHtmlShapeNode n = PreferCreated(ctx, "picture");
                    if (n == null)
                    {
                        return "无 picture";
                    }

                    string err = ContentAssert.AttachPictureSrc(
                        (PowerPoint.Presentation)ctx.Presentation,
                        ctx.AfterCreate,
                        ctx.ExportDir);
                    if (err != null)
                    {
                        return "导出失败: " + err;
                    }

                    if (string.IsNullOrEmpty(n.DataSrc))
                    {
                        return "导出后无 data-src";
                    }

                    return null;
                });

            Add(list, ref page, "pic-replace-src-geo-keep",
                assets => ContentHtml.PictureCreate(assets.RedPng, 55, 20, 35, 40),
                (assets, ids) => ContentHtml.PictureUpdate(ids[0], assets.BluePng),
                ctx =>
                {
                    PptHtmlShapeNode before = PreferCreated(ctx, "picture", useReplace: false);
                    PptHtmlShapeNode after = PreferCreated(ctx, "picture", useReplace: true);
                    if (before == null || after == null)
                    {
                        return "换图前后缺 picture";
                    }

                    if (!ContentAssert.GeoClose(before.Style, after.Style))
                    {
                        return "换图应保几何: 前=" + before.Style + " 后=" + after.Style;
                    }

                    string err1 = ContentAssert.AttachPictureSrc(
                        (PowerPoint.Presentation)ctx.Presentation,
                        ctx.AfterCreate,
                        Path.Combine(ctx.ExportDir, "before"));
                    string err2 = ContentAssert.AttachPictureSrc(
                        (PowerPoint.Presentation)ctx.Presentation,
                        ctx.AfterReplace,
                        Path.Combine(ctx.ExportDir, "after"));
                    if (err1 != null || err2 != null)
                    {
                        return "导出失败: " + (err1 ?? err2);
                    }

                    if (string.IsNullOrEmpty(before.DataSrc) || string.IsNullOrEmpty(after.DataSrc))
                    {
                        return "导出 data-src 为空 before="
                            + (before.DataSrc ?? "(null)")
                            + " after="
                            + (after.DataSrc ?? "(null)");
                    }

                    if (string.Equals(before.DataSrc, after.DataSrc, StringComparison.OrdinalIgnoreCase))
                    {
                        return "换图后 data-src 未变: " + after.DataSrc;
                    }

                    return null;
                });

            Add(list, ref page, "mix-create-three",
                assets => ContentHtml.MixCreate(
                    "混排标题",
                    new[]
                    {
                        new[] { "品类", "销量" },
                        new[] { "甲", "12" },
                        new[] { "乙", "9" }
                    },
                    assets.RedPng),
                null,
                ctx =>
                {
                    PptHtmlShapeNode tb = PreferCreated(ctx, "textbox");
                    if (tb == null)
                    {
                        return "缺 textbox";
                    }

                    string tbl = ContentAssert.TableCellsMismatch(
                        PreferCreated(ctx, "table"),
                        new[]
                        {
                            new[] { "品类", "销量" },
                            new[] { "甲", "12" },
                            new[] { "乙", "9" }
                        });
                    if (tbl != null)
                    {
                        return tbl;
                    }

                    if (PreferCreated(ctx, "picture") == null)
                    {
                        return "缺 picture";
                    }

                    if (!string.Equals(ContentAssert.NormText(tb.Text), "混排标题", StringComparison.Ordinal))
                    {
                        return "标题不一致: " + tb.Text;
                    }

                    return null;
                });

            Add(list, ref page, "pic-neg-http",
                assets => ContentHtml.Section(
                    "  <img data-shape-type=\"picture\" data-src=\"https://example.com/a.png\" "
                    + "style=\"position:absolute;left:10%;top:10%;width:20%;height:20%\" />"),
                null,
                null,
                expectApplyErrorContains: "不存在");

            return list;
        }

        private static void Add(
            List<ContentCase> list,
            ref int page,
            string name,
            Func<ContentAssets, string> create,
            Func<ContentAssets, IList<string>, string> replace,
            Func<ContentAssertContext, string> match,
            string expectApplyErrorContains = null)
        {
            page++;
            list.Add(new ContentCase
            {
                Index = list.Count + 1,
                Batch = 1,
                Page = page,
                Name = name,
                CreateHtml = create,
                ReplaceHtml = replace,
                Match = match,
                ExpectApplyErrorContains = expectApplyErrorContains
            });
        }

        private static PptHtmlShapeNode PreferCreated(
            ContentAssertContext ctx,
            string shapeType,
            bool useReplace = false)
        {
            PptHtmlReadResult result = useReplace ? ctx.AfterReplace : ctx.AfterCreate;
            if (ctx.CreatedShapeIds != null)
            {
                for (int i = 0; i < ctx.CreatedShapeIds.Count; i++)
                {
                    PptHtmlShapeNode n = ContentAssert.FindById(result, ctx.CreatedShapeIds[i]);
                    if (n != null
                        && string.Equals(n.ShapeType, shapeType, StringComparison.OrdinalIgnoreCase))
                    {
                        return n;
                    }
                }
            }

            // 跳过页眉用例标签（字号约 9、贴顶）
            if (result?.Shapes == null)
            {
                return null;
            }

            for (int i = 0; i < result.Shapes.Count; i++)
            {
                PptHtmlShapeNode n = result.Shapes[i];
                if (n == null
                    || !string.Equals(n.ShapeType, shapeType, StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                if (string.Equals(shapeType, "textbox", StringComparison.OrdinalIgnoreCase)
                    && !string.IsNullOrEmpty(n.Text)
                    && n.Text.IndexOf("B1P", StringComparison.Ordinal) >= 0)
                {
                    continue;
                }

                return n;
            }

            return null;
        }
    }
}
