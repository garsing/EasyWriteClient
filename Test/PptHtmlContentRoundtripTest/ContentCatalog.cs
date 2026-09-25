using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using WordAddIn1.PresentationHost;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PptHtmlContentRoundtripTest
{
    /// <summary>P0/P1 内容往返目录。含表格/图片/文本框矩阵（tbl-mx / pic-mx / tb-mx）。</summary>
    internal static partial class ContentCatalog
    {
        public const int BatchCount = 1;

        public static int Total => Build().Count;

        public static List<ContentCase> Build()
        {
            var list = new List<ContentCase>();
            int page = 0;

            AddTextboxCases(list, ref page);
            AddTextboxMatrixCases(list, ref page);
            AddTableCases(list, ref page);
            AddPictureCases(list, ref page);
            AddPictureMatrixCases(list, ref page);
            AddMixCases(list, ref page);
            AddGroupCases(list, ref page);
            AddMixChartCases(list, ref page);
            AddOoxmlCases(list, ref page);

            return list;
        }

        private static void AddTextboxCases(List<ContentCase> list, ref int page)
        {
            Add(list, ref page, "tb-create-text",
                a => ContentHtml.TextboxCreate("你好 EasyWrite"),
                null,
                ctx => TextEq(PreferCreated(ctx, "textbox"), "你好 EasyWrite"));

            Add(list, ref page, "tb-replace-text",
                a => ContentHtml.TextboxCreate("原稿", extraAttrs: "data-fill=\"#FFEECC\" data-font-size=\"18pt\""),
                (a, ids) => ContentHtml.TextboxUpdate(ids[0], "改后正文"),
                ctx =>
                {
                    PptHtmlShapeNode n = PreferCreated(ctx, "textbox", useReplace: true);
                    string e = TextEq(n, "改后正文");
                    if (e != null) return e;
                    if (!HexEq(n.Fill, "#FFEECC")) return "瘦稿改字丢 fill: " + n.Fill;
                    if (!ContentAssert.FontSizeClose(n.FontSizePt, 18)) return "瘦稿改字丢字号: " + n.FontSizePt;
                    return null;
                });

            Add(list, ref page, "tb-create-font-size",
                a => ContentHtml.TextboxCreate("字号", extraAttrs: "data-font-size=\"28pt\""),
                null,
                ctx =>
                {
                    PptHtmlShapeNode n = PreferCreated(ctx, "textbox");
                    return n == null ? "无 textbox"
                        : ContentAssert.FontSizeClose(n.FontSizePt, 28) ? null : "字号期望 28 实际 " + n.FontSizePt;
                });

            Add(list, ref page, "tb-create-font-color",
                a => ContentHtml.TextboxCreate("字色", extraAttrs: "data-font-color=\"#0070C0\""),
                null,
                ctx =>
                {
                    PptHtmlShapeNode n = PreferCreated(ctx, "textbox");
                    return n == null ? "无 textbox"
                        : HexEq(n.FontColor, "#0070C0") ? null : "字色期望 #0070C0 实际 " + n.FontColor;
                });

            Add(list, ref page, "tb-create-font-bold-on",
                a => ContentHtml.TextboxCreate("粗体开", extraAttrs: "data-font-bold=\"true\""),
                null,
                ctx =>
                {
                    PptHtmlShapeNode n = PreferCreated(ctx, "textbox");
                    return n == null ? "无 textbox" : n.FontBold == true ? null : "加粗期望 true";
                });

            Add(list, ref page, "tb-create-font-bold-off",
                a => ContentHtml.TextboxCreate("粗体关", extraAttrs: "data-font-bold=\"false\""),
                null,
                ctx =>
                {
                    PptHtmlShapeNode n = PreferCreated(ctx, "textbox");
                    return n == null ? "无 textbox" : n.FontBold == false ? null : "加粗期望 false 实际 " + n.FontBold;
                });

            Add(list, ref page, "tb-create-font-name",
                a => ContentHtml.TextboxCreate("雅黑", extraAttrs: "data-font-name=\"微软雅黑\""),
                null,
                ctx =>
                {
                    PptHtmlShapeNode n = PreferCreated(ctx, "textbox");
                    if (n == null) return "无 textbox";
                    if (string.IsNullOrEmpty(n.FontName)) return "字体名为空";
                    if (n.FontName.IndexOf("雅黑", StringComparison.Ordinal) < 0
                        && n.FontName.IndexOf("YaHei", StringComparison.OrdinalIgnoreCase) < 0)
                    {
                        return "字体期望含雅黑 实际 " + n.FontName;
                    }

                    return null;
                });

            Add(list, ref page, "tb-create-align-left",
                a => ContentHtml.TextboxCreate("左齐", extraAttrs: "data-align=\"left\""),
                null,
                ctx => AlignEq(PreferCreated(ctx, "textbox"), "left"));

            Add(list, ref page, "tb-create-align-center",
                a => ContentHtml.TextboxCreate("居中", extraAttrs: "data-align=\"center\""),
                null,
                ctx => AlignEq(PreferCreated(ctx, "textbox"), "center"));

            Add(list, ref page, "tb-create-align-right",
                a => ContentHtml.TextboxCreate("右齐", extraAttrs: "data-align=\"right\""),
                null,
                ctx => AlignEq(PreferCreated(ctx, "textbox"), "right"));

            Add(list, ref page, "tb-create-align-justify",
                a => ContentHtml.TextboxCreate("两端对齐文字两端对齐文字两端对齐", extraAttrs: "data-align=\"justify\""),
                null,
                ctx => AlignEq(PreferCreated(ctx, "textbox"), "justify"));

            Add(list, ref page, "tb-create-valign-top",
                a => ContentHtml.TextboxCreate("顶齐", height: 30, extraAttrs: "data-valign=\"top\""),
                null,
                ctx => ValignEq(PreferCreated(ctx, "textbox"), "top"));

            Add(list, ref page, "tb-create-valign-middle",
                a => ContentHtml.TextboxCreate("中齐", height: 30, extraAttrs: "data-valign=\"middle\""),
                null,
                ctx => ValignEq(PreferCreated(ctx, "textbox"), "middle"));

            Add(list, ref page, "tb-create-valign-bottom",
                a => ContentHtml.TextboxCreate("底齐", height: 30, extraAttrs: "data-valign=\"bottom\""),
                null,
                ctx => ValignEq(PreferCreated(ctx, "textbox"), "bottom"));

            Add(list, ref page, "tb-create-fill",
                a => ContentHtml.TextboxCreate("填色", extraAttrs: "data-fill=\"#FFF2CC\""),
                null,
                ctx =>
                {
                    PptHtmlShapeNode n = PreferCreated(ctx, "textbox");
                    return n == null ? "无 textbox"
                        : HexEq(n.Fill, "#FFF2CC") ? null : "fill 期望 #FFF2CC 实际 " + n.Fill;
                });

            Add(list, ref page, "tb-create-fill-none",
                a => ContentHtml.TextboxCreate("无填充", extraAttrs: "data-fill=\"none\""),
                null,
                ctx =>
                {
                    PptHtmlShapeNode n = PreferCreated(ctx, "textbox");
                    if (n == null) return "无 textbox";
                    string fill = (n.Fill ?? "").Trim();
                    return string.Equals(fill, "none", StringComparison.OrdinalIgnoreCase)
                        || string.IsNullOrEmpty(fill)
                        ? null
                        : "fill 期望 none 实际 " + n.Fill;
                });

            Add(list, ref page, "tb-create-line",
                a => ContentHtml.TextboxCreate(
                    "描边",
                    extraAttrs: "data-line-color=\"#C00000\" data-line-width=\"2pt\""),
                null,
                ctx =>
                {
                    PptHtmlShapeNode n = PreferCreated(ctx, "textbox");
                    if (n == null) return "无 textbox";
                    if (!HexEq(n.LineColor, "#C00000")) return "线色期望 #C00000 实际 " + n.LineColor;
                    if (!n.LineWidthPt.HasValue || Math.Abs(n.LineWidthPt.Value - 2) > 0.6)
                    {
                        return "线宽期望 2 实际 " + n.LineWidthPt;
                    }

                    return null;
                });

            Add(list, ref page, "tb-create-margin",
                a => ContentHtml.TextboxCreate(
                    "边距",
                    extraAttrs: "data-margin-left=\"8pt\" data-margin-right=\"6pt\" "
                        + "data-margin-top=\"4pt\" data-margin-bottom=\"5pt\""),
                null,
                ctx =>
                {
                    PptHtmlShapeNode n = PreferCreated(ctx, "textbox");
                    if (n == null) return "无 textbox";
                    if (!Near(n.MarginLeftPt, 8)) return "margin-left 期望 8 实际 " + n.MarginLeftPt;
                    if (!Near(n.MarginRightPt, 6)) return "margin-right 期望 6 实际 " + n.MarginRightPt;
                    if (!Near(n.MarginTopPt, 4)) return "margin-top 期望 4 实际 " + n.MarginTopPt;
                    if (!Near(n.MarginBottomPt, 5)) return "margin-bottom 期望 5 实际 " + n.MarginBottomPt;
                    return null;
                });

            Add(list, ref page, "tb-create-para-space",
                a => ContentHtml.TextboxCreate(
                    "段距",
                    extraAttrs: "data-space-before=\"10pt\" data-space-after=\"8pt\" data-line-spacing=\"1.5\""),
                null,
                ctx =>
                {
                    PptHtmlShapeNode n = PreferCreated(ctx, "textbox");
                    if (n == null) return "无 textbox";
                    if (!Near(n.SpaceBeforePt, 10)) return "space-before 期望 10 实际 " + n.SpaceBeforePt;
                    if (!Near(n.SpaceAfterPt, 8)) return "space-after 期望 8 实际 " + n.SpaceAfterPt;
                    if (string.IsNullOrEmpty(n.LineSpacing)
                        || n.LineSpacing.IndexOf("1.5", StringComparison.Ordinal) < 0)
                    {
                        return "line-spacing 期望含 1.5 实际 " + n.LineSpacing;
                    }

                    return null;
                });

            Add(list, ref page, "tb-create-bullet",
                a => ContentHtml.TextboxCreate("条目一", extraAttrs: "data-bullet=\"bullet\""),
                null,
                ctx =>
                {
                    PptHtmlShapeNode n = PreferCreated(ctx, "textbox");
                    if (n == null) return "无 textbox";
                    return string.Equals(n.Bullet, "bullet", StringComparison.OrdinalIgnoreCase)
                        ? null
                        : "bullet 期望 bullet 实际 " + n.Bullet;
                });

            Add(list, ref page, "tb-create-bullet-number",
                a => ContentHtml.TextboxCreate("编号一", extraAttrs: "data-bullet=\"number\""),
                null,
                ctx =>
                {
                    PptHtmlShapeNode n = PreferCreated(ctx, "textbox");
                    if (n == null) return "无 textbox";
                    return string.Equals(n.Bullet, "number", StringComparison.OrdinalIgnoreCase)
                        ? null
                        : "bullet 期望 number 实际 " + n.Bullet;
                });

            Add(list, ref page, "tb-create-multi-para",
                a => ContentHtml.TextboxCreate("第一段\n第二段\n第三段"),
                null,
                ctx =>
                {
                    PptHtmlShapeNode n = PreferCreated(ctx, "textbox");
                    if (n == null) return "无 textbox";
                    string t = ContentAssert.NormText(n.Text);
                    if (t.IndexOf("第一段", StringComparison.Ordinal) < 0
                        || t.IndexOf("第二段", StringComparison.Ordinal) < 0
                        || t.IndexOf("第三段", StringComparison.Ordinal) < 0)
                    {
                        return "多段文案不全: " + t;
                    }

                    return null;
                });

            Add(list, ref page, "tb-create-geo",
                a => ContentHtml.TextboxCreate("几何", left: 12, top: 15, width: 33, height: 22),
                null,
                ctx =>
                {
                    PptHtmlShapeNode n = PreferCreated(ctx, "textbox");
                    if (n == null) return "无 textbox";
                    if (!ContentAssert.TryParseGeo(n.Style, out double l, out double t, out double w, out double h))
                    {
                        return "无几何 style";
                    }

                    if (Math.Abs(l - 12) > 0.8 || Math.Abs(t - 15) > 0.8
                        || Math.Abs(w - 33) > 0.8 || Math.Abs(h - 22) > 0.8)
                    {
                        return "几何偏离: " + n.Style;
                    }

                    return null;
                });

            Add(list, ref page, "tb-create-z",
                a => ContentHtml.TextboxCreate("叠放", extraAttrs: "data-z=\"5\""),
                null,
                ctx =>
                {
                    PptHtmlShapeNode n = PreferCreated(ctx, "textbox");
                    if (n == null) return "无 textbox";
                    return n.Z.HasValue && n.Z.Value >= 1 ? null : "z 未读到: " + n.Z;
                });

            Add(list, ref page, "tb-replace-font-keep-text",
                a => ContentHtml.TextboxCreate("保留正文", extraAttrs: "data-font-size=\"16pt\""),
                (a, ids) => ContentHtml.TextboxUpdate(ids[0], "保留正文", "data-font-size=\"26pt\" data-font-bold=\"true\""),
                ctx =>
                {
                    PptHtmlShapeNode n = PreferCreated(ctx, "textbox", useReplace: true);
                    string e = TextEq(n, "保留正文");
                    if (e != null) return e;
                    if (!ContentAssert.FontSizeClose(n.FontSizePt, 26)) return "换字号失败: " + n.FontSizePt;
                    if (n.FontBold != true) return "换粗体失败";
                    return null;
                });

            Add(list, ref page, "tb-create-font-combo",
                a => ContentHtml.TextboxCreate(
                    "组合皮",
                    extraAttrs: "data-font-size=\"20pt\" data-font-color=\"#7030A0\" data-font-bold=\"true\" "
                        + "data-align=\"center\" data-valign=\"middle\" data-fill=\"#E2F0D9\""),
                null,
                ctx =>
                {
                    PptHtmlShapeNode n = PreferCreated(ctx, "textbox");
                    if (n == null) return "无 textbox";
                    if (!ContentAssert.FontSizeClose(n.FontSizePt, 20)) return "字号";
                    if (!HexEq(n.FontColor, "#7030A0")) return "字色";
                    if (n.FontBold != true) return "粗体";
                    if (!string.Equals(n.Align, "center", StringComparison.OrdinalIgnoreCase)) return "align";
                    if (!string.Equals(n.Valign, "middle", StringComparison.OrdinalIgnoreCase)) return "valign";
                    if (!HexEq(n.Fill, "#E2F0D9")) return "fill";
                    return null;
                });
        }

        private static void AddTableCases(List<ContentCase> list, ref int page)
        {
            string[][] g2 =
            {
                new[] { "A1", "B1" },
                new[] { "A2", "B2" }
            };
            string[][] g2b =
            {
                new[] { "X1", "Y1" },
                new[] { "X2", "Y2" }
            };
            string[][] g3 =
            {
                new[] { "品类", "Q1", "Q2" },
                new[] { "甲", "1", "2" },
                new[] { "乙", "3", "4" }
            };
            string[][] g1x4 =
            {
                new[] { "一", "二", "三", "四" }
            };

            Add(list, ref page, "tbl-create-grid-2x2",
                a => ContentHtml.TableCreate(g2),
                null,
                ctx => ContentAssert.TableCellsMismatch(PreferCreated(ctx, "table"), g2));

            Add(list, ref page, "tbl-replace-grid-2x2",
                a => ContentHtml.TableCreate(g2),
                (a, ids) => ContentHtml.TableUpdate(ids[0], g2b),
                ctx => ContentAssert.TableCellsMismatch(
                    PreferCreated(ctx, "table", useReplace: true), g2b));

            Add(list, ref page, "tbl-create-grid-3x3",
                a => ContentHtml.TableCreate(g3),
                null,
                ctx => ContentAssert.TableCellsMismatch(PreferCreated(ctx, "table"), g3));

            Add(list, ref page, "tbl-create-grid-1x4",
                a => ContentHtml.TableCreate(g1x4, width: 60, height: 12),
                null,
                ctx => ContentAssert.TableCellsMismatch(PreferCreated(ctx, "table"), g1x4));

            Add(list, ref page, "tbl-create-fill",
                a => ContentHtml.TableCreate(g2, extraAttrs: "data-fill=\"#DEEBF7\""),
                null,
                ctx =>
                {
                    PptHtmlShapeNode n = PreferCreated(ctx, "table");
                    if (n == null) return "无 table";
                    string cell = ContentAssert.TableCellsMismatch(n, g2);
                    if (cell != null) return cell;
                    return HexEq(n.Fill, "#DEEBF7") ? null : "表 fill 期望 #DEEBF7 实际 " + n.Fill;
                });

            Add(list, ref page, "tbl-create-line",
                a => ContentHtml.TableCreate(
                    g2,
                    extraAttrs: "data-line-color=\"#548235\" data-line-width=\"1.5pt\""),
                null,
                ctx =>
                {
                    // 表形描边在部分 Office 上写 Line.Weight 会抛范围错误；引擎已跳过表描边。
                    // 本例只确认带描边属性时网格仍能创建成功。
                    return ContentAssert.TableCellsMismatch(PreferCreated(ctx, "table"), g2);
                });

            Add(list, ref page, "tbl-create-geo",
                a => ContentHtml.TableCreate(g2, left: 20, top: 25, width: 50, height: 35),
                null,
                ctx =>
                {
                    PptHtmlShapeNode n = PreferCreated(ctx, "table");
                    if (n == null) return "无 table";
                    if (!ContentAssert.TryParseGeo(n.Style, out double l, out double t, out double w, out double h))
                    {
                        return "无几何";
                    }

                    if (Math.Abs(l - 20) > 0.8 || Math.Abs(t - 25) > 0.8
                        || Math.Abs(w - 50) > 0.8 || Math.Abs(h - 35) > 0.8)
                    {
                        return "几何偏离: " + n.Style;
                    }

                    return null;
                });

            Add(list, ref page, "tbl-lean-grid-keep-geo",
                a => ContentHtml.TableCreate(g2, left: 18, top: 28, width: 40, height: 28),
                (a, ids) => ContentHtml.TableUpdate(ids[0], g2b),
                ctx =>
                {
                    PptHtmlShapeNode before = PreferCreated(ctx, "table", useReplace: false);
                    PptHtmlShapeNode after = PreferCreated(ctx, "table", useReplace: true);
                    if (before == null || after == null) return "缺 table";
                    string cell = ContentAssert.TableCellsMismatch(after, g2b);
                    if (cell != null) return cell;
                    return ContentAssert.GeoClose(before.Style, after.Style)
                        ? null
                        : "瘦稿换表不应丢几何";
                });

            // D21：原地扩表（2×2 → 3×3），禁止删表重建
            Add(list, ref page, "tbl-grow-2x2-to-3x3",
                a => ContentHtml.TableCreate(g2),
                (a, ids) => ContentHtml.TableUpdate(ids[0], g3),
                ctx => ContentAssert.TableCellsMismatch(
                    PreferCreated(ctx, "table", useReplace: true), g3));

            // 合并表头 + 三线 + 列宽%（增强表方言）
            Add(list, ref page, "tbl-create-merge-three-line",
                a => ContentHtml.TableCreateMergeThreeLine(),
                null,
                ctx => ContentAssert.TableMergeThreeLineOk(PreferCreated(ctx, "table")));

            // 属性交叉矩阵（≥100）：tbl-mx-*
            AddTableMatrixCases(list, ref page);
        }

        private static void AddPictureCases(List<ContentCase> list, ref int page)
        {
            Add(list, ref page, "pic-create",
                a => ContentHtml.PictureCreate(a.RedPng),
                null,
                ctx =>
                {
                    PptHtmlShapeNode n = PreferCreated(ctx, "picture");
                    if (n == null) return "无 picture";
                    string err = ContentAssert.AttachPictureSrc(
                        ctx.Presentation, ctx.AfterCreate, ctx.ExportDir);
                    if (err != null) return "导出失败: " + err;
                    return string.IsNullOrEmpty(n.DataSrc) ? "导出后无 data-src" : null;
                });

            Add(list, ref page, "pic-replace-src-geo-keep",
                a => ContentHtml.PictureCreate(a.RedPng, 55, 20, 35, 40),
                (a, ids) => ContentHtml.PictureUpdate(ids[0], a.BluePng),
                ctx =>
                {
                    PptHtmlShapeNode before = PreferCreated(ctx, "picture", useReplace: false);
                    PptHtmlShapeNode after = PreferCreated(ctx, "picture", useReplace: true);
                    if (before == null || after == null) return "换图前后缺 picture";
                    if (!ContentAssert.GeoClose(before.Style, after.Style))
                    {
                        return "换图应保几何";
                    }

                    string err2 = ContentAssert.AttachPictureSrc(
                        ctx.Presentation,
                        ctx.AfterReplace,
                        Path.Combine(ctx.ExportDir, "after"));
                    if (err2 != null) return "导出失败: " + err2;
                    if (string.IsNullOrEmpty(before.DataSrc) || string.IsNullOrEmpty(after.DataSrc))
                    {
                        return "data-src 空 before=" + before.DataSrc + " after=" + after.DataSrc;
                    }

                    return string.Equals(before.DataSrc, after.DataSrc, StringComparison.OrdinalIgnoreCase)
                        ? "换图后 data-src 未变"
                        : null;
                });

            Add(list, ref page, "pic-create-geo",
                a => ContentHtml.PictureCreate(a.BluePng, 10, 40, 25, 30),
                null,
                ctx =>
                {
                    PptHtmlShapeNode n = PreferCreated(ctx, "picture");
                    if (n == null) return "无 picture";
                    if (!ContentAssert.TryParseGeo(n.Style, out double l, out double t, out double w, out double h))
                    {
                        return "无几何";
                    }

                    if (Math.Abs(l - 10) > 0.8 || Math.Abs(t - 40) > 0.8
                        || Math.Abs(w - 25) > 0.8 || Math.Abs(h - 30) > 0.8)
                    {
                        return "几何偏离: " + n.Style;
                    }

                    return null;
                });

            Add(list, ref page, "pic-replace-to-green",
                a => ContentHtml.PictureCreate(a.RedPng, 50, 15, 30, 35),
                (a, ids) => ContentHtml.PictureUpdate(ids[0], a.GreenPng),
                ctx =>
                {
                    PptHtmlShapeNode before = PreferCreated(ctx, "picture", useReplace: false);
                    PptHtmlShapeNode after = PreferCreated(ctx, "picture", useReplace: true);
                    if (before == null || after == null) return "缺 picture";
                    string err = ContentAssert.AttachPictureSrc(
                        ctx.Presentation,
                        ctx.AfterReplace,
                        Path.Combine(ctx.ExportDir, "after"));
                    if (err != null) return err;
                    if (string.IsNullOrEmpty(before.DataSrc) || string.IsNullOrEmpty(after.DataSrc))
                    {
                        return "data-src 空";
                    }

                    return string.Equals(before.DataSrc, after.DataSrc, StringComparison.OrdinalIgnoreCase)
                        ? "绿图 src 未变"
                        : null;
                });

            // picture 更新路径故意不写 HTML 几何（换图保框）；改测 textbox 显式改几何
            Add(list, ref page, "tb-update-geo-only",
                a => ContentHtml.TextboxCreate("几何框", left: 55, top: 20, width: 35, height: 20),
                (a, ids) => ContentHtml.Section(
                    "  <div ShapeId=\"" + ids[0] + "\" style=\"" + ContentHtml.Geo(8, 8, 40, 25) + "\">几何框</div>"),
                ctx =>
                {
                    PptHtmlShapeNode after = PreferCreated(ctx, "textbox", useReplace: true);
                    if (after == null) return "无 textbox";
                    if (!ContentAssert.TryParseGeo(after.Style, out double l, out double t, out double w, out double h))
                    {
                        return "无几何";
                    }

                    if (Math.Abs(l - 8) > 1 || Math.Abs(t - 8) > 1
                        || Math.Abs(w - 40) > 1 || Math.Abs(h - 25) > 1)
                    {
                        return "显式几何未生效: " + after.Style;
                    }

                    return null;
                });

            Add(list, ref page, "pic-neg-http",
                a => ContentHtml.Section(
                    "  <img data-shape-type=\"picture\" data-src=\"https://example.com/a.png\" "
                    + "style=\"position:absolute;left:10%;top:10%;width:20%;height:20%\" />"),
                null,
                null,
                expectApplyErrorContains: "不存在");

            Add(list, ref page, "pic-neg-data-uri",
                a => ContentHtml.Section(
                    "  <img data-shape-type=\"picture\" "
                    + "data-src=\"data:image/png;base64,iVBORw0KGgo=\" "
                    + "style=\"position:absolute;left:10%;top:10%;width:20%;height:20%\" />"),
                null,
                null,
                expectApplyErrorContains: "data");
        }

        private static void AddMixCases(List<ContentCase> list, ref int page)
        {
            string[][] mixGrid =
            {
                new[] { "品类", "销量" },
                new[] { "甲", "12" },
                new[] { "乙", "9" }
            };
            string[][] mixGrid2 =
            {
                new[] { "品类", "销量" },
                new[] { "丙", "7" },
                new[] { "丁", "5" }
            };

            Add(list, ref page, "mix-create-three",
                a => ContentHtml.MixCreate("混排标题", mixGrid, a.RedPng),
                null,
                ctx =>
                {
                    PptHtmlShapeNode tb = PreferCreated(ctx, "textbox");
                    if (tb == null) return "缺 textbox";
                    string tbl = ContentAssert.TableCellsMismatch(PreferCreated(ctx, "table"), mixGrid);
                    if (tbl != null) return tbl;
                    if (PreferCreated(ctx, "picture") == null) return "缺 picture";
                    return TextEq(tb, "混排标题");
                });

            Add(list, ref page, "mix-lean-textbox",
                a => ContentHtml.MixCreate("原标题", mixGrid, a.RedPng),
                (a, ids) =>
                {
                    string tbId = FindIdByType(ids, a, "textbox") ?? ids[0];
                    return ContentHtml.TextboxUpdate(tbId, "只改标题");
                },
                ctx =>
                {
                    PptHtmlShapeNode tb = PreferCreated(ctx, "textbox", useReplace: true);
                    string e = TextEq(tb, "只改标题");
                    if (e != null) return e;
                    string tbl = ContentAssert.TableCellsMismatch(
                        PreferCreated(ctx, "table", useReplace: true), mixGrid);
                    if (tbl != null) return "改标题误伤表: " + tbl;
                    if (PreferCreated(ctx, "picture", useReplace: true) == null) return "图片丢了";
                    return null;
                });

            Add(list, ref page, "mix-lean-table",
                a => ContentHtml.MixCreate("标题不动", mixGrid, a.RedPng),
                (a, ids) =>
                {
                    string tblId = FindIdByType(ids, a, "table") ?? (ids.Count > 1 ? ids[1] : ids[0]);
                    return ContentHtml.TableUpdate(tblId, mixGrid2);
                },
                ctx =>
                {
                    string e = TextEq(PreferCreated(ctx, "textbox", useReplace: true), "标题不动");
                    if (e != null) return "换表误伤标题: " + e;
                    return ContentAssert.TableCellsMismatch(
                        PreferCreated(ctx, "table", useReplace: true), mixGrid2);
                });

            Add(list, ref page, "mix-lean-picture",
                a => ContentHtml.MixCreate("标题仍在", mixGrid, a.RedPng),
                (a, ids) =>
                {
                    string picId = FindIdByType(ids, a, "picture") ?? ids[ids.Count - 1];
                    return ContentHtml.PictureUpdate(picId, a.BluePng);
                },
                ctx =>
                {
                    string e = TextEq(PreferCreated(ctx, "textbox", useReplace: true), "标题仍在");
                    if (e != null) return "换图误伤标题: " + e;
                    string tbl = ContentAssert.TableCellsMismatch(
                        PreferCreated(ctx, "table", useReplace: true), mixGrid);
                    if (tbl != null) return "换图误伤表: " + tbl;
                    PptHtmlShapeNode before = PreferCreated(ctx, "picture", useReplace: false);
                    PptHtmlShapeNode after = PreferCreated(ctx, "picture", useReplace: true);
                    if (before == null || after == null) return "缺 picture";
                    string err = ContentAssert.AttachPictureSrc(
                        ctx.Presentation,
                        ctx.AfterReplace,
                        Path.Combine(ctx.ExportDir, "after"));
                    if (err != null) return err;
                    if (string.IsNullOrEmpty(before.DataSrc) || string.IsNullOrEmpty(after.DataSrc))
                    {
                        return "data-src 空";
                    }

                    return string.Equals(before.DataSrc, after.DataSrc, StringComparison.OrdinalIgnoreCase)
                        ? "混排换图 src 未变"
                        : null;
                });

            Add(list, ref page, "mix-read-content-only",
                a => ContentHtml.MixCreate("骨架可见", mixGrid, a.GreenPng),
                null,
                ctx =>
                {
                    if (PreferCreated(ctx, "textbox") == null) return "content_only 缺 textbox";
                    if (PreferCreated(ctx, "table") == null) return "content_only 缺 table";
                    if (PreferCreated(ctx, "picture") == null) return "content_only 缺 picture";
                    return null;
                },
                contentOnly: true);
        }

        private static string FindIdByType(IList<string> ids, ContentAssets assets, string type)
        {
            // createdIds 按创建顺序：mix 固定 textbox → table → picture
            if (ids == null || ids.Count == 0) return null;
            if (string.Equals(type, "textbox", StringComparison.OrdinalIgnoreCase)) return ids[0];
            if (string.Equals(type, "table", StringComparison.OrdinalIgnoreCase))
            {
                return ids.Count > 1 ? ids[1] : ids[0];
            }

            if (string.Equals(type, "picture", StringComparison.OrdinalIgnoreCase))
            {
                return ids[ids.Count - 1];
            }

            return ids[0];
        }

        private static string TextEq(PptHtmlShapeNode n, string expect)
        {
            if (n == null) return "节点为空";
            return string.Equals(ContentAssert.NormText(n.Text), expect, StringComparison.Ordinal)
                ? null
                : "正文期望 " + expect + " 实际 " + n.Text;
        }

        private static string AlignEq(PptHtmlShapeNode n, string expect)
        {
            if (n == null) return "无 textbox";
            return string.Equals(n.Align, expect, StringComparison.OrdinalIgnoreCase)
                ? null
                : "align 期望 " + expect + " 实际 " + n.Align;
        }

        private static string ValignEq(PptHtmlShapeNode n, string expect)
        {
            if (n == null) return "无 textbox";
            return string.Equals(n.Valign, expect, StringComparison.OrdinalIgnoreCase)
                ? null
                : "valign 期望 " + expect + " 实际 " + n.Valign;
        }

        private static bool HexEq(string actual, string expect)
        {
            return string.Equals(
                ContentAssert.HexOrNull(actual),
                ContentAssert.HexOrNull(expect),
                StringComparison.OrdinalIgnoreCase);
        }

        private static bool Near(double? actual, double expect, double eps = 0.8)
        {
            return actual.HasValue && Math.Abs(actual.Value - expect) <= eps;
        }

        private static string EscapeAttr(string s)
        {
            if (string.IsNullOrEmpty(s)) return "";
            return s.Replace("&", "&amp;").Replace("\"", "&quot;").Replace("<", "&lt;");
        }

        private static void Add(
            List<ContentCase> list,
            ref int page,
            string name,
            Func<ContentAssets, string> create,
            Func<ContentAssets, IList<string>, string> replace,
            Func<ContentAssertContext, string> match,
            string expectApplyErrorContains = null,
            bool contentOnly = false)
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
                ExpectApplyErrorContains = expectApplyErrorContains,
                ContentOnly = contentOnly
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

            if (result?.Shapes == null) return null;

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
