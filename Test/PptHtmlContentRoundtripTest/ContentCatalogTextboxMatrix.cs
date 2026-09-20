using System;
using System.Collections.Generic;
using System.Globalization;
using WordAddIn1.PresentationHost;

namespace PptHtmlContentRoundtripTest
{
    /// <summary>
    /// 文本框属性交叉矩阵（tb-mx-*）。覆盖字皮组合、对齐交叉、填充/线、
    /// 边距/段距、bullet、几何、全皮抽样与换文瘦稿；合计 ≥100。
    /// </summary>
    internal static partial class ContentCatalog
    {
        private static void AddTextboxMatrixCases(List<ContentCase> list, ref int page)
        {
            int before = list.Count;
            AddTbMxFontCombos(list, ref page);
            AddTbMxAlignValign(list, ref page);
            AddTbMxFillLine(list, ref page);
            AddTbMxMarginPara(list, ref page);
            AddTbMxBulletAlign(list, ref page);
            AddTbMxGeo(list, ref page);
            AddTbMxSkinCross(list, ref page);
            AddTbMxUpdates(list, ref page);
            int added = list.Count - before;
            if (added < 100)
            {
                throw new InvalidOperationException(
                    "文本框矩阵用例不足 100（实际 " + added + "），请扩 ContentCatalogTextboxMatrix。");
            }
        }

        private static void AddTbMxFontCombos(List<ContentCase> list, ref int page)
        {
            double[] sizes = { 12, 16, 20, 28 };
            string[] colors = { "#0070C0", "#C00000", "#548235" };
            bool[] bolds = { true, false };

            int i = 0;
            foreach (double size in sizes)
            {
                foreach (string color in colors)
                {
                    foreach (bool bold in bolds)
                    {
                        double es = size;
                        string ec = color;
                        bool eb = bold;
                        string name = "tb-mx-font-" + i.ToString("00", CultureInfo.InvariantCulture)
                            + "-s" + Tag(es)
                            + "-" + ec.TrimStart('#')
                            + (eb ? "-b" : "-n");
                        Add(list, ref page, name,
                            a => ContentHtml.TextboxCreate(
                                "字皮" + i,
                                extraAttrs: "data-font-size=\"" + Pt(es) + "\" data-font-color=\"" + ec
                                    + "\" data-font-bold=\"" + (eb ? "true" : "false") + "\""),
                            null,
                            ctx =>
                            {
                                PptHtmlShapeNode n = PreferCreated(ctx, "textbox");
                                if (n == null) return "无 textbox";
                                return ContentAssert.FirstFail(
                                    ContentAssert.FontSizeClose(n.FontSizePt, es)
                                        ? null
                                        : "字号期望 " + es + " 实际 " + n.FontSizePt,
                                    HexEq(n.FontColor, ec) ? null : "字色期望 " + ec + " 实际 " + n.FontColor,
                                    n.FontBold == eb ? null : "粗体期望 " + eb + " 实际 " + n.FontBold);
                            });
                        i++;
                    }
                }
            }
        }

        private static void AddTbMxAlignValign(List<ContentCase> list, ref int page)
        {
            string[] aligns = { "left", "center", "right", "justify" };
            string[] valigns = { "top", "middle", "bottom" };
            int i = 0;
            foreach (string align in aligns)
            {
                foreach (string valign in valigns)
                {
                    string ea = align;
                    string ev = valign;
                    string body = ea == "justify"
                        ? "两端对齐文字两端对齐文字两端对齐"
                        : "对齐" + i;
                    Add(list, ref page, "tb-mx-align-" + ea + "-" + ev,
                        a => ContentHtml.TextboxCreate(
                            body,
                            height: 28,
                            extraAttrs: "data-align=\"" + ea + "\" data-valign=\"" + ev + "\""),
                        null,
                        ctx =>
                        {
                            PptHtmlShapeNode n = PreferCreated(ctx, "textbox");
                            return ContentAssert.FirstFail(
                                AlignEq(n, ea),
                                ValignEq(n, ev));
                        });
                    i++;
                }
            }
        }

        private static void AddTbMxFillLine(List<ContentCase> list, ref int page)
        {
            var fills = new[] { "#FFF2CC", "#DEEBF7", "#E2F0D9", "#FCE4D6", "none" };
            var lines = new[]
            {
                new { C = "#C00000", W = 1.5 },
                new { C = "#0070C0", W = 2.0 },
                new { C = "#7030A0", W = 2.5 }
            };

            int i = 0;
            foreach (string fill in fills)
            {
                foreach (var line in lines)
                {
                    if (i >= 12) break;
                    string ef = fill;
                    string elc = line.C;
                    double elw = line.W;
                    string tag = ef == "none" ? "none" : ef.TrimStart('#');
                    Add(list, ref page, "tb-mx-fillline-" + i.ToString("00", CultureInfo.InvariantCulture)
                            + "-" + tag + "-" + elc.TrimStart('#'),
                        a => ContentHtml.TextboxCreate(
                            "框皮" + i,
                            extraAttrs: "data-fill=\"" + ef + "\" data-line-color=\"" + elc
                                + "\" data-line-width=\"" + Pt(elw) + "\""),
                        null,
                        ctx =>
                        {
                            PptHtmlShapeNode n = PreferCreated(ctx, "textbox");
                            if (n == null) return "无 textbox";
                            string fillErr = FillEq(n, ef);
                            if (fillErr != null) return fillErr;
                            if (!HexEq(n.LineColor, elc)) return "线色期望 " + elc + " 实际 " + n.LineColor;
                            if (!n.LineWidthPt.HasValue || Math.Abs(n.LineWidthPt.Value - elw) > 0.6)
                            {
                                return "线宽期望 " + elw + " 实际 " + n.LineWidthPt;
                            }

                            return null;
                        });
                    i++;
                }
            }
        }

        private static void AddTbMxMarginPara(List<ContentCase> list, ref int page)
        {
            var margins = new[]
            {
                new { L = 8.0, R = 6.0, T = 4.0, B = 5.0 },
                new { L = 2.0, R = 2.0, T = 2.0, B = 2.0 },
                new { L = 12.0, R = 10.0, T = 6.0, B = 8.0 },
                new { L = 0.0, R = 4.0, T = 8.0, B = 0.0 },
                new { L = 5.0, R = 5.0, T = 10.0, B = 10.0 },
                new { L = 16.0, R = 3.0, T = 3.0, B = 16.0 }
            };

            for (int i = 0; i < margins.Length; i++)
            {
                var m = margins[i];
                Add(list, ref page, "tb-mx-margin-" + i.ToString("00", CultureInfo.InvariantCulture),
                    a => ContentHtml.TextboxCreate(
                        "边距" + i,
                        extraAttrs: "data-margin-left=\"" + Pt(m.L) + "\" data-margin-right=\"" + Pt(m.R)
                            + "\" data-margin-top=\"" + Pt(m.T) + "\" data-margin-bottom=\"" + Pt(m.B) + "\""),
                    null,
                    ctx =>
                    {
                        PptHtmlShapeNode n = PreferCreated(ctx, "textbox");
                        if (n == null) return "无 textbox";
                        if (!Near(n.MarginLeftPt, m.L)) return "margin-left 期望 " + m.L + " 实际 " + n.MarginLeftPt;
                        if (!Near(n.MarginRightPt, m.R)) return "margin-right 期望 " + m.R + " 实际 " + n.MarginRightPt;
                        if (!Near(n.MarginTopPt, m.T)) return "margin-top 期望 " + m.T + " 实际 " + n.MarginTopPt;
                        if (!Near(n.MarginBottomPt, m.B)) return "margin-bottom 期望 " + m.B + " 实际 " + n.MarginBottomPt;
                        return null;
                    });
            }

            var paras = new[]
            {
                new { Bf = 10.0, Af = 8.0, Ls = "1.5" },
                new { Bf = 0.0, Af = 12.0, Ls = "1.25" },
                new { Bf = 6.0, Af = 6.0, Ls = "1.75" },
                new { Bf = 14.0, Af = 4.0, Ls = "1.2" },
                new { Bf = 4.0, Af = 14.0, Ls = "1.8" },
                new { Bf = 8.0, Af = 0.0, Ls = "1.15" }
            };

            for (int i = 0; i < paras.Length; i++)
            {
                var p = paras[i];
                Add(list, ref page, "tb-mx-para-" + i.ToString("00", CultureInfo.InvariantCulture),
                    a => ContentHtml.TextboxCreate(
                        "段距" + i,
                        extraAttrs: "data-space-before=\"" + Pt(p.Bf) + "\" data-space-after=\"" + Pt(p.Af)
                            + "\" data-line-spacing=\"" + p.Ls + "\""),
                    null,
                    ctx =>
                    {
                        PptHtmlShapeNode n = PreferCreated(ctx, "textbox");
                        if (n == null) return "无 textbox";
                        if (!Near(n.SpaceBeforePt, p.Bf)) return "space-before 期望 " + p.Bf + " 实际 " + n.SpaceBeforePt;
                        if (!Near(n.SpaceAfterPt, p.Af)) return "space-after 期望 " + p.Af + " 实际 " + n.SpaceAfterPt;
                        return LineSpacingClose(n.LineSpacing, p.Ls, n.FontSizePt)
                            ? null
                            : "line-spacing 期望 " + p.Ls + " 实际 " + n.LineSpacing;
                    });
            }

            // 整数倍行距探针（对照 WPP exact:N 读回）；正式矩阵用 1.25/1.75 避开
            foreach (var dbg in new[]
            {
                new { Tag = "1", Ls = "1.0" },
                new { Tag = "2", Ls = "2.0" }
            })
            {
                string ls = dbg.Ls;
                Add(list, ref page, "tb-mx-dbg-ls-" + dbg.Tag,
                    a => ContentHtml.TextboxCreate(
                        "行距探针" + dbg.Tag,
                        extraAttrs: "data-line-spacing=\"" + ls + "\""),
                    null,
                    ctx =>
                    {
                        PptHtmlShapeNode n = PreferCreated(ctx, "textbox");
                        if (n == null) return "无 textbox";
                        return LineSpacingClose(n.LineSpacing, ls, n.FontSizePt)
                            ? null
                            : "line-spacing 期望倍数 " + ls + " 实际 " + n.LineSpacing;
                    });
            }

            // 肉眼对照：左 data-line-spacing=1.0（疑似写成 exact:1pt），右 1.5 倍数
            string multi =
                "第一行文字AAAA\n第二行文字BBBB\n第三行文字CCCC\n第四行文字DDDD\n第五行文字EEEE";
            Add(list, ref page, "tb-mx-dbg-ls-visual",
                a => ContentHtml.Section(
                    "  <div data-shape-type=\"textbox\" style=\""
                    + ContentHtml.Geo(5, 8, 42, 70)
                    + "\" data-font-size=\"18pt\" data-line-spacing=\"1.0\" data-fill=\"#FFF2CC\">"
                    + "【左】约定写倍数 1.0\n" + multi + "</div>\n"
                    + "  <div data-shape-type=\"textbox\" style=\""
                    + ContentHtml.Geo(52, 8, 42, 70)
                    + "\" data-font-size=\"18pt\" data-line-spacing=\"1.5\" data-fill=\"#DEEBF7\">"
                    + "【右】约定写倍数 1.5\n" + multi + "</div>"),
                null,
                ctx =>
                {
                    // 演示页：不卡断言，只保证两个 textbox 建出来
                    int n = 0;
                    if (ctx.AfterCreate?.Shapes != null)
                    {
                        for (int i = 0; i < ctx.AfterCreate.Shapes.Count; i++)
                        {
                            if (string.Equals(
                                    ctx.AfterCreate.Shapes[i].ShapeType,
                                    "textbox",
                                    StringComparison.OrdinalIgnoreCase))
                            {
                                n++;
                            }
                        }
                    }

                    return n >= 2 ? null : "演示页 textbox 不足 2 实际 " + n;
                });
        }

        private static void AddTbMxBulletAlign(List<ContentCase> list, ref int page)
        {
            string[] bullets = { "bullet", "number" };
            string[] aligns = { "left", "center", "right", "justify" };
            int i = 0;
            foreach (string bullet in bullets)
            {
                foreach (string align in aligns)
                {
                    string eb = bullet;
                    string ea = align;
                    string body = ea == "justify" ? "列表两端对齐文字列表两端对齐" : "列表" + i;
                    Add(list, ref page, "tb-mx-bullet-" + eb + "-" + ea,
                        a => ContentHtml.TextboxCreate(
                            body,
                            extraAttrs: "data-bullet=\"" + eb + "\" data-align=\"" + ea + "\""),
                        null,
                        ctx =>
                        {
                            PptHtmlShapeNode n = PreferCreated(ctx, "textbox");
                            if (n == null) return "无 textbox";
                            if (!string.Equals(n.Bullet, eb, StringComparison.OrdinalIgnoreCase))
                            {
                                return "bullet 期望 " + eb + " 实际 " + n.Bullet;
                            }

                            return AlignEq(n, ea);
                        });
                    i++;
                }
            }
        }

        private static void AddTbMxGeo(List<ContentCase> list, ref int page)
        {
            double[][] geos =
            {
                new[] { 5.0, 5.0, 30.0, 18.0 }, new[] { 10.0, 12.0, 40.0, 22.0 }, new[] { 15.0, 20.0, 35.0, 25.0 },
                new[] { 20.0, 8.0, 45.0, 16.0 }, new[] { 25.0, 30.0, 28.0, 28.0 }, new[] { 30.0, 15.0, 50.0, 20.0 },
                new[] { 35.0, 40.0, 25.0, 30.0 }, new[] { 40.0, 10.0, 35.0, 35.0 }, new[] { 45.0, 25.0, 30.0, 20.0 },
                new[] { 50.0, 5.0, 40.0, 18.0 }, new[] { 8.0, 45.0, 55.0, 20.0 }, new[] { 12.0, 50.0, 30.0, 25.0 },
                new[] { 18.0, 35.0, 42.0, 22.0 }, new[] { 22.0, 55.0, 38.0, 18.0 }, new[] { 28.0, 42.0, 32.0, 24.0 },
                new[] { 33.0, 18.0, 28.0, 40.0 }, new[] { 38.0, 48.0, 36.0, 20.0 }, new[] { 42.0, 22.0, 40.0, 26.0 },
                new[] { 48.0, 12.0, 25.0, 32.0 }, new[] { 55.0, 35.0, 30.0, 28.0 }
            };

            for (int i = 0; i < geos.Length; i++)
            {
                double l = geos[i][0], t = geos[i][1], w = geos[i][2], h = geos[i][3];
                double el = l, et = t, ew = w, eh = h;
                Add(list, ref page, "tb-mx-geo-" + i.ToString("00", CultureInfo.InvariantCulture)
                        + "-" + Tag(l) + "x" + Tag(t) + "-" + Tag(w) + "x" + Tag(h),
                    a => ContentHtml.TextboxCreate("几何" + i, el, et, ew, eh),
                    null,
                    ctx => TbGeoEq(PreferCreated(ctx, "textbox"), el, et, ew, eh));
            }
        }

        private static void AddTbMxSkinCross(List<ContentCase> list, ref int page)
        {
            var specs = new[]
            {
                new
                {
                    Tag = "title",
                    Text = "标题皮",
                    Attr = "data-font-size=\"28pt\" data-font-color=\"#1F4E79\" data-font-bold=\"true\" "
                        + "data-align=\"center\" data-valign=\"middle\" data-fill=\"#D6DCE4\" "
                        + "data-font-name=\"微软雅黑\"",
                    Size = 28.0, Color = "#1F4E79", Bold = true, Align = "center", Valign = "middle",
                    Fill = "#D6DCE4", Yahei = true, LineC = (string)null, LineW = (double?)null
                },
                new
                {
                    Tag = "body",
                    Text = "正文皮",
                    Attr = "data-font-size=\"16pt\" data-font-color=\"#333333\" data-font-bold=\"false\" "
                        + "data-align=\"left\" data-valign=\"top\" data-fill=\"#FFF2CC\" "
                        + "data-line-color=\"#C00000\" data-line-width=\"1.5pt\"",
                    Size = 16.0, Color = "#333333", Bold = false, Align = "left", Valign = "top",
                    Fill = "#FFF2CC", Yahei = false, LineC = "#C00000", LineW = (double?)1.5
                },
                new
                {
                    Tag = "callout",
                    Text = "强调框",
                    Attr = "data-font-size=\"18pt\" data-font-color=\"#FFFFFF\" data-font-bold=\"true\" "
                        + "data-align=\"center\" data-valign=\"middle\" data-fill=\"#C00000\" "
                        + "data-line-color=\"#7F0000\" data-line-width=\"2pt\"",
                    Size = 18.0, Color = "#FFFFFF", Bold = true, Align = "center", Valign = "middle",
                    Fill = "#C00000", Yahei = false, LineC = "#7F0000", LineW = (double?)2.0
                },
                new
                {
                    Tag = "note",
                    Text = "脚注皮",
                    Attr = "data-font-size=\"12pt\" data-font-color=\"#666666\" data-font-bold=\"false\" "
                        + "data-align=\"right\" data-valign=\"bottom\" data-fill=\"none\" "
                        + "data-margin-left=\"4pt\" data-margin-right=\"4pt\"",
                    Size = 12.0, Color = "#666666", Bold = false, Align = "right", Valign = "bottom",
                    Fill = "none", Yahei = false, LineC = (string)null, LineW = (double?)null
                },
                new
                {
                    Tag = "quote",
                    Text = "引用段",
                    Attr = "data-font-size=\"20pt\" data-font-color=\"#7030A0\" data-font-bold=\"true\" "
                        + "data-align=\"justify\" data-valign=\"middle\" data-fill=\"#E2F0D9\" "
                        + "data-space-before=\"8pt\" data-space-after=\"8pt\" data-line-spacing=\"1.5\"",
                    Size = 20.0, Color = "#7030A0", Bold = true, Align = "justify", Valign = "middle",
                    Fill = "#E2F0D9", Yahei = false, LineC = (string)null, LineW = (double?)null
                },
                new
                {
                    Tag = "badge",
                    Text = "徽标",
                    Attr = "data-font-size=\"14pt\" data-font-color=\"#FFFFFF\" data-font-bold=\"true\" "
                        + "data-align=\"center\" data-valign=\"middle\" data-fill=\"#0070C0\" "
                        + "data-font-name=\"微软雅黑\" data-line-color=\"#002060\" data-line-width=\"2.5pt\"",
                    Size = 14.0, Color = "#FFFFFF", Bold = true, Align = "center", Valign = "middle",
                    Fill = "#0070C0", Yahei = true, LineC = "#002060", LineW = (double?)2.5
                },
                new
                {
                    Tag = "warn",
                    Text = "警示",
                    Attr = "data-font-size=\"22pt\" data-font-color=\"#C00000\" data-font-bold=\"true\" "
                        + "data-align=\"left\" data-valign=\"top\" data-fill=\"#FCE4D6\" "
                        + "data-line-color=\"#C00000\" data-line-width=\"2pt\" data-bullet=\"bullet\"",
                    Size = 22.0, Color = "#C00000", Bold = true, Align = "left", Valign = "top",
                    Fill = "#FCE4D6", Yahei = false, LineC = "#C00000", LineW = (double?)2.0
                },
                new
                {
                    Tag = "list",
                    Text = "编号皮",
                    Attr = "data-font-size=\"16pt\" data-font-color=\"#1F4E79\" data-font-bold=\"false\" "
                        + "data-align=\"left\" data-valign=\"top\" data-fill=\"#DEEBF7\" "
                        + "data-bullet=\"number\" data-margin-left=\"10pt\"",
                    Size = 16.0, Color = "#1F4E79", Bold = false, Align = "left", Valign = "top",
                    Fill = "#DEEBF7", Yahei = false, LineC = (string)null, LineW = (double?)null
                },
                new
                {
                    Tag = "tall",
                    Text = "高框中齐",
                    Attr = "data-font-size=\"18pt\" data-font-color=\"#548235\" data-font-bold=\"true\" "
                        + "data-align=\"center\" data-valign=\"middle\" data-fill=\"#E2F0D9\"",
                    Size = 18.0, Color = "#548235", Bold = true, Align = "center", Valign = "middle",
                    Fill = "#E2F0D9", Yahei = false, LineC = (string)null, LineW = (double?)null
                },
                new
                {
                    Tag = "wide",
                    Text = "宽框底齐",
                    Attr = "data-font-size=\"15pt\" data-font-color=\"#002060\" data-font-bold=\"false\" "
                        + "data-align=\"justify\" data-valign=\"bottom\" data-fill=\"#FFF2CC\" "
                        + "data-line-color=\"#ED7D31\" data-line-width=\"1.5pt\"",
                    Size = 15.0, Color = "#002060", Bold = false, Align = "justify", Valign = "bottom",
                    Fill = "#FFF2CC", Yahei = false, LineC = "#ED7D31", LineW = (double?)1.5
                },
                new
                {
                    Tag = "compact",
                    Text = "紧凑",
                    Attr = "data-font-size=\"11pt\" data-font-color=\"#595959\" data-font-bold=\"false\" "
                        + "data-align=\"left\" data-valign=\"top\" data-fill=\"none\" "
                        + "data-margin-left=\"2pt\" data-margin-right=\"2pt\" data-margin-top=\"2pt\" "
                        + "data-margin-bottom=\"2pt\"",
                    Size = 11.0, Color = "#595959", Bold = false, Align = "left", Valign = "top",
                    Fill = "none", Yahei = false, LineC = (string)null, LineW = (double?)null
                },
                new
                {
                    Tag = "hero",
                    Text = "主标题",
                    Attr = "data-font-size=\"32pt\" data-font-color=\"#FFFFFF\" data-font-bold=\"true\" "
                        + "data-align=\"center\" data-valign=\"middle\" data-fill=\"#1F4E79\" "
                        + "data-font-name=\"微软雅黑\" data-z=\"3\"",
                    Size = 32.0, Color = "#FFFFFF", Bold = true, Align = "center", Valign = "middle",
                    Fill = "#1F4E79", Yahei = true, LineC = (string)null, LineW = (double?)null
                }
            };

            for (int i = 0; i < specs.Length; i++)
            {
                var s = specs[i];
                double left = 8 + (i % 4) * 18;
                double top = 8 + (i / 4) * 28;
                double height = s.Tag == "tall" ? 40 : (s.Tag == "wide" ? 22 : 24);
                double width = s.Tag == "wide" ? 70 : 40;
                Add(list, ref page, "tb-mx-x-skin-" + s.Tag,
                    a => ContentHtml.TextboxCreate(s.Text, left, top, width, height, s.Attr),
                    null,
                    ctx =>
                    {
                        PptHtmlShapeNode n = PreferCreated(ctx, "textbox");
                        if (n == null) return "无 textbox";
                        string err = ContentAssert.FirstFail(
                            ContentAssert.FontSizeClose(n.FontSizePt, s.Size)
                                ? null
                                : "字号",
                            HexEq(n.FontColor, s.Color) ? null : "字色",
                            n.FontBold == s.Bold ? null : "粗体",
                            AlignEq(n, s.Align),
                            ValignEq(n, s.Valign),
                            FillEq(n, s.Fill));
                        if (err != null) return err;
                        if (s.Yahei)
                        {
                            if (string.IsNullOrEmpty(n.FontName)
                                || (n.FontName.IndexOf("雅黑", StringComparison.Ordinal) < 0
                                    && n.FontName.IndexOf("YaHei", StringComparison.OrdinalIgnoreCase) < 0))
                            {
                                return "字体期望含雅黑 实际 " + n.FontName;
                            }
                        }

                        if (s.LineC != null)
                        {
                            if (!HexEq(n.LineColor, s.LineC)) return "线色 " + n.LineColor;
                            if (!s.LineW.HasValue
                                || !n.LineWidthPt.HasValue
                                || Math.Abs(n.LineWidthPt.Value - s.LineW.Value) > 0.6)
                            {
                                return "线宽 " + n.LineWidthPt;
                            }
                        }

                        if (s.Tag == "warn" || s.Tag == "list")
                        {
                            string expectBullet = s.Tag == "warn" ? "bullet" : "number";
                            if (!string.Equals(n.Bullet, expectBullet, StringComparison.OrdinalIgnoreCase))
                            {
                                return "bullet 期望 " + expectBullet + " 实际 " + n.Bullet;
                            }
                        }

                        if (s.Tag == "quote")
                        {
                            if (!Near(n.SpaceBeforePt, 8) || !Near(n.SpaceAfterPt, 8))
                            {
                                return "段距";
                            }

                            if (!LineSpacingClose(n.LineSpacing, "1.5", n.FontSizePt))
                            {
                                return "行距 " + n.LineSpacing;
                            }
                        }

                        if (s.Tag == "hero" && !(n.Z.HasValue && n.Z.Value >= 1))
                        {
                            return "z 未读到: " + n.Z;
                        }

                        return null;
                    });
            }
        }

        private static void AddTbMxUpdates(List<ContentCase> list, ref int page)
        {
            // 改字保皮
            Add(list, ref page, "tb-mx-upd-text-keep-skin",
                a => ContentHtml.TextboxCreate(
                    "原稿皮",
                    extraAttrs: "data-fill=\"#FFEECC\" data-font-size=\"18pt\" data-font-color=\"#0070C0\" "
                        + "data-font-bold=\"true\" data-align=\"center\""),
                (a, ids) => ContentHtml.TextboxUpdate(ids[0], "改后正文"),
                ctx =>
                {
                    PptHtmlShapeNode n = PreferCreated(ctx, "textbox", useReplace: true);
                    return ContentAssert.FirstFail(
                        TextEq(n, "改后正文"),
                        HexEq(n.Fill, "#FFEECC") ? null : "瘦稿丢 fill: " + n.Fill,
                        ContentAssert.FontSizeClose(n.FontSizePt, 18) ? null : "瘦稿丢字号: " + n.FontSizePt,
                        HexEq(n.FontColor, "#0070C0") ? null : "瘦稿丢字色: " + n.FontColor,
                        n.FontBold == true ? null : "瘦稿丢粗体",
                        AlignEq(n, "center"));
                });

            // 改字号粗体，保 fill / 正文
            Add(list, ref page, "tb-mx-upd-font-keep-fill",
                a => ContentHtml.TextboxCreate(
                    "保留正文",
                    extraAttrs: "data-fill=\"#DEEBF7\" data-font-size=\"14pt\" data-line-color=\"#C00000\" "
                        + "data-line-width=\"2pt\""),
                (a, ids) => ContentHtml.TextboxUpdate(
                    ids[0],
                    "保留正文",
                    "data-font-size=\"24pt\" data-font-bold=\"true\" data-font-color=\"#C00000\""),
                ctx =>
                {
                    PptHtmlShapeNode n = PreferCreated(ctx, "textbox", useReplace: true);
                    return ContentAssert.FirstFail(
                        TextEq(n, "保留正文"),
                        ContentAssert.FontSizeClose(n.FontSizePt, 24) ? null : "字号",
                        n.FontBold == true ? null : "粗体",
                        HexEq(n.FontColor, "#C00000") ? null : "字色",
                        HexEq(n.Fill, "#DEEBF7") ? null : "丢 fill: " + n.Fill,
                        HexEq(n.LineColor, "#C00000") ? null : "丢线色: " + n.LineColor);
                });

            // 只改几何
            Add(list, ref page, "tb-mx-upd-geo-only",
                a => ContentHtml.TextboxCreate("几何框", left: 10, top: 10, width: 30, height: 20,
                    extraAttrs: "data-fill=\"#E2F0D9\" data-font-size=\"16pt\""),
                (a, ids) => ContentHtml.TextboxUpdate(
                    ids[0],
                    "几何框",
                    "style=\"" + ContentHtml.Geo(40, 35, 45, 28) + "\""),
                ctx =>
                {
                    PptHtmlShapeNode n = PreferCreated(ctx, "textbox", useReplace: true);
                    return ContentAssert.FirstFail(
                        TextEq(n, "几何框"),
                        TbGeoEq(n, 40, 35, 45, 28),
                        HexEq(n.Fill, "#E2F0D9") ? null : "改框丢 fill",
                        ContentAssert.FontSizeClose(n.FontSizePt, 16) ? null : "改框丢字号");
                });

            // 改对齐，保字皮+fill
            Add(list, ref page, "tb-mx-upd-align-keep-font",
                a => ContentHtml.TextboxCreate(
                    "对齐原稿",
                    height: 28,
                    extraAttrs: "data-align=\"left\" data-valign=\"top\" data-font-size=\"18pt\" "
                        + "data-font-color=\"#548235\" data-fill=\"#FFF2CC\""),
                (a, ids) => ContentHtml.TextboxUpdate(
                    ids[0],
                    "对齐原稿",
                    "data-align=\"right\" data-valign=\"bottom\""),
                ctx =>
                {
                    PptHtmlShapeNode n = PreferCreated(ctx, "textbox", useReplace: true);
                    return ContentAssert.FirstFail(
                        AlignEq(n, "right"),
                        ValignEq(n, "bottom"),
                        ContentAssert.FontSizeClose(n.FontSizePt, 18) ? null : "字号",
                        HexEq(n.FontColor, "#548235") ? null : "字色",
                        HexEq(n.Fill, "#FFF2CC") ? null : "fill");
                });

            // 换 fill+line，保正文与字号
            Add(list, ref page, "tb-mx-upd-chrome-keep-text",
                a => ContentHtml.TextboxCreate(
                    "框色原稿",
                    extraAttrs: "data-fill=\"#DEEBF7\" data-line-color=\"#0070C0\" data-line-width=\"1pt\" "
                        + "data-font-size=\"17pt\""),
                (a, ids) => ContentHtml.TextboxUpdate(
                    ids[0],
                    "框色原稿",
                    "data-fill=\"#FCE4D6\" data-line-color=\"#C00000\" data-line-width=\"2.5pt\""),
                ctx =>
                {
                    PptHtmlShapeNode n = PreferCreated(ctx, "textbox", useReplace: true);
                    return ContentAssert.FirstFail(
                        TextEq(n, "框色原稿"),
                        ContentAssert.FontSizeClose(n.FontSizePt, 17) ? null : "字号",
                        HexEq(n.Fill, "#FCE4D6") ? null : "fill",
                        HexEq(n.LineColor, "#C00000") ? null : "线色",
                        n.LineWidthPt.HasValue && Math.Abs(n.LineWidthPt.Value - 2.5) <= 0.6
                            ? null
                            : "线宽 " + n.LineWidthPt);
                });

            // 多段改字保皮
            Add(list, ref page, "tb-mx-upd-multipara-keep-skin",
                a => ContentHtml.TextboxCreate(
                    "第一段\n第二段",
                    extraAttrs: "data-font-size=\"15pt\" data-align=\"center\" data-fill=\"#E2F0D9\""),
                (a, ids) => ContentHtml.TextboxUpdate(ids[0], "甲段\n乙段\n丙段"),
                ctx =>
                {
                    PptHtmlShapeNode n = PreferCreated(ctx, "textbox", useReplace: true);
                    if (n == null) return "无 textbox";
                    string t = ContentAssert.NormText(n.Text);
                    if (t.IndexOf("甲段", StringComparison.Ordinal) < 0
                        || t.IndexOf("乙段", StringComparison.Ordinal) < 0
                        || t.IndexOf("丙段", StringComparison.Ordinal) < 0)
                    {
                        return "多段不全: " + t;
                    }

                    return ContentAssert.FirstFail(
                        ContentAssert.FontSizeClose(n.FontSizePt, 15) ? null : "字号",
                        AlignEq(n, "center"),
                        HexEq(n.Fill, "#E2F0D9") ? null : "fill");
                });

            // 加 bullet，保 fill
            Add(list, ref page, "tb-mx-upd-add-bullet-keep-fill",
                a => ContentHtml.TextboxCreate(
                    "条目原文",
                    extraAttrs: "data-fill=\"#FFF2CC\" data-font-size=\"16pt\""),
                (a, ids) => ContentHtml.TextboxUpdate(
                    ids[0],
                    "条目原文",
                    "data-bullet=\"bullet\" data-align=\"left\""),
                ctx =>
                {
                    PptHtmlShapeNode n = PreferCreated(ctx, "textbox", useReplace: true);
                    return ContentAssert.FirstFail(
                        TextEq(n, "条目原文"),
                        string.Equals(n.Bullet, "bullet", StringComparison.OrdinalIgnoreCase)
                            ? null
                            : "bullet " + n.Bullet,
                        HexEq(n.Fill, "#FFF2CC") ? null : "fill",
                        ContentAssert.FontSizeClose(n.FontSizePt, 16) ? null : "字号");
                });

            // 改边距，保字皮
            Add(list, ref page, "tb-mx-upd-margin-keep-font",
                a => ContentHtml.TextboxCreate(
                    "边距原稿",
                    extraAttrs: "data-font-size=\"19pt\" data-font-bold=\"true\" data-font-color=\"#0070C0\" "
                        + "data-margin-left=\"2pt\" data-margin-right=\"2pt\""),
                (a, ids) => ContentHtml.TextboxUpdate(
                    ids[0],
                    "边距原稿",
                    "data-margin-left=\"10pt\" data-margin-right=\"8pt\" data-margin-top=\"6pt\" "
                        + "data-margin-bottom=\"6pt\""),
                ctx =>
                {
                    PptHtmlShapeNode n = PreferCreated(ctx, "textbox", useReplace: true);
                    return ContentAssert.FirstFail(
                        Near(n.MarginLeftPt, 10) ? null : "ml " + n.MarginLeftPt,
                        Near(n.MarginRightPt, 8) ? null : "mr " + n.MarginRightPt,
                        Near(n.MarginTopPt, 6) ? null : "mt " + n.MarginTopPt,
                        Near(n.MarginBottomPt, 6) ? null : "mb " + n.MarginBottomPt,
                        ContentAssert.FontSizeClose(n.FontSizePt, 19) ? null : "字号",
                        n.FontBold == true ? null : "粗体",
                        HexEq(n.FontColor, "#0070C0") ? null : "字色");
                });

            // 字体名 × 字号 × 色
            string[] names = { "微软雅黑", "宋体", "黑体" };
            double[] nameSizes = { 14, 18, 22 };
            for (int i = 0; i < names.Length; i++)
            {
                for (int j = 0; j < nameSizes.Length; j++)
                {
                    string fn = names[i];
                    double fs = nameSizes[j];
                    string tag = (fn.IndexOf("雅黑", StringComparison.Ordinal) >= 0 ? "yahei"
                        : fn.IndexOf("宋", StringComparison.Ordinal) >= 0 ? "song" : "hei")
                        + "-s" + Tag(fs);
                    Add(list, ref page, "tb-mx-fontname-" + tag,
                        a => ContentHtml.TextboxCreate(
                            "字体" + tag,
                            extraAttrs: "data-font-name=\"" + fn + "\" data-font-size=\"" + Pt(fs)
                                + "\" data-font-color=\"#333333\""),
                        null,
                        ctx =>
                        {
                            PptHtmlShapeNode n = PreferCreated(ctx, "textbox");
                            if (n == null) return "无 textbox";
                            if (!ContentAssert.FontSizeClose(n.FontSizePt, fs)) return "字号 " + n.FontSizePt;
                            if (!HexEq(n.FontColor, "#333333")) return "字色 " + n.FontColor;
                            if (string.IsNullOrEmpty(n.FontName)) return "字体名为空";
                            bool ok = fn.IndexOf("雅黑", StringComparison.Ordinal) >= 0
                                ? (n.FontName.IndexOf("雅黑", StringComparison.Ordinal) >= 0
                                    || n.FontName.IndexOf("YaHei", StringComparison.OrdinalIgnoreCase) >= 0)
                                : fn.IndexOf("宋", StringComparison.Ordinal) >= 0
                                    ? (n.FontName.IndexOf("宋", StringComparison.Ordinal) >= 0
                                        || n.FontName.IndexOf("SimSun", StringComparison.OrdinalIgnoreCase) >= 0)
                                    : (n.FontName.IndexOf("黑", StringComparison.Ordinal) >= 0
                                        || n.FontName.IndexOf("SimHei", StringComparison.OrdinalIgnoreCase) >= 0
                                        || n.FontName.IndexOf("Hei", StringComparison.OrdinalIgnoreCase) >= 0);
                            return ok ? null : "字体期望 " + fn + " 实际 " + n.FontName;
                        });
                }
            }
        }

        /// <summary>
        /// 读回可能是 "1.5" / "exact:12"。
        /// WPP 降级：期望倍数时，若读回 exact:(字号×倍数) 或裸数值≈字号×倍数也算接近
        /// （WPS 有时把定距误报成倍数方言）。
        /// </summary>
        private static bool LineSpacingClose(string actual, string expect, double? fontSizePt = null)
        {
            if (string.IsNullOrEmpty(actual) || string.IsNullOrEmpty(expect)) return false;

            string aRaw = actual.Trim();
            string eRaw = expect.Trim();
            bool aExact = aRaw.StartsWith("exact:", StringComparison.OrdinalIgnoreCase);
            bool eExact = eRaw.StartsWith("exact:", StringComparison.OrdinalIgnoreCase);

            string a = aExact ? aRaw.Substring(6).Trim() : aRaw;
            string e = eExact ? eRaw.Substring(6).Trim() : eRaw;
            if (!TryLeadingDouble(e, out double want)) return false;
            if (!TryLeadingDouble(a, out double got)) return false;

            if (aExact == eExact)
            {
                if (string.Equals(a, e, StringComparison.OrdinalIgnoreCase)) return true;
                if (Math.Abs(got - want) <= 0.08) return true;
            }

            // WPP：倍数降级定距后，读回 exact:font×mult，或裸 font×mult
            if (!eExact && fontSizePt.HasValue && fontSizePt.Value >= 1.0)
            {
                double expectExact = fontSizePt.Value * want;
                double tol = Math.Max(1.0, fontSizePt.Value * 0.25);
                if (Math.Abs(got - expectExact) <= tol) return true;
            }

            return false;
        }

        private static bool TryLeadingDouble(string s, out double value)
        {
            value = 0;
            if (string.IsNullOrEmpty(s)) return false;
            string dig = s.Trim();
            if (dig.EndsWith("pt", StringComparison.OrdinalIgnoreCase))
            {
                dig = dig.Substring(0, dig.Length - 2).Trim();
            }

            int i = 0;
            while (i < dig.Length
                && (char.IsDigit(dig[i]) || dig[i] == '.' || dig[i] == ','))
            {
                i++;
            }

            if (i == 0) return false;
            string num = dig.Substring(0, i).Replace(',', '.');
            return double.TryParse(num, NumberStyles.Float, CultureInfo.InvariantCulture, out value);
        }

        private static string FillEq(PptHtmlShapeNode n, string expect)
        {
            if (string.Equals(expect, "none", StringComparison.OrdinalIgnoreCase))
            {
                string fill = (n.Fill ?? "").Trim();
                return string.Equals(fill, "none", StringComparison.OrdinalIgnoreCase)
                    || string.IsNullOrEmpty(fill)
                    ? null
                    : "fill 期望 none 实际 " + n.Fill;
            }

            return HexEq(n.Fill, expect) ? null : "fill 期望 " + expect + " 实际 " + n.Fill;
        }

        private static string TbGeoEq(
            PptHtmlShapeNode n,
            double left,
            double top,
            double width,
            double height,
            double eps = 0.8)
        {
            if (n == null) return "无 textbox";
            if (!ContentAssert.TryParseGeo(n.Style, out double l, out double t, out double w, out double h))
            {
                return "无几何 style";
            }

            if (Math.Abs(l - left) > eps || Math.Abs(t - top) > eps
                || Math.Abs(w - width) > eps || Math.Abs(h - height) > eps)
            {
                return "几何偏离: " + n.Style;
            }

            return null;
        }

        private static string Pt(double v)
        {
            return v.ToString("0.##", CultureInfo.InvariantCulture) + "pt";
        }
    }
}
