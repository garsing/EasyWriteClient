using System;
using System.Collections.Generic;
using System.IO;
using WordAddIn1.PresentationHost;

namespace PptHtmlContentRoundtripTest
{
    /// <summary>同页 chart + 文本框/表/图。前缀 mixc-，50 条 × 两宿主 = 100。</summary>
    internal static partial class ContentCatalog
    {
        private static readonly string[] MixcPieCats = { "A", "B", "C" };
        private static readonly string[] MixcPieVals = { "40", "35", "25" };
        private static readonly string[] MixcPieCats2 = { "X", "Y", "Z" };
        private static readonly string[] MixcPieVals2 = { "10", "20", "70" };

        private static readonly string[][] MixcPageGrid =
        {
            new[] { "品类", "销量" },
            new[] { "页甲", "12" },
            new[] { "页乙", "9" }
        };

        private static readonly string[][] MixcPageGrid2 =
        {
            new[] { "品类", "销量" },
            new[] { "页丙", "7" },
            new[] { "页丁", "5" }
        };

        private static void AddMixChartCases(List<ContentCase> list, ref int page)
        {
            AddMixcCreate(list, ref page);
            AddMixcDirty(list, ref page);
            AddMixcLean(list, ref page);
            AddMixcRead(list, ref page);
            AddMixcGeoNeg(list, ref page);
        }

        private static void AddMixcCreate(List<ContentCase> list, ref int page)
        {
            Add(list, ref page, "mixc-create-tb-tbl-pie",
                a => ContentHtml.MixTitleTablePie("混排带饼", MixcPageGrid),
                null,
                ctx => MixcTitleTableChartOk(ctx, "混排带饼", MixcPageGrid, MixcPieCats, MixcPieVals, "pie"));

            Add(list, ref page, "mixc-create-tb-pic-pie",
                a => ContentHtml.MixTitlePicPie("图旁饼", a.RedPng),
                null,
                ctx => MixcTitlePicChartOk(ctx, "图旁饼", MixcPieCats, MixcPieVals, "pie"));

            Add(list, ref page, "mixc-create-tb-tbl-pic-pie",
                a => ContentHtml.MixTitleTablePicPie("四件同页", MixcPageGrid, a.RedPng),
                null,
                ctx => MixcFourOk(ctx, "四件同页", MixcPageGrid, MixcPieCats, MixcPieVals));

            Add(list, ref page, "mixc-create-tb-pie",
                a => ContentHtml.MixParts(
                    ContentHtml.MixTitleMarkup("只有标题饼"),
                    ContentHtml.ChartFrag("pie2d", 8, 22, 50, 50)),
                null,
                ctx =>
                {
                    string e = TextEq(PreferCreated(ctx, "textbox"), "只有标题饼");
                    return e ?? ContentAssert.ChartDataMismatch(
                        PreferCreated(ctx, "chart"), MixcPieCats, MixcPieVals, "pie");
                });

            Add(list, ref page, "mixc-create-tbl-pie",
                a => ContentHtml.MixParts(
                    ContentHtml.MixPageTableMarkup(MixcPageGrid),
                    ContentHtml.ChartFrag("pie2d", 52, 22, 42, 50)),
                null,
                ctx => MixcTableChartOk(ctx, MixcPageGrid, MixcPieCats, MixcPieVals, "pie"));

            Add(list, ref page, "mixc-create-pic-pie",
                a => ContentHtml.MixParts(
                    ContentHtml.MixPicMarkup(a.RedPng),
                    ContentHtml.ChartFrag("pie2d", 48, 22, 46, 50)),
                null,
                ctx =>
                {
                    if (PreferCreated(ctx, "picture") == null) return "缺 picture";
                    return ContentAssert.ChartDataMismatch(
                        PreferCreated(ctx, "chart"), MixcPieCats, MixcPieVals, "pie");
                });

            Add(list, ref page, "mixc-create-tb-tbl-col",
                a => ContentHtml.MixParts(
                    ContentHtml.MixTitleMarkup("标题柱"),
                    ContentHtml.MixPageTableMarkup(MixcPageGrid),
                    ContentHtml.ChartFrag("column", 52, 22, 42, 50)),
                null,
                ctx => MixcTitleTableChartOk(ctx, "标题柱", MixcPageGrid, MixcPieCats, MixcPieVals, "column"));

            Add(list, ref page, "mixc-create-tb-pic-col",
                a => ContentHtml.MixParts(
                    ContentHtml.MixTitleMarkup("图旁柱"),
                    ContentHtml.MixPicMarkup(a.RedPng),
                    ContentHtml.ChartFrag("column", 48, 22, 46, 50)),
                null,
                ctx => MixcTitlePicChartOk(ctx, "图旁柱", MixcPieCats, MixcPieVals, "column"));

            Add(list, ref page, "mixc-create-tb-col",
                a => ContentHtml.MixParts(
                    ContentHtml.MixTitleMarkup("只标题柱"),
                    ContentHtml.ChartFrag("column", 8, 22, 50, 50)),
                null,
                ctx =>
                {
                    string e = TextEq(PreferCreated(ctx, "textbox"), "只标题柱");
                    return e ?? ContentAssert.ChartDataMismatch(
                        PreferCreated(ctx, "chart"), MixcPieCats, MixcPieVals, "column");
                });

            Add(list, ref page, "mixc-create-tbl-col",
                a => ContentHtml.MixParts(
                    ContentHtml.MixPageTableMarkup(MixcPageGrid),
                    ContentHtml.ChartFrag("column", 52, 22, 42, 50)),
                null,
                ctx => MixcTableChartOk(ctx, MixcPageGrid, MixcPieCats, MixcPieVals, "column"));

            Add(list, ref page, "mixc-create-two-pie",
                a => ContentHtml.MixParts(
                    ContentHtml.ChartFrag("pie2d", 6, 18, 40, 55, MixcPieCats, MixcPieVals),
                    ContentHtml.ChartFrag("pie2d", 52, 18, 40, 55, MixcPieCats2, MixcPieVals2)),
                null,
                ctx => MixcTwoChartOk(ctx.AfterCreate, MixcPieCats, MixcPieVals, MixcPieCats2, MixcPieVals2, "pie", "pie"));

            Add(list, ref page, "mixc-create-pie-col",
                a => ContentHtml.MixParts(
                    ContentHtml.ChartFrag("pie2d", 6, 18, 40, 55),
                    ContentHtml.ChartFrag("column", 52, 18, 40, 55, MixcPieCats2, MixcPieVals2)),
                null,
                ctx => MixcTwoChartOk(
                    ctx.AfterCreate, MixcPieCats, MixcPieVals, MixcPieCats2, MixcPieVals2, "pie", "column"));

            Add(list, ref page, "mixc-create-tb-two-pie",
                a => ContentHtml.MixParts(
                    ContentHtml.MixTitleMarkup("双饼标题"),
                    ContentHtml.ChartFrag("pie2d", 6, 22, 40, 50),
                    ContentHtml.ChartFrag("pie2d", 52, 22, 40, 50, MixcPieCats2, MixcPieVals2)),
                null,
                ctx =>
                {
                    string e = TextEq(PreferCreated(ctx, "textbox"), "双饼标题");
                    return e ?? MixcTwoChartOk(
                        ctx.AfterCreate, MixcPieCats, MixcPieVals, MixcPieCats2, MixcPieVals2, "pie", "pie");
                });

            Add(list, ref page, "mixc-create-tb-bar",
                a => ContentHtml.MixParts(
                    ContentHtml.MixTitleMarkup("标题条"),
                    ContentHtml.ChartFrag("bar", 8, 22, 50, 50)),
                null,
                ctx =>
                {
                    string e = TextEq(PreferCreated(ctx, "textbox"), "标题条");
                    return e ?? ContentAssert.ChartDataMismatch(
                        PreferCreated(ctx, "chart"), MixcPieCats, MixcPieVals, "bar");
                });

            Add(list, ref page, "mixc-create-tb-line",
                a => ContentHtml.MixParts(
                    ContentHtml.MixTitleMarkup("标题折"),
                    ContentHtml.ChartFrag("line", 8, 22, 50, 50)),
                null,
                ctx =>
                {
                    string e = TextEq(PreferCreated(ctx, "textbox"), "标题折");
                    return e ?? ContentAssert.ChartDataMismatch(
                        PreferCreated(ctx, "chart"), MixcPieCats, MixcPieVals, "line");
                });

            Add(list, ref page, "mixc-create-tb-tbl-bar",
                a => ContentHtml.MixParts(
                    ContentHtml.MixTitleMarkup("标题表条"),
                    ContentHtml.MixPageTableMarkup(MixcPageGrid),
                    ContentHtml.ChartFrag("bar", 52, 22, 42, 50)),
                null,
                ctx => MixcTitleTableChartOk(ctx, "标题表条", MixcPageGrid, MixcPieCats, MixcPieVals, "bar"));
        }

        private static void AddMixcDirty(List<ContentCase> list, ref int page)
        {
            AddGroup(list, ref page, "mixc-create-after-leaf", s =>
                MixcSeq(
                    s,
                    ContentHtml.MixTitleTable("先铺叶子", MixcPageGrid),
                    ContentHtml.PieChart(52, 22, 42, 50),
                    pageRead => MixcPageTitleTableChart(pageRead, "先铺叶子", MixcPageGrid, MixcPieCats, MixcPieVals, "pie")));

            AddGroup(list, ref page, "mixc-create-after-tb", s =>
                MixcSeq(
                    s,
                    ContentHtml.TextboxCreate("先标题"),
                    ContentHtml.PieChart(8, 28, 50, 50),
                    pageRead =>
                    {
                        string e = TextEq(ContentAssert.FindContentTextbox(pageRead), "先标题");
                        return e ?? ContentAssert.ChartDataMismatch(
                            ContentAssert.FindByType(pageRead, "chart"), MixcPieCats, MixcPieVals, "pie");
                    }));

            AddGroup(list, ref page, "mixc-create-after-pic", s =>
                MixcSeq(
                    s,
                    ContentHtml.PictureCreate(MixcPic(s)),
                    ContentHtml.PieChart(48, 22, 46, 50),
                    pageRead =>
                    {
                        if (ContentAssert.FindByType(pageRead, "picture") == null) return "缺 picture";
                        return ContentAssert.ChartDataMismatch(
                            ContentAssert.FindByType(pageRead, "chart"), MixcPieCats, MixcPieVals, "pie");
                    }));

            AddGroup(list, ref page, "mixc-create-after-tbl", s =>
                MixcSeq(
                    s,
                    ContentHtml.TableCreate(MixcPageGrid),
                    ContentHtml.PieChart(52, 22, 42, 50),
                    pageRead => MixcPageTableChart(pageRead, MixcPageGrid, MixcPieCats, MixcPieVals, "pie")));

            AddGroup(list, ref page, "mixc-create-chart-then-tb", s =>
                MixcSeq(
                    s,
                    ContentHtml.PieChart(8, 28, 50, 50),
                    ContentHtml.TextboxCreate("后标题", 6, 6, 50, 12),
                    pageRead =>
                    {
                        string e = TextEq(ContentAssert.FindContentTextbox(pageRead), "后标题");
                        return e ?? ContentAssert.ChartDataMismatch(
                            ContentAssert.FindByType(pageRead, "chart"), MixcPieCats, MixcPieVals, "pie");
                    }));

            AddGroup(list, ref page, "mixc-create-chart-then-tbl", s =>
                MixcSeq(
                    s,
                    ContentHtml.PieChart(52, 22, 42, 50),
                    ContentHtml.TableCreate(MixcPageGrid, 6, 22, 42, 50),
                    pageRead => MixcPageTableChart(pageRead, MixcPageGrid, MixcPieCats, MixcPieVals, "pie")));

            AddGroup(list, ref page, "mixc-create-chart-then-pic", s =>
                MixcSeq(
                    s,
                    ContentHtml.PieChart(48, 22, 46, 50),
                    ContentHtml.PictureCreate(MixcPic(s), 6, 22, 36, 50),
                    pageRead =>
                    {
                        if (ContentAssert.FindByType(pageRead, "picture") == null) return "缺 picture";
                        return ContentAssert.ChartDataMismatch(
                            ContentAssert.FindByType(pageRead, "chart"), MixcPieCats, MixcPieVals, "pie");
                    }));

            AddGroup(list, ref page, "mixc-create-after-three", s =>
                MixcSeq(
                    s,
                    ContentHtml.MixTitleTablePic("先三后饼", MixcPageGrid, MixcPic(s)),
                    ContentHtml.PieChart(58, 22, 36, 50),
                    pageRead => MixcPageFour(pageRead, "先三后饼", MixcPageGrid, MixcPieCats, MixcPieVals)));

            AddGroup(list, ref page, "mixc-create-after-col", s =>
                MixcSeq(
                    s,
                    ContentHtml.MixTitleTable("先铺再柱", MixcPageGrid),
                    ContentHtml.ChartCreate("column"),
                    pageRead => MixcPageTitleTableChart(
                        pageRead, "先铺再柱", MixcPageGrid, MixcPieCats, MixcPieVals, "column")));

            AddGroup(list, ref page, "mixc-create-pie-then-col", s =>
                MixcSeq(
                    s,
                    ContentHtml.PieChart(6, 18, 40, 55),
                    ContentHtml.ChartCreate("column", 52, 18, 40, 55, MixcPieCats2, MixcPieVals2),
                    pageRead => MixcTwoChartOk(
                        pageRead, MixcPieCats, MixcPieVals, MixcPieCats2, MixcPieVals2, "pie", "column")));
        }

        private static void AddMixcLean(List<ContentCase> list, ref int page)
        {
            Add(list, ref page, "mixc-lean-title",
                a => ContentHtml.MixTitleTablePie("原标题", MixcPageGrid),
                (a, ids) => ContentHtml.TextboxUpdate(MixcId(ids, 0), "只改标题"),
                ctx =>
                {
                    string e = TextEq(PreferCreated(ctx, "textbox", useReplace: true), "只改标题");
                    if (e != null) return e;
                    string tbl = ContentAssert.TableCellsMismatch(
                        PreferCreated(ctx, "table", useReplace: true), MixcPageGrid);
                    if (tbl != null) return "改标题误伤页表: " + tbl;
                    return ContentAssert.ChartDataMismatch(
                        PreferCreated(ctx, "chart", useReplace: true), MixcPieCats, MixcPieVals, "pie");
                });

            Add(list, ref page, "mixc-lean-table",
                a => ContentHtml.MixTitleTablePie("标题不动", MixcPageGrid),
                (a, ids) => ContentHtml.TableUpdate(MixcId(ids, 1), MixcPageGrid2),
                ctx =>
                {
                    string e = TextEq(PreferCreated(ctx, "textbox", useReplace: true), "标题不动");
                    if (e != null) return "换页表误伤标题: " + e;
                    string tbl = ContentAssert.TableCellsMismatch(
                        PreferCreated(ctx, "table", useReplace: true), MixcPageGrid2);
                    if (tbl != null) return tbl;
                    if (ContentAssert.GridHasText(PreferCreated(ctx, "table", useReplace: true), "A"))
                    {
                        return "页表串进了饼类目";
                    }

                    return ContentAssert.ChartDataMismatch(
                        PreferCreated(ctx, "chart", useReplace: true), MixcPieCats, MixcPieVals, "pie");
                });

            Add(list, ref page, "mixc-lean-chart",
                a => ContentHtml.MixTitleTablePie("标题仍在", MixcPageGrid),
                (a, ids) => ContentHtml.ChartUpdate(MixcId(ids, 2), MixcPieCats2, MixcPieVals2),
                ctx =>
                {
                    string e = TextEq(PreferCreated(ctx, "textbox", useReplace: true), "标题仍在");
                    if (e != null) return "换饼误伤标题: " + e;
                    string tbl = ContentAssert.TableCellsMismatch(
                        PreferCreated(ctx, "table", useReplace: true), MixcPageGrid);
                    if (tbl != null) return "换饼误伤页表: " + tbl;
                    PptHtmlShapeNode chart = PreferCreated(ctx, "chart", useReplace: true);
                    if (ContentAssert.GridHasText(chart, "页甲"))
                    {
                        return "饼内嵌表串进了页表";
                    }

                    return ContentAssert.ChartDataMismatch(chart, MixcPieCats2, MixcPieVals2, "pie");
                });

            Add(list, ref page, "mixc-lean-picture",
                a => ContentHtml.MixTitlePicPie("图旁标题", a.RedPng),
                (a, ids) => ContentHtml.PictureUpdate(MixcId(ids, 1), a.BluePng),
                ctx => MixcLeanPicOk(ctx, "图旁标题"));

            Add(list, ref page, "mixc-lean-title-four",
                a => ContentHtml.MixTitleTablePicPie("四件原题", MixcPageGrid, a.RedPng),
                (a, ids) => ContentHtml.TextboxUpdate(MixcId(ids, 0), "四件新题"),
                ctx => MixcFourOk(ctx, "四件新题", MixcPageGrid, MixcPieCats, MixcPieVals, useReplace: true));

            Add(list, ref page, "mixc-lean-table-four",
                a => ContentHtml.MixTitleTablePicPie("四件换表", MixcPageGrid, a.RedPng),
                (a, ids) => ContentHtml.TableUpdate(MixcId(ids, 1), MixcPageGrid2),
                ctx => MixcFourOk(ctx, "四件换表", MixcPageGrid2, MixcPieCats, MixcPieVals, useReplace: true));

            Add(list, ref page, "mixc-lean-chart-four",
                a => ContentHtml.MixTitleTablePicPie("四件换饼", MixcPageGrid, a.RedPng),
                (a, ids) => ContentHtml.ChartUpdate(MixcId(ids, 3), MixcPieCats2, MixcPieVals2),
                ctx => MixcFourOk(ctx, "四件换饼", MixcPageGrid, MixcPieCats2, MixcPieVals2, useReplace: true));

            Add(list, ref page, "mixc-lean-pic-four",
                a => ContentHtml.MixTitleTablePicPie("四件换图", MixcPageGrid, a.RedPng),
                (a, ids) => ContentHtml.PictureUpdate(MixcId(ids, 2), a.BluePng),
                ctx =>
                {
                    string four = MixcFourOk(
                        ctx, "四件换图", MixcPageGrid, MixcPieCats, MixcPieVals, useReplace: true);
                    return four ?? MixcPicSrcChanged(ctx);
                });

            Add(list, ref page, "mixc-lean-col-data",
                a => ContentHtml.MixParts(
                    ContentHtml.MixTitleMarkup("柱数原"),
                    ContentHtml.MixPageTableMarkup(MixcPageGrid),
                    ContentHtml.ChartFrag("column", 52, 22, 42, 50)),
                (a, ids) => ContentHtml.ChartUpdate(
                    MixcId(ids, 2), MixcPieCats2, MixcPieVals2, "规模", "column"),
                ctx =>
                {
                    string e = TextEq(PreferCreated(ctx, "textbox", useReplace: true), "柱数原");
                    if (e != null) return e;
                    string tbl = ContentAssert.TableCellsMismatch(
                        PreferCreated(ctx, "table", useReplace: true), MixcPageGrid);
                    if (tbl != null) return "换柱误伤页表: " + tbl;
                    return ContentAssert.ChartDataMismatch(
                        PreferCreated(ctx, "chart", useReplace: true), MixcPieCats2, MixcPieVals2, "column");
                });

            Add(list, ref page, "mixc-lean-two-pie-one",
                a => ContentHtml.MixParts(
                    ContentHtml.MixTitleMarkup("双饼瘦"),
                    ContentHtml.ChartFrag("pie2d", 6, 22, 40, 50),
                    ContentHtml.ChartFrag("pie2d", 52, 22, 40, 50, MixcPieCats2, MixcPieVals2)),
                (a, ids) => ContentHtml.ChartUpdate(MixcId(ids, 2), MixcPieCats2, MixcPieVals2),
                ctx =>
                {
                    string e = TextEq(PreferCreated(ctx, "textbox", useReplace: true), "双饼瘦");
                    if (e != null) return e;
                    return MixcTwoChartOk(
                        ctx.AfterReplace,
                        MixcPieCats,
                        MixcPieVals,
                        MixcPieCats2,
                        MixcPieVals2,
                        "pie",
                        "pie");
                });

            Add(list, ref page, "mixc-lean-title-keep-geo",
                a => ContentHtml.MixTitleTablePie("几何标题", MixcPageGrid),
                (a, ids) => ContentHtml.TextboxUpdate(MixcId(ids, 0), "几何已改"),
                ctx =>
                {
                    string e = TextEq(PreferCreated(ctx, "textbox", useReplace: true), "几何已改");
                    if (e != null) return e;
                    return MixcGeoSame(
                        PreferCreated(ctx, "chart"),
                        PreferCreated(ctx, "chart", useReplace: true),
                        "改标题挪了饼");
                });

            Add(list, ref page, "mixc-lean-chart-keep-geo",
                a => ContentHtml.MixTitleTablePie("饼几何", MixcPageGrid),
                (a, ids) => ContentHtml.ChartUpdate(MixcId(ids, 2), MixcPieCats2, MixcPieVals2),
                ctx =>
                {
                    string data = ContentAssert.ChartDataMismatch(
                        PreferCreated(ctx, "chart", useReplace: true), MixcPieCats2, MixcPieVals2, "pie");
                    if (data != null) return data;
                    return MixcGeoSame(
                        PreferCreated(ctx, "chart"),
                        PreferCreated(ctx, "chart", useReplace: true),
                        "换数挪了饼框");
                });
        }

        private static void AddMixcRead(List<ContentCase> list, ref int page)
        {
            Add(list, ref page, "mixc-read-page-count",
                a => ContentHtml.MixTitleTablePie("计数标题", MixcPageGrid),
                null,
                ctx =>
                {
                    if (ContentAssert.CountType(ctx.AfterCreate, "chart") != 1) return "页上 chart 数不是 1";
                    if (ContentAssert.CountType(ctx.AfterCreate, "table") != 1) return "页上 table 数不是 1";
                    return TextEq(PreferCreated(ctx, "textbox"), "计数标题");
                });

            AddGroup(list, ref page, "mixc-read-chart-shape", s =>
            {
                if (!s.TryApply(ContentHtml.MixTitleTablePie("点饼读", MixcPageGrid), true, out List<string> ids, out string err)
                    || ids == null
                    || ids.Count < 3)
                {
                    return "铺底失败: " + err;
                }

                if (!s.TryReadShape(ids[2], out PptHtmlReadResult focused, out err))
                {
                    return "读饼失败: " + err;
                }

                if (ContentAssert.CountType(focused, "table") > 0)
                {
                    return "点饼却读到页表";
                }

                PptHtmlShapeNode chart = ContentAssert.FindByType(focused, "chart")
                    ?? ContentAssert.FindNthByType(focused, "chart", 0);
                if (chart == null && focused?.Shapes != null && focused.Shapes.Count == 1)
                {
                    chart = focused.Shapes[0];
                }

                if (ContentAssert.GridHasText(chart, "页甲"))
                {
                    return "饼详读串进页表";
                }

                return ContentAssert.ChartDataMismatch(chart, MixcPieCats, MixcPieVals, "pie");
            });

            AddGroup(list, ref page, "mixc-read-table-shape", s =>
            {
                if (!s.TryApply(ContentHtml.MixTitleTablePie("点表读", MixcPageGrid), true, out List<string> ids, out string err)
                    || ids == null
                    || ids.Count < 2)
                {
                    return "铺底失败: " + err;
                }

                if (!s.TryReadShape(ids[1], out PptHtmlReadResult focused, out err))
                {
                    return "读表失败: " + err;
                }

                if (ContentAssert.CountType(focused, "chart") > 0)
                {
                    return "点表却读到 chart";
                }

                PptHtmlShapeNode table = ContentAssert.FindByType(focused, "table")
                    ?? (focused?.Shapes != null && focused.Shapes.Count == 1 ? focused.Shapes[0] : null);
                if (ContentAssert.GridHasText(table, "A"))
                {
                    return "页表详读串进饼类目";
                }

                return ContentAssert.TableCellsMismatch(table, MixcPageGrid);
            });

            Add(list, ref page, "mixc-read-two-pie-distinct",
                a => ContentHtml.MixParts(
                    ContentHtml.ChartFrag("pie2d", 6, 18, 40, 55),
                    ContentHtml.ChartFrag("pie2d", 52, 18, 40, 55, MixcPieCats2, MixcPieVals2)),
                null,
                ctx => MixcTwoChartOk(
                    ctx.AfterCreate, MixcPieCats, MixcPieVals, MixcPieCats2, MixcPieVals2, "pie", "pie"));

            AddGroup(list, ref page, "mixc-read-tb-shape", s =>
            {
                if (!s.TryApply(ContentHtml.MixTitleTablePie("点标题", MixcPageGrid), true, out List<string> ids, out string err)
                    || ids == null
                    || ids.Count < 1)
                {
                    return "铺底失败: " + err;
                }

                if (!s.TryReadShape(ids[0], out PptHtmlReadResult focused, out err))
                {
                    return "读标题失败: " + err;
                }

                PptHtmlShapeNode tb = ContentAssert.FindByType(focused, "textbox")
                    ?? (focused?.Shapes != null && focused.Shapes.Count > 0 ? focused.Shapes[0] : null);
                return TextEq(tb, "点标题");
            });

            Add(list, ref page, "mixc-read-four-types",
                a => ContentHtml.MixTitleTablePicPie("四类都在", MixcPageGrid, a.GreenPng),
                null,
                ctx =>
                {
                    if (ContentAssert.CountType(ctx.AfterCreate, "textbox") < 1) return "缺 textbox";
                    if (ContentAssert.CountType(ctx.AfterCreate, "table") != 1) return "table 数不对";
                    if (ContentAssert.CountType(ctx.AfterCreate, "picture") < 1) return "缺 picture";
                    if (ContentAssert.CountType(ctx.AfterCreate, "chart") != 1) return "chart 数不对";
                    return null;
                });

            Add(list, ref page, "mixc-read-chart-not-page-cells",
                a => ContentHtml.MixTitleTablePie("饼不吃表", MixcPageGrid),
                null,
                ctx =>
                {
                    PptHtmlShapeNode chart = PreferCreated(ctx, "chart");
                    if (ContentAssert.GridHasText(chart, "品类") || ContentAssert.GridHasText(chart, "页甲"))
                    {
                        return "饼内嵌表串进了页表";
                    }

                    return ContentAssert.ChartDataMismatch(chart, MixcPieCats, MixcPieVals, "pie");
                });

            Add(list, ref page, "mixc-read-table-not-chart-cats",
                a => ContentHtml.MixTitleTablePie("表不吃饼", MixcPageGrid),
                null,
                ctx =>
                {
                    PptHtmlShapeNode table = PreferCreated(ctx, "table");
                    if (ContentAssert.GridHasText(table, "A") || ContentAssert.GridHasText(table, "份额"))
                    {
                        return "页表串进了饼类目";
                    }

                    return ContentAssert.TableCellsMismatch(table, MixcPageGrid);
                });
        }

        private static void AddMixcGeoNeg(List<ContentCase> list, ref int page)
        {
            Add(list, ref page, "mixc-geo-pie-right",
                a => ContentHtml.MixTitleTablePie("饼在右", MixcPageGrid),
                null,
                ctx =>
                {
                    PptHtmlShapeNode chart = PreferCreated(ctx, "chart");
                    if (chart == null || !ContentAssert.TryParseGeo(chart.Style, out double left, out _, out _, out _))
                    {
                        return "饼无几何";
                    }

                    return left >= 45 ? null : "饼应在右侧 left=" + left.ToString("0.##");
                });

            AddGroup(list, ref page, "mixc-neg-chart-id-as-table", s =>
            {
                if (!s.TryApply(ContentHtml.MixTitleTablePie("误写饼", MixcPageGrid), true, out List<string> ids, out string err)
                    || ids == null
                    || ids.Count < 3)
                {
                    return "铺底失败: " + err;
                }

                s.TryApply(ContentHtml.TableUpdate(ids[2], MixcPageGrid2), false, out _, out _);
                if (!s.TryReadFullPage(out PptHtmlReadResult pageRead, out err))
                {
                    return err;
                }

                string tbl = ContentAssert.TableCellsMismatch(
                    ContentAssert.FindByType(pageRead, "table"), MixcPageGrid);
                if (tbl != null) return "误写饼伤了页表: " + tbl;
                if (ContentAssert.GridHasText(ContentAssert.FindByType(pageRead, "chart"), "页丙"))
                {
                    return "页表写进了饼";
                }

                return ContentAssert.ChartDataMismatch(
                    ContentAssert.FindByType(pageRead, "chart"), MixcPieCats, MixcPieVals, "pie");
            });

            AddGroup(list, ref page, "mixc-neg-table-id-as-chart", s =>
            {
                if (!s.TryApply(ContentHtml.MixTitleTablePie("误写表", MixcPageGrid), true, out List<string> ids, out string err)
                    || ids == null
                    || ids.Count < 3)
                {
                    return "铺底失败: " + err;
                }

                s.TryApply(ContentHtml.ChartUpdate(ids[1], MixcPieCats2, MixcPieVals2), false, out _, out _);
                if (!s.TryReadFullPage(out PptHtmlReadResult pageRead, out err))
                {
                    return err;
                }

                string tbl = ContentAssert.TableCellsMismatch(
                    ContentAssert.FindByType(pageRead, "table"), MixcPageGrid);
                if (tbl != null) return "饼稿写进了页表: " + tbl;
                return ContentAssert.ChartDataMismatch(
                    ContentAssert.FindByType(pageRead, "chart"), MixcPieCats, MixcPieVals, "pie");
            });

            Add(list, ref page, "mixc-neg-empty-chart",
                a => ContentHtml.Section(
                    "  <div data-shape-type=\"chart\" data-chart-type=\"pie2d\" style=\""
                    + ContentHtml.Geo(10, 20, 40, 40)
                    + "\"></div>"),
                null,
                null,
                expectApplyErrorContains: "空图");
        }

        private static string MixcSeq(
            ContentGroupSession s,
            string first,
            string second,
            Func<PptHtmlReadResult, string> match)
        {
            if (!s.TryApply(first, true, out _, out string err))
            {
                return "先铺失败: " + err;
            }

            if (!s.TryApply(second, true, out List<string> ids, out err) || ids == null || ids.Count < 1)
            {
                return "后铺失败: " + err;
            }

            if (!s.TryReadFullPage(out PptHtmlReadResult pageRead, out err))
            {
                return "读页失败: " + err;
            }

            return match(pageRead);
        }

        private static string MixcPic(ContentGroupSession s)
        {
            if (s?.Assets != null && !string.IsNullOrEmpty(s.Assets.RedPng))
            {
                return s.Assets.RedPng;
            }

            return Path.Combine(Path.GetTempPath(), "mixc-missing.png");
        }

        private static string MixcTitleTableChartOk(
            ContentAssertContext ctx,
            string title,
            string[][] pageGrid,
            string[] cats,
            string[] vals,
            string typeHint,
            bool useReplace = false)
        {
            string e = TextEq(PreferCreated(ctx, "textbox", useReplace), title);
            if (e != null) return e;
            return MixcTableChartOk(ctx, pageGrid, cats, vals, typeHint, useReplace);
        }

        private static string MixcTitlePicChartOk(
            ContentAssertContext ctx,
            string title,
            string[] cats,
            string[] vals,
            string typeHint)
        {
            string e = TextEq(PreferCreated(ctx, "textbox"), title);
            if (e != null) return e;
            if (PreferCreated(ctx, "picture") == null) return "缺 picture";
            return ContentAssert.ChartDataMismatch(PreferCreated(ctx, "chart"), cats, vals, typeHint);
        }

        private static string MixcTableChartOk(
            ContentAssertContext ctx,
            string[][] pageGrid,
            string[] cats,
            string[] vals,
            string typeHint,
            bool useReplace = false)
        {
            return MixcPageTableChart(
                useReplace ? ctx.AfterReplace : ctx.AfterCreate,
                pageGrid,
                cats,
                vals,
                typeHint,
                PreferCreated(ctx, "table", useReplace),
                PreferCreated(ctx, "chart", useReplace));
        }

        private static string MixcPageTitleTableChart(
            PptHtmlReadResult page,
            string title,
            string[][] grid,
            string[] cats,
            string[] vals,
            string typeHint)
        {
            string e = TextEq(ContentAssert.FindContentTextbox(page), title);
            if (e != null) return e;
            return MixcPageTableChart(page, grid, cats, vals, typeHint);
        }

        private static string MixcPageTableChart(
            PptHtmlReadResult page,
            string[][] grid,
            string[] cats,
            string[] vals,
            string typeHint,
            PptHtmlShapeNode table = null,
            PptHtmlShapeNode chart = null)
        {
            table = table ?? ContentAssert.FindByType(page, "table");
            chart = chart ?? ContentAssert.FindByType(page, "chart");
            string tbl = ContentAssert.TableCellsMismatch(table, grid);
            if (tbl != null) return tbl;
            if (cats != null && cats.Length > 0 && ContentAssert.GridHasText(table, cats[0]))
            {
                return "页表串进了图类目";
            }

            if (ContentAssert.GridHasText(chart, "品类") || ContentAssert.GridHasText(chart, "页甲"))
            {
                return "图内嵌表串进了页表";
            }

            // 柱/条/折系列读回类目常是 1/2/3，混排套对数值和类型即可。
            bool cartesian = !string.IsNullOrEmpty(typeHint)
                && typeHint.IndexOf("pie", StringComparison.OrdinalIgnoreCase) < 0;
            return ContentAssert.ChartDataMismatch(
                chart,
                cartesian ? null : cats,
                vals,
                typeHint);
        }

        private static string MixcFourOk(
            ContentAssertContext ctx,
            string title,
            string[][] grid,
            string[] cats,
            string[] vals,
            bool useReplace = false)
        {
            string e = MixcTitleTableChartOk(ctx, title, grid, cats, vals, "pie", useReplace);
            if (e != null) return e;
            if (PreferCreated(ctx, "picture", useReplace) == null) return "缺 picture";
            return null;
        }

        private static string MixcPageFour(
            PptHtmlReadResult page,
            string title,
            string[][] grid,
            string[] cats,
            string[] vals)
        {
            if (ContentAssert.FindByType(page, "picture") == null) return "缺 picture";
            return MixcPageTitleTableChart(page, title, grid, cats, vals, "pie");
        }

        private static string MixcTwoChartOk(
            PptHtmlReadResult page,
            string[] cats1,
            string[] vals1,
            string[] cats2,
            string[] vals2,
            string type1,
            string type2)
        {
            if (ContentAssert.CountType(page, "chart") < 2)
            {
                return "页上 chart 不足 2";
            }

            string a = ContentAssert.ChartDataMismatch(
                ContentAssert.FindNthByType(page, "chart", 0), cats1, vals1, type1);
            if (a != null) return "图1: " + a;
            string b = ContentAssert.ChartDataMismatch(
                ContentAssert.FindNthByType(page, "chart", 1), cats2, vals2, type2);
            return b == null ? null : "图2: " + b;
        }

        private static string MixcLeanPicOk(ContentAssertContext ctx, string title)
        {
            string e = TextEq(PreferCreated(ctx, "textbox", useReplace: true), title);
            if (e != null) return "换图误伤标题: " + e;
            string chart = ContentAssert.ChartDataMismatch(
                PreferCreated(ctx, "chart", useReplace: true), MixcPieCats, MixcPieVals, "pie");
            if (chart != null) return "换图误伤图: " + chart;
            return MixcPicSrcChanged(ctx);
        }

        private static string MixcPicSrcChanged(ContentAssertContext ctx)
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
                ? "混排换图 src 未变"
                : null;
        }

        private static string MixcGeoSame(PptHtmlShapeNode before, PptHtmlShapeNode after, string fail)
        {
            if (before == null || after == null) return "缺 chart 几何";
            return ContentAssert.GeoClose(before.Style, after.Style, 1.2) ? null : fail + " " + after.Style;
        }

        private static string MixcId(IList<string> ids, int index)
        {
            if (ids == null || index < 0 || index >= ids.Count)
            {
                return null;
            }

            return ids[index];
        }
    }
}
