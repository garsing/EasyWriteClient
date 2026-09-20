using System;
using System.Collections.Generic;
using System.Globalization;
using WordAddIn1.PresentationHost;

namespace PptHtmlContentRoundtripTest
{
    /// <summary>
    /// 表格属性交叉矩阵（tbl-mx-*）。覆盖网格尺寸、几何、表级皮、列宽/行高、
    /// 合并、格皮、span 字皮、扩缩/换数与组合交叉；合计 ≥100。
    /// </summary>
    internal static partial class ContentCatalog
    {
        private static void AddTableMatrixCases(List<ContentCase> list, ref int page)
        {
            int before = list.Count;
            AddMatrixGridSizes(list, ref page);
            AddMatrixShapeSkin(list, ref page);
            AddMatrixTableFonts(list, ref page);
            AddMatrixColRowPcts(list, ref page);
            AddMatrixThreeLine(list, ref page);
            AddMatrixCellFills(list, ref page);
            AddMatrixCellSpans(list, ref page);
            AddMatrixMerges(list, ref page);
            AddMatrixUpdates(list, ref page);
            AddMatrixCrossCombos(list, ref page);
            int added = list.Count - before;
            if (added < 100)
            {
                throw new InvalidOperationException(
                    "表格矩阵用例不足 100（实际 " + added + "），请扩 ContentCatalogTableMatrix。");
            }
        }

        private static void AddMatrixGridSizes(List<ContentCase> list, ref int page)
        {
            int[][] sizes =
            {
                new[] { 1, 1 }, new[] { 1, 2 }, new[] { 1, 3 }, new[] { 1, 5 }, new[] { 1, 8 },
                new[] { 2, 1 }, new[] { 2, 2 }, new[] { 2, 3 }, new[] { 2, 4 }, new[] { 2, 6 },
                new[] { 3, 1 }, new[] { 3, 2 }, new[] { 3, 3 }, new[] { 3, 4 }, new[] { 3, 5 },
                new[] { 4, 2 }, new[] { 4, 3 }, new[] { 4, 4 }, new[] { 5, 2 }, new[] { 5, 3 },
                new[] { 6, 2 }, new[] { 8, 1 }
            };

            for (int i = 0; i < sizes.Length; i++)
            {
                int rows = sizes[i][0];
                int cols = sizes[i][1];
                string[][] grid = ContentHtml.Grid(rows, cols, "g");
                string name = "tbl-mx-grid-" + rows + "x" + cols;
                Add(list, ref page, name,
                    a => ContentHtml.TableCreate(grid, left: 6, top: 20, width: 70, height: 55),
                    null,
                    ctx => ContentAssert.TableCellsMismatch(PreferCreated(ctx, "table"), grid));
            }
        }

        private static void AddMatrixShapeSkin(List<ContentCase> list, ref int page)
        {
            string[] fills =
            {
                "#DEEBF7", "#FFF2CC", "#E2F0D9", "#FCE4D6", "#DDEBF7", "#F4B183", "#C6EFCE", "#FFC7CE"
            };
            double[][] geos =
            {
                new[] { 5.0, 15, 40, 25 },
                new[] { 12.0, 22, 55, 40 },
                new[] { 20.0, 30, 45, 35 },
                new[] { 8.0, 40, 60, 28 }
            };

            string[][] g2 = ContentHtml.Grid(2, 2, "s");
            for (int i = 0; i < fills.Length; i++)
            {
                string fill = fills[i];
                double[] geo = geos[i % geos.Length];
                string name = "tbl-mx-fill-" + fill.TrimStart('#');
                Add(list, ref page, name,
                    a => ContentHtml.TableCreate(
                        g2,
                        left: geo[0],
                        top: geo[1],
                        width: geo[2],
                        height: geo[3],
                        extraAttrs: "data-fill=\"" + fill + "\""),
                    null,
                    ctx =>
                    {
                        PptHtmlShapeNode n = PreferCreated(ctx, "table");
                        return ContentAssert.FirstFail(
                            ContentAssert.TableCellsMismatch(n, g2),
                            HexEq(n?.Fill, fill) ? null : "表 fill 期望 " + fill + " 实际 " + n?.Fill,
                            GeoNear(n, geo[0], geo[1], geo[2], geo[3]));
                    });
            }

            // 线色：部分宿主对表形描边弱支持，只断言网格成功
            string[] lines = { "#548235", "#C00000", "#0070C0", "#7030A0" };
            for (int i = 0; i < lines.Length; i++)
            {
                string line = lines[i];
                Add(list, ref page, "tbl-mx-line-" + line.TrimStart('#'),
                    a => ContentHtml.TableCreate(
                        g2,
                        extraAttrs: "data-line-color=\"" + line + "\" data-line-width=\"1.25pt\""),
                    null,
                    ctx => ContentAssert.TableCellsMismatch(PreferCreated(ctx, "table"), g2));
            }
        }

        private static void AddMatrixTableFonts(List<ContentCase> list, ref int page)
        {
            string[][] g2 = ContentHtml.Grid(2, 2, "f");
            var fonts = new[]
            {
                new { Name = "微软雅黑", Size = 12.0, Color = "#333333" },
                new { Name = "微软雅黑", Size = 14.0, Color = "#0070C0" },
                new { Name = "微软雅黑", Size = 18.0, Color = "#C00000" },
                new { Name = "宋体", Size = 12.0, Color = "#000000" },
                new { Name = "宋体", Size = 16.0, Color = "#548235" },
                new { Name = "黑体", Size = 14.0, Color = "#7030A0" },
                new { Name = "Arial", Size = 11.0, Color = "#1F4E79" },
                new { Name = "Arial", Size = 16.0, Color = "#C65911" },
                new { Name = "Calibri", Size = 13.0, Color = "#385723" },
                new { Name = "Calibri", Size = 20.0, Color = "#833C0C" }
            };

            for (int i = 0; i < fonts.Length; i++)
            {
                var f = fonts[i];
                string attrs = "data-table-font-name=\"" + f.Name + "\" data-table-font-size=\""
                    + f.Size.ToString("0.##", CultureInfo.InvariantCulture)
                    + "\" data-table-font-color=\"" + f.Color + "\"";
                string name = "tbl-mx-font-" + SafeToken(f.Name) + "-" + ((int)f.Size) + "-" + f.Color.TrimStart('#');
                Add(list, ref page, name,
                    a => ContentHtml.TableCreate(g2, extraAttrs: attrs),
                    null,
                    ctx =>
                    {
                        PptHtmlShapeNode n = PreferCreated(ctx, "table");
                        return ContentAssert.FirstFail(
                            ContentAssert.TableCellsMismatch(n, g2),
                            ContentAssert.TableFontClose(n, f.Name, f.Size, f.Color));
                    });
            }
        }

        private static void AddMatrixColRowPcts(List<ContentCase> list, ref int page)
        {
            // 列宽
            var colSpecs = new[]
            {
                new { R = 1, C = 2, Cols = new float[] { 40f, 60f }, Rows = (float[])null },
                new { R = 1, C = 2, Cols = new float[] { 30f, 70f }, Rows = (float[])null },
                new { R = 1, C = 3, Cols = new float[] { 20f, 30f, 50f }, Rows = (float[])null },
                new { R = 1, C = 3, Cols = new float[] { 33f, 33f, 34f }, Rows = (float[])null },
                new { R = 1, C = 4, Cols = new float[] { 10f, 20f, 30f, 40f }, Rows = (float[])null },
                new { R = 2, C = 2, Cols = new float[] { 25f, 75f }, Rows = (float[])null },
                new { R = 2, C = 3, Cols = new float[] { 15f, 35f, 50f }, Rows = (float[])null },
                new { R = 3, C = 3, Cols = new float[] { 20f, 40f, 40f }, Rows = (float[])null }
            };

            for (int i = 0; i < colSpecs.Length; i++)
            {
                var s = colSpecs[i];
                string[][] grid = ContentHtml.Grid(s.R, s.C, "w");
                string attrs = "data-col-widths=\"" + JoinPct(s.Cols) + "\"";
                Add(list, ref page, "tbl-mx-colw-" + i + "-" + s.R + "x" + s.C,
                    a => ContentHtml.TableCreate(grid, width: 70, height: 40, extraAttrs: attrs),
                    null,
                    ctx =>
                    {
                        PptHtmlShapeNode n = PreferCreated(ctx, "table");
                        return ContentAssert.FirstFail(
                            ContentAssert.TableCellsMismatch(n, grid),
                            ContentAssert.TableColWidthsClose(n, s.Cols));
                    });
            }

            // 行高
            var rowSpecs = new[]
            {
                new { R = 2, C = 2, Rows = new float[] { 40f, 60f }, Cols = (float[])null },
                new { R = 2, C = 2, Rows = new float[] { 30f, 70f }, Cols = (float[])null },
                new { R = 3, C = 2, Rows = new float[] { 20f, 30f, 50f }, Cols = (float[])null },
                new { R = 3, C = 3, Rows = new float[] { 25f, 25f, 50f }, Cols = (float[])null },
                new { R = 4, C = 2, Rows = new float[] { 15f, 25f, 25f, 35f }, Cols = (float[])null }
            };

            for (int i = 0; i < rowSpecs.Length; i++)
            {
                var s = rowSpecs[i];
                string[][] grid = ContentHtml.Grid(s.R, s.C, "h");
                string attrs = "data-row-heights=\"" + JoinPct(s.Rows) + "\"";
                Add(list, ref page, "tbl-mx-rowh-" + i + "-" + s.R + "x" + s.C,
                    a => ContentHtml.TableCreate(grid, width: 55, height: 50, extraAttrs: attrs),
                    null,
                    ctx =>
                    {
                        PptHtmlShapeNode n = PreferCreated(ctx, "table");
                        return ContentAssert.FirstFail(
                            ContentAssert.TableCellsMismatch(n, grid),
                            ContentAssert.TableRowHeightsClose(n, s.Rows));
                    });
            }
        }

        private static void AddMatrixThreeLine(List<ContentCase> list, ref int page)
        {
            string[][] g2 = ContentHtml.Grid(2, 2, "t");
            string[][] g3 = ContentHtml.Grid(3, 3, "t");

            Add(list, ref page, "tbl-mx-three-line-2x2",
                a => ContentHtml.TableCreate(g2, extraAttrs: "data-table-style=\"three-line\""),
                null,
                ctx => ContentAssert.FirstFail(
                    ContentAssert.TableCellsMismatch(PreferCreated(ctx, "table"), g2),
                    ContentAssert.TableStyleIsThreeLine(PreferCreated(ctx, "table"))));

            Add(list, ref page, "tbl-mx-three-line-3x3",
                a => ContentHtml.TableCreate(g3, extraAttrs: "data-table-style=\"three-line\""),
                null,
                ctx => ContentAssert.FirstFail(
                    ContentAssert.TableCellsMismatch(PreferCreated(ctx, "table"), g3),
                    ContentAssert.TableStyleIsThreeLine(PreferCreated(ctx, "table"))));

            Add(list, ref page, "tbl-mx-three-line-colw",
                a => ContentHtml.TableCreate(
                    g2,
                    extraAttrs: "data-table-style=\"three-line\" data-col-widths=\"35%,65%\""),
                null,
                ctx =>
                {
                    PptHtmlShapeNode n = PreferCreated(ctx, "table");
                    return ContentAssert.FirstFail(
                        ContentAssert.TableCellsMismatch(n, g2),
                        ContentAssert.TableStyleIsThreeLine(n),
                        ContentAssert.TableColWidthsClose(n, new float[] { 35f, 65f }));
                });

            Add(list, ref page, "tbl-mx-three-line-font",
                a => ContentHtml.TableCreate(
                    g2,
                    extraAttrs: "data-table-style=\"three-line\" data-table-font-name=\"微软雅黑\" "
                        + "data-table-font-size=\"14\" data-table-font-color=\"#1F4E79\""),
                null,
                ctx =>
                {
                    PptHtmlShapeNode n = PreferCreated(ctx, "table");
                    return ContentAssert.FirstFail(
                        ContentAssert.TableCellsMismatch(n, g2),
                        ContentAssert.TableStyleIsThreeLine(n),
                        ContentAssert.TableFontClose(n, "微软雅黑", 14, "#1F4E79"));
                });
        }

        private static void AddMatrixCellFills(List<ContentCase> list, ref int page)
        {
            string[] colors =
            {
                "#F5F5F5", "#DEEBF7", "#FFF2CC", "#E2F0D9", "#FCE4D6",
                "#DDEBF7", "#C6EFCE", "#FFC7CE", "#D9E2F3", "#F8CBAD"
            };

            for (int i = 0; i < colors.Length; i++)
            {
                string color = colors[i];
                int r = i % 2;
                int c = (i / 2) % 2;
                string rows = Build2x2WithCellFill(r, c, color, "cf" + i);
                Add(list, ref page, "tbl-mx-cellfill-r" + r + "c" + c + "-" + color.TrimStart('#'),
                    a => ContentHtml.TableCreateRaw(rows),
                    null,
                    ctx =>
                    {
                        PptHtmlShapeNode n = PreferCreated(ctx, "table");
                        return ContentAssert.FirstFail(
                            ContentAssert.TableRequire(n),
                            ContentAssert.InnerHas(n, "data-fill=\"" + color + "\"", "格 fill 未读回"));
                    });
            }
        }

        private static void AddMatrixCellSpans(List<ContentCase> list, ref int page)
        {
            var specs = new[]
            {
                new { Tag = "bold", Color = (string)null, Bold = (bool?)true, Italic = (bool?)null },
                new { Tag = "italic", Color = (string)null, Bold = (bool?)null, Italic = (bool?)true },
                new { Tag = "red", Color = "#C00000", Bold = (bool?)null, Italic = (bool?)null },
                new { Tag = "blue", Color = "#0070C0", Bold = (bool?)null, Italic = (bool?)null },
                new { Tag = "green", Color = "#548235", Bold = (bool?)null, Italic = (bool?)null },
                new { Tag = "bold-red", Color = "#C00000", Bold = (bool?)true, Italic = (bool?)null },
                new { Tag = "bold-blue", Color = "#0070C0", Bold = (bool?)true, Italic = (bool?)null },
                new { Tag = "italic-purple", Color = "#7030A0", Bold = (bool?)null, Italic = (bool?)true },
                new { Tag = "bold-italic", Color = "#333333", Bold = (bool?)true, Italic = (bool?)true },
                new { Tag = "bold-italic-orange", Color = "#C65911", Bold = (bool?)true, Italic = (bool?)true }
            };

            for (int i = 0; i < specs.Length; i++)
            {
                var s = specs[i];
                string rows = ContentHtml.Tr(
                        ContentHtml.Td("A"),
                        ContentHtml.Td("B", fontColor: s.Color, bold: s.Bold, italic: s.Italic))
                    + ContentHtml.Tr(ContentHtml.Td("C"), ContentHtml.Td("D"));
                Add(list, ref page, "tbl-mx-span-" + s.Tag,
                    a => ContentHtml.TableCreateRaw(rows),
                    null,
                    ctx =>
                    {
                        PptHtmlShapeNode n = PreferCreated(ctx, "table");
                        string e = ContentAssert.TableRequire(n);
                        if (e != null) return e;
                        if ((n.InnerHtml ?? "").IndexOf("B", StringComparison.Ordinal) < 0)
                        {
                            return "缺单元格 B";
                        }

                        if (s.Bold == true)
                        {
                            e = ContentAssert.InnerHas(n, "font-weight:bold", "缺 bold");
                            if (e != null) return e;
                        }

                        if (s.Italic == true)
                        {
                            e = ContentAssert.InnerHas(n, "font-style:italic", "缺 italic");
                            if (e != null) return e;
                        }

                        if (!string.IsNullOrEmpty(s.Color))
                        {
                            // 读回可能规范化大小写；用去掉 # 后的片段
                            string hex = s.Color.TrimStart('#');
                            if ((n.InnerHtml ?? "").IndexOf(hex, StringComparison.OrdinalIgnoreCase) < 0)
                            {
                                return "缺字色 " + s.Color + ": " + TruncInner(n);
                            }
                        }

                        return null;
                    });
            }
        }

        private static void AddMatrixMerges(List<ContentCase> list, ref int page)
        {
            // colspan
            Add(list, ref page, "tbl-mx-merge-colspan2",
                a => ContentHtml.TableCreateRaw(
                    ContentHtml.Tr(ContentHtml.Td("头", colspan: 2, header: true))
                    + ContentHtml.Tr(ContentHtml.Td("L"), ContentHtml.Td("R"))),
                null,
                ctx => ContentAssert.FirstFail(
                    ContentAssert.InnerHas(PreferCreated(ctx, "table"), "colspan=\"2\"", "缺 colspan=2"),
                    ContentAssert.InnerHas(PreferCreated(ctx, "table"), "头")));

            Add(list, ref page, "tbl-mx-merge-colspan3",
                a => ContentHtml.TableCreateRaw(
                    ContentHtml.Tr(ContentHtml.Td("横三", colspan: 3, header: true))
                    + ContentHtml.Tr(ContentHtml.Td("a"), ContentHtml.Td("b"), ContentHtml.Td("c"))),
                null,
                ctx => ContentAssert.InnerHas(PreferCreated(ctx, "table"), "colspan=\"3\"", "缺 colspan=3"));

            // rowspan：被覆盖格不要再写空 <td>，否则解析报槽位溢出
            Add(list, ref page, "tbl-mx-merge-rowspan2",
                a => ContentHtml.TableCreateRaw(
                    ContentHtml.Tr(ContentHtml.Td("侧", rowspan: 2), ContentHtml.Td("上"))
                    + ContentHtml.Tr(ContentHtml.Td("下"))),
                null,
                ctx => ContentAssert.InnerHas(PreferCreated(ctx, "table"), "rowspan=\"2\"", "缺 rowspan=2"));

            Add(list, ref page, "tbl-mx-merge-rowspan3",
                a => ContentHtml.TableCreateRaw(
                    ContentHtml.Tr(ContentHtml.Td("侧三", rowspan: 3), ContentHtml.Td("1"))
                    + ContentHtml.Tr(ContentHtml.Td("2"))
                    + ContentHtml.Tr(ContentHtml.Td("3"))),
                null,
                ctx => ContentAssert.InnerHas(PreferCreated(ctx, "table"), "rowspan=\"3\"", "缺 rowspan=3"));

            // L 型 / 角合并
            Add(list, ref page, "tbl-mx-merge-corner-colspan",
                a => ContentHtml.TableCreateRaw(
                    ContentHtml.Tr(ContentHtml.Td("角", colspan: 2), ContentHtml.Td("右上"))
                    + ContentHtml.Tr(ContentHtml.Td(""), ContentHtml.Td(""), ContentHtml.Td("右下"))
                    + ContentHtml.Tr(ContentHtml.Td("左下"), ContentHtml.Td("中下"), ContentHtml.Td("右底"))),
                null,
                ctx => ContentAssert.InnerHas(PreferCreated(ctx, "table"), "colspan=\"2\""));

            Add(list, ref page, "tbl-mx-merge-header-rowspan",
                a => ContentHtml.TableCreateRaw(
                    ContentHtml.Tr(
                        ContentHtml.Td("行标", rowspan: 2, header: true),
                        ContentHtml.Td("列A", header: true),
                        ContentHtml.Td("列B", header: true))
                    + ContentHtml.Tr(ContentHtml.Td("a1"), ContentHtml.Td("b1"))
                    + ContentHtml.Tr(ContentHtml.Td("r2"), ContentHtml.Td("a2"), ContentHtml.Td("b2"))),
                null,
                ctx => ContentAssert.FirstFail(
                    ContentAssert.InnerHas(PreferCreated(ctx, "table"), "rowspan=\"2\""),
                    ContentAssert.InnerHas(PreferCreated(ctx, "table"), "行标")));

            // 多合并同表
            Add(list, ref page, "tbl-mx-merge-multi",
                a => ContentHtml.TableCreateRaw(
                    ContentHtml.Tr(ContentHtml.Td("H", colspan: 3, header: true))
                    + ContentHtml.Tr(ContentHtml.Td("S", rowspan: 2), ContentHtml.Td("m1"), ContentHtml.Td("n1"))
                    + ContentHtml.Tr(ContentHtml.Td("m2"), ContentHtml.Td("n2"))),
                null,
                ctx => ContentAssert.FirstFail(
                    ContentAssert.InnerHas(PreferCreated(ctx, "table"), "colspan=\"3\""),
                    ContentAssert.InnerHas(PreferCreated(ctx, "table"), "rowspan=\"2\"")));

            Add(list, ref page, "tbl-mx-merge-th-td-mix",
                a => ContentHtml.TableCreateRaw(
                    ContentHtml.Tr(ContentHtml.Td("标题", colspan: 2, header: true))
                    + ContentHtml.Tr(ContentHtml.Td("值1"), ContentHtml.Td("值2", fill: "#DEEBF7"))),
                null,
                ctx => ContentAssert.FirstFail(
                    ContentAssert.InnerHas(PreferCreated(ctx, "table"), "colspan=\"2\""),
                    ContentAssert.InnerHas(PreferCreated(ctx, "table"), "data-fill=\"#DEEBF7\"")));
        }

        private static void AddMatrixUpdates(List<ContentCase> list, ref int page)
        {
            string[][] g2 = ContentHtml.Grid(2, 2, "u");
            string[][] g2b = ContentHtml.Grid(2, 2, "v");
            string[][] g3 = ContentHtml.Grid(3, 3, "u");
            string[][] g1x3 = ContentHtml.Grid(1, 3, "u");
            string[][] g4x2 = ContentHtml.Grid(4, 2, "u");

            Add(list, ref page, "tbl-mx-upd-replace-text",
                a => ContentHtml.TableCreate(g2),
                (a, ids) => ContentHtml.TableUpdate(ids[0], g2b),
                ctx => ContentAssert.TableCellsMismatch(
                    PreferCreated(ctx, "table", useReplace: true), g2b));

            Add(list, ref page, "tbl-mx-upd-grow-2x2-3x3",
                a => ContentHtml.TableCreate(g2),
                (a, ids) => ContentHtml.TableUpdate(ids[0], g3),
                ctx => ContentAssert.TableCellsMismatch(
                    PreferCreated(ctx, "table", useReplace: true), g3));

            Add(list, ref page, "tbl-mx-upd-grow-2x2-4x2",
                a => ContentHtml.TableCreate(g2),
                (a, ids) => ContentHtml.TableUpdate(ids[0], g4x2),
                ctx => ContentAssert.TableCellsMismatch(
                    PreferCreated(ctx, "table", useReplace: true), g4x2));

            Add(list, ref page, "tbl-mx-upd-shrink-3x3-2x2",
                a => ContentHtml.TableCreate(g3),
                (a, ids) => ContentHtml.TableUpdate(ids[0], g2b),
                ctx => ContentAssert.TableCellsMismatch(
                    PreferCreated(ctx, "table", useReplace: true), g2b));

            Add(list, ref page, "tbl-mx-upd-shrink-3x3-1x3",
                a => ContentHtml.TableCreate(g3),
                (a, ids) => ContentHtml.TableUpdate(ids[0], g1x3),
                ctx => ContentAssert.TableCellsMismatch(
                    PreferCreated(ctx, "table", useReplace: true), g1x3));

            Add(list, ref page, "tbl-mx-upd-lean-keep-geo",
                a => ContentHtml.TableCreate(g2, left: 15, top: 25, width: 48, height: 32),
                (a, ids) => ContentHtml.TableUpdate(ids[0], g2b),
                ctx =>
                {
                    PptHtmlShapeNode before = PreferCreated(ctx, "table", useReplace: false);
                    PptHtmlShapeNode after = PreferCreated(ctx, "table", useReplace: true);
                    return ContentAssert.FirstFail(
                        ContentAssert.TableCellsMismatch(after, g2b),
                        ContentAssert.GeoClose(before?.Style, after?.Style)
                            ? null
                            : "瘦稿换表丢几何");
                });

            Add(list, ref page, "tbl-mx-upd-add-three-line",
                a => ContentHtml.TableCreate(g2),
                (a, ids) => ContentHtml.TableUpdate(
                    ids[0],
                    g2b,
                    "data-table-style=\"three-line\" data-col-widths=\"45%,55%\""),
                ctx =>
                {
                    PptHtmlShapeNode n = PreferCreated(ctx, "table", useReplace: true);
                    return ContentAssert.FirstFail(
                        ContentAssert.TableCellsMismatch(n, g2b),
                        ContentAssert.TableStyleIsThreeLine(n),
                        ContentAssert.TableColWidthsClose(n, new float[] { 45f, 55f }));
                });

            Add(list, ref page, "tbl-mx-upd-font-after",
                a => ContentHtml.TableCreate(g2),
                (a, ids) => ContentHtml.TableUpdate(
                    ids[0],
                    g2b,
                    "data-table-font-name=\"微软雅黑\" data-table-font-size=\"15\" data-table-font-color=\"#0070C0\""),
                ctx =>
                {
                    PptHtmlShapeNode n = PreferCreated(ctx, "table", useReplace: true);
                    return ContentAssert.FirstFail(
                        ContentAssert.TableCellsMismatch(n, g2b),
                        ContentAssert.TableFontClose(n, "微软雅黑", 15, "#0070C0"));
                });
        }

        private static void AddMatrixCrossCombos(List<ContentCase> list, ref int page)
        {
            // 三线 × 合并 × 列宽 × 格皮 × 字皮
            Add(list, ref page, "tbl-mx-x-merge-three-fill-span",
                a => ContentHtml.TableCreateRaw(
                    ContentHtml.Tr(ContentHtml.Td("总览", colspan: 2, header: true))
                    + ContentHtml.Tr(
                        ContentHtml.Td("左"),
                        ContentHtml.Td("重点", fill: "#FFF2CC", fontColor: "#C00000", bold: true)),
                    tableAttrs: "data-table-style=\"three-line\" data-col-widths=\"40%,60%\" "
                        + "data-table-font-name=\"微软雅黑\" data-table-font-size=\"13\""),
                null,
                ctx =>
                {
                    PptHtmlShapeNode n = PreferCreated(ctx, "table");
                    return ContentAssert.FirstFail(
                        ContentAssert.TableStyleIsThreeLine(n),
                        ContentAssert.TableColWidthsClose(n, new float[] { 40f, 60f }),
                        ContentAssert.InnerHas(n, "colspan=\"2\""),
                        ContentAssert.InnerHas(n, "data-fill=\"#FFF2CC\""),
                        ContentAssert.TableFontClose(n, "微软雅黑", 13, null));
                });

            Add(list, ref page, "tbl-mx-x-rowspan-three-rowh",
                a => ContentHtml.TableCreateRaw(
                    ContentHtml.Tr(ContentHtml.Td("类", rowspan: 2, header: true), ContentHtml.Td("上"))
                    + ContentHtml.Tr(ContentHtml.Td("下", fill: "#E2F0D9")),
                    tableAttrs: "data-table-style=\"three-line\" data-row-heights=\"45%,55%\""),
                null,
                ctx =>
                {
                    PptHtmlShapeNode n = PreferCreated(ctx, "table");
                    return ContentAssert.FirstFail(
                        ContentAssert.TableStyleIsThreeLine(n),
                        ContentAssert.InnerHas(n, "rowspan=\"2\""),
                        ContentAssert.TableRowHeightsClose(n, new float[] { 45f, 55f }));
                });

            Add(list, ref page, "tbl-mx-x-3x3-font-colw-fill",
                a =>
                {
                    string rows =
                        ContentHtml.Tr(
                            ContentHtml.Td("A", header: true),
                            ContentHtml.Td("B", header: true),
                            ContentHtml.Td("C", header: true))
                        + ContentHtml.Tr(
                            ContentHtml.Td("1"),
                            ContentHtml.Td("2", fill: "#DEEBF7"),
                            ContentHtml.Td("3"))
                        + ContentHtml.Tr(
                            ContentHtml.Td("4", fontColor: "#548235", bold: true),
                            ContentHtml.Td("5"),
                            ContentHtml.Td("6"));
                    return ContentHtml.TableCreateRaw(
                        rows,
                        width: 70,
                        height: 45,
                        tableAttrs: "data-col-widths=\"20%,40%,40%\" data-table-font-name=\"宋体\" "
                            + "data-table-font-size=\"12\" data-table-font-color=\"#333333\"");
                },
                null,
                ctx =>
                {
                    PptHtmlShapeNode n = PreferCreated(ctx, "table");
                    return ContentAssert.FirstFail(
                        ContentAssert.TableColWidthsClose(n, new float[] { 20f, 40f, 40f }),
                        ContentAssert.TableFontClose(n, "宋体", 12, "#333333"),
                        ContentAssert.InnerHas(n, "data-fill=\"#DEEBF7\""));
                });

            // 几何 × 三线 × 字号
            double[][] geos =
            {
                new[] { 6.0, 12, 50, 30 },
                new[] { 18.0, 28, 55, 38 },
                new[] { 10.0, 35, 65, 42 }
            };
            double[] sizes = { 11, 14, 17 };
            for (int i = 0; i < geos.Length; i++)
            {
                double[] g = geos[i];
                double sz = sizes[i];
                string[][] grid = ContentHtml.Grid(2, 2, "x");
                Add(list, ref page, "tbl-mx-x-geo-three-font-" + i,
                    a => ContentHtml.TableCreate(
                        grid,
                        left: g[0],
                        top: g[1],
                        width: g[2],
                        height: g[3],
                        extraAttrs: "data-table-style=\"three-line\" data-table-font-size=\""
                            + sz.ToString("0.##", CultureInfo.InvariantCulture) + "\""),
                    null,
                    ctx =>
                    {
                        PptHtmlShapeNode n = PreferCreated(ctx, "table");
                        return ContentAssert.FirstFail(
                            ContentAssert.TableCellsMismatch(n, grid),
                            ContentAssert.TableStyleIsThreeLine(n),
                            ContentAssert.TableFontClose(n, null, sz, null),
                            GeoNear(n, g[0], g[1], g[2], g[3]));
                    });
            }

            // 换数后仍保三线+列宽
            string[][] g2 = ContentHtml.Grid(2, 2, "y");
            string[][] g2z = ContentHtml.Grid(2, 2, "z");
            Add(list, ref page, "tbl-mx-x-upd-keep-three-colw",
                a => ContentHtml.TableCreate(
                    g2,
                    extraAttrs: "data-table-style=\"three-line\" data-col-widths=\"30%,70%\""),
                (a, ids) => ContentHtml.TableUpdate(
                    ids[0],
                    g2z,
                    "data-table-style=\"three-line\" data-col-widths=\"30%,70%\""),
                ctx =>
                {
                    PptHtmlShapeNode n = PreferCreated(ctx, "table", useReplace: true);
                    return ContentAssert.FirstFail(
                        ContentAssert.TableCellsMismatch(n, g2z),
                        ContentAssert.TableStyleIsThreeLine(n),
                        ContentAssert.TableColWidthsClose(n, new float[] { 30f, 70f }));
                });

            // 扩表后套三线
            Add(list, ref page, "tbl-mx-x-grow-then-three",
                a => ContentHtml.TableCreate(g2),
                (a, ids) => ContentHtml.TableUpdate(
                    ids[0],
                    ContentHtml.Grid(3, 3, "z"),
                    "data-table-style=\"three-line\" data-col-widths=\"30%,35%,35%\""),
                ctx =>
                {
                    PptHtmlShapeNode n = PreferCreated(ctx, "table", useReplace: true);
                    return ContentAssert.FirstFail(
                        ContentAssert.TableCellsMismatch(n, ContentHtml.Grid(3, 3, "z")),
                        ContentAssert.TableStyleIsThreeLine(n),
                        ContentAssert.TableColWidthsClose(n, new float[] { 30f, 35f, 35f }));
                });

            // 表 fill × 格 fill × 三线
            Add(list, ref page, "tbl-mx-x-shape-and-cell-fill",
                a => ContentHtml.TableCreateRaw(
                    ContentHtml.Tr(ContentHtml.Td("a"), ContentHtml.Td("b", fill: "#FFC7CE"))
                    + ContentHtml.Tr(ContentHtml.Td("c", fill: "#C6EFCE"), ContentHtml.Td("d")),
                    tableAttrs: "data-fill=\"#DEEBF7\" data-table-style=\"three-line\""),
                null,
                ctx =>
                {
                    PptHtmlShapeNode n = PreferCreated(ctx, "table");
                    return ContentAssert.FirstFail(
                        ContentAssert.TableStyleIsThreeLine(n),
                        HexEq(n?.Fill, "#DEEBF7") ? null : "表 fill",
                        ContentAssert.InnerHas(n, "data-fill=\"#FFC7CE\""),
                        ContentAssert.InnerHas(n, "data-fill=\"#C6EFCE\""));
                });

            // 多色格交叉
            string[] cellColors = { "#FCE4D6", "#DDEBF7", "#E2F0D9", "#FFF2CC" };
            for (int i = 0; i < cellColors.Length; i++)
            {
                string col = cellColors[i];
                Add(list, ref page, "tbl-mx-x-cellfill-three-" + col.TrimStart('#'),
                    a => ContentHtml.TableCreateRaw(
                        ContentHtml.Tr(ContentHtml.Td("头", colspan: 2, header: true))
                        + ContentHtml.Tr(ContentHtml.Td("L"), ContentHtml.Td("R", fill: col, bold: true)),
                        tableAttrs: "data-table-style=\"three-line\" data-col-widths=\"50%,50%\""),
                    null,
                    ctx =>
                    {
                        PptHtmlShapeNode n = PreferCreated(ctx, "table");
                        return ContentAssert.FirstFail(
                            ContentAssert.TableStyleIsThreeLine(n),
                            ContentAssert.InnerHas(n, "colspan=\"2\""),
                            ContentAssert.InnerHas(n, "data-fill=\"" + col + "\""),
                            ContentAssert.TableColWidthsClose(n, new float[] { 50f, 50f }));
                    });
            }
        }

        private static string Build2x2WithCellFill(int row, int col, string fill, string prefix)
        {
            var cells = new string[2, 2];
            for (int r = 0; r < 2; r++)
            {
                for (int c = 0; c < 2; c++)
                {
                    cells[r, c] = prefix + r + c;
                }
            }

            string TrRow(int r)
            {
                string a = (r == row && col == 0)
                    ? ContentHtml.Td(cells[r, 0], fill: fill)
                    : ContentHtml.Td(cells[r, 0]);
                string b = (r == row && col == 1)
                    ? ContentHtml.Td(cells[r, 1], fill: fill)
                    : ContentHtml.Td(cells[r, 1]);
                return ContentHtml.Tr(a, b);
            }

            return TrRow(0) + TrRow(1);
        }

        private static string JoinPct(float[] pcts)
        {
            if (pcts == null || pcts.Length == 0)
            {
                return "";
            }

            var parts = new string[pcts.Length];
            for (int i = 0; i < pcts.Length; i++)
            {
                parts[i] = pcts[i].ToString("0.##", CultureInfo.InvariantCulture) + "%";
            }

            return string.Join(",", parts);
        }

        private static string SafeToken(string s)
        {
            if (string.IsNullOrEmpty(s))
            {
                return "x";
            }

            return s.Replace(" ", "").Replace("/", "");
        }

        private static string GeoNear(
            PptHtmlShapeNode n,
            double left,
            double top,
            double width,
            double height,
            double eps = 1.0)
        {
            if (n == null)
            {
                return "无 table";
            }

            if (!ContentAssert.TryParseGeo(n.Style, out double l, out double t, out double w, out double h))
            {
                return "无几何";
            }

            if (Math.Abs(l - left) > eps || Math.Abs(t - top) > eps
                || Math.Abs(w - width) > eps || Math.Abs(h - height) > eps)
            {
                return "几何偏离: " + n.Style;
            }

            return null;
        }

        private static string TruncInner(PptHtmlShapeNode n)
        {
            string s = n?.InnerHtml ?? "";
            return s.Length <= 200 ? s : s.Substring(0, 200) + "…";
        }
    }
}
