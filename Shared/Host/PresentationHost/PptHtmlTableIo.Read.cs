using System;
using System.Collections.Generic;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Text;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace WordAddIn1.PresentationHost
{
    /// <summary>PPT Interop：TryRead / 合并拓扑读出。</summary>
    internal static partial class PptHtmlTableIo
    {
        public static bool TryRead(
            PowerPoint.Shape shape,
            out PptHtmlTableGrid grid,
            out PptHtmlTableStyleSnap style,
            out string error,
            List<string> warnings)
        {
            grid = null;
            style = null;
            error = null;
            if (warnings == null)
            {
                warnings = new List<string>();
            }

            try
            {
                if (shape == null || shape.HasTable != Office.MsoTriState.msoTrue)
                {
                    error = "目标不是表格";
                    return false;
                }

                PowerPoint.Table table = shape.Table;
                int rows = table.Rows.Count;
                int cols = table.Columns.Count;
                bool truncated = false;
                int useRows = rows;
                int useCols = cols;
                if (useRows > PptHtmlTableGrid.MaxRows || useCols > PptHtmlTableGrid.MaxCols)
                {
                    truncated = true;
                    useRows = Math.Min(useRows, PptHtmlTableGrid.MaxRows);
                    useCols = Math.Min(useCols, PptHtmlTableGrid.MaxCols);
                    warnings.Add("表格超过 " + PptHtmlTableGrid.MaxRows + "×" + PptHtmlTableGrid.MaxCols
                        + "，读出已截断");
                }

                var mergeMap = ResolveMergeMap(shape, table, useRows, useCols, warnings);
                var cells = new List<PptHtmlTableCell>(useRows * useCols);
                for (int r = 1; r <= useRows; r++)
                {
                    for (int c = 1; c <= useCols; c++)
                    {
                        int key = Pack(r, c);
                        if (mergeMap.Covered.Contains(key))
                        {
                            cells.Add(new PptHtmlTableCell { IsCovered = true, Text = "" });
                            continue;
                        }

                        int rs = 1;
                        int cs = 1;
                        if (mergeMap.Anchors.TryGetValue(key, out MergeSpan span))
                        {
                            rs = span.RowSpan;
                            cs = span.ColSpan;
                        }

                        PowerPoint.Cell cell = table.Cell(r, c);
                        var outCell = new PptHtmlTableCell
                        {
                            IsCovered = false,
                            RowSpan = rs,
                            ColSpan = cs,
                            Text = ReadCellText(cell)
                        };
                        TryReadCellSkin(cell, outCell, warnings);
                        cells.Add(outCell);
                    }
                }

                grid = new PptHtmlTableGrid
                {
                    RowCount = useRows,
                    ColCount = useCols,
                    Cells = cells
                };

                style = new PptHtmlTableStyleSnap { Truncated = truncated };
                TryReadColRowPcts(shape, table, useRows, useCols, style, warnings);
                TryInferTableFont(table, useRows, useCols, style, warnings);
                return true;
            }
            catch (Exception ex)
            {
                error = "读表格失败: " + ex.Message;
                return false;
            }
        }

        private static void TryReadCellSkin(
            PowerPoint.Cell pptCell,
            PptHtmlTableCell cell,
            List<string> warnings)
        {
            try
            {
                PowerPoint.TextRange tr = pptCell.Shape.TextFrame.TextRange;
                try
                {
                    if (tr.Font.Bold == Office.MsoTriState.msoTrue)
                    {
                        cell.FontBold = true;
                    }
                }
                catch
                {
                }

                try
                {
                    if (tr.Font.Italic == Office.MsoTriState.msoTrue)
                    {
                        cell.FontItalic = true;
                    }
                }
                catch
                {
                }

                try
                {
                    int rgb = tr.Font.Color.RGB;
                    cell.FontColor = FormatRgb(rgb);
                }
                catch
                {
                }

                try
                {
                    if (pptCell.Shape.Fill.Visible == Office.MsoTriState.msoTrue)
                    {
                        cell.Fill = FormatRgb(pptCell.Shape.Fill.ForeColor.RGB);
                    }
                }
                catch
                {
                }
            }
            catch (Exception ex)
            {
                warnings.Add("读单元格样式失败: " + ex.Message);
            }
        }

        private static void TryReadColRowPcts(
            PowerPoint.Shape shape,
            PowerPoint.Table table,
            int rows,
            int cols,
            PptHtmlTableStyleSnap style,
            List<string> warnings)
        {
            try
            {
                float tableW = shape.Width;
                float tableH = shape.Height;
                if (tableW > 0 && cols > 0)
                {
                    var pcts = new float[cols];
                    for (int c = 1; c <= cols; c++)
                    {
                        pcts[c - 1] = table.Columns[c].Width / tableW * 100f;
                    }

                    style.ColWidthPcts = pcts;
                }

                if (tableH > 0 && rows > 0)
                {
                    var pcts = new float[rows];
                    for (int r = 1; r <= rows; r++)
                    {
                        pcts[r - 1] = table.Rows[r].Height / tableH * 100f;
                    }

                    style.RowHeightPcts = pcts;
                }
            }
            catch (Exception ex)
            {
                warnings.Add("读列宽/行高失败: " + ex.Message);
            }
        }

        private static void TryInferTableFont(
            PowerPoint.Table table,
            int rows,
            int cols,
            PptHtmlTableStyleSnap style,
            List<string> warnings)
        {
            try
            {
                if (rows < 1 || cols < 1)
                {
                    return;
                }

                PowerPoint.TextRange tr = table.Cell(1, 1).Shape.TextFrame.TextRange;
                try
                {
                    string name = tr.Font.Name;
                    if (!string.IsNullOrWhiteSpace(name))
                    {
                        style.FontName = name;
                    }
                }
                catch
                {
                }

                try
                {
                    style.FontSizePt = tr.Font.Size;
                }
                catch
                {
                }

                try
                {
                    style.FontColor = FormatRgb(tr.Font.Color.RGB);
                }
                catch
                {
                }
            }
            catch (Exception ex)
            {
                warnings.Add("推断整表字体失败: " + ex.Message);
            }
        }

        private static string ReadCellText(PowerPoint.Cell cell)
        {
            try
            {
                return cell.Shape.TextFrame.TextRange.Text ?? "";
            }
            catch
            {
                return "";
            }
        }

        private static MergeMap BuildMergeMap(
            PowerPoint.Table table,
            int rows,
            int cols,
            List<string> warnings)
        {
            var map = new MergeMap();
            var shapeKey = new Dictionary<int, string>();
            for (int r = 1; r <= rows; r++)
            {
                for (int c = 1; c <= cols; c++)
                {
                    string id;
                    try
                    {
                        PowerPoint.Shape sh = table.Cell(r, c).Shape;
                        id = sh.Id.ToString(CultureInfo.InvariantCulture)
                            + "@"
                            + sh.Left.ToString("0.###", CultureInfo.InvariantCulture)
                            + ","
                            + sh.Top.ToString("0.###", CultureInfo.InvariantCulture)
                            + ","
                            + sh.Width.ToString("0.###", CultureInfo.InvariantCulture)
                            + ","
                            + sh.Height.ToString("0.###", CultureInfo.InvariantCulture);
                    }
                    catch
                    {
                        id = "r" + r + "c" + c;
                    }

                    shapeKey[Pack(r, c)] = id;
                }
            }

            var visited = new HashSet<int>();
            for (int r = 1; r <= rows; r++)
            {
                for (int c = 1; c <= cols; c++)
                {
                    int key = Pack(r, c);
                    if (visited.Contains(key))
                    {
                        continue;
                    }

                    string id = shapeKey[key];
                    int maxR = r;
                    int maxC = c;
                    for (int rr = r; rr <= rows; rr++)
                    {
                        for (int cc = c; cc <= cols; cc++)
                        {
                            int k2 = Pack(rr, cc);
                            if (shapeKey[k2] == id)
                            {
                                if (rr > maxR)
                                {
                                    maxR = rr;
                                }

                                if (cc > maxC)
                                {
                                    maxC = cc;
                                }
                            }
                        }
                    }

                    int rs = maxR - r + 1;
                    int cs = maxC - c + 1;
                    bool rectangular = true;
                    for (int rr = r; rr <= maxR && rectangular; rr++)
                    {
                        for (int cc = c; cc <= maxC; cc++)
                        {
                            if (shapeKey[Pack(rr, cc)] != id)
                            {
                                rectangular = false;
                                break;
                            }
                        }
                    }

                    if (!rectangular)
                    {
                        visited.Add(key);
                        map.Anchors[key] = new MergeSpan { RowSpan = 1, ColSpan = 1 };
                        continue;
                    }

                    map.Anchors[key] = new MergeSpan { RowSpan = rs, ColSpan = cs };
                    for (int rr = r; rr <= maxR; rr++)
                    {
                        for (int cc = c; cc <= maxC; cc++)
                        {
                            int k2 = Pack(rr, cc);
                            visited.Add(k2);
                            if (rr != r || cc != c)
                            {
                                map.Covered.Add(k2);
                            }
                        }
                    }
                }
            }

            return map;
        }
    }
}
