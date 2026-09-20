using System;
using System.Collections.Generic;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Text;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace WordAddIn1.PresentationHost
{
    /// <summary>PPT Interop：TryApply / 扩缩 / Merge / 皮。</summary>
    internal static partial class PptHtmlTableIo
    {
        public static bool TryApply(
            PowerPoint.Shape shape,
            PptHtmlTableGrid grid,
            PptHtmlTableStyleSnap style,
            out string error,
            List<string> warnings)
        {
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

                if (grid == null)
                {
                    if (style == null)
                    {
                        error = "缺少表格网格";
                        return false;
                    }

                    PowerPoint.Table tableOnly = shape.Table;
                    TryApplyTableStyle(tableOnly, style, warnings);
                    if (style.ColWidthPcts != null || style.RowHeightPcts != null)
                    {
                        TryApplyColRowPcts(shape, tableOnly, style.ColWidthPcts, style.RowHeightPcts, warnings);
                    }

                    return true;
                }

                if (grid.RowCount < 1 || grid.ColCount < 1)
                {
                    error = "表格网格无效";
                    return false;
                }

                if (grid.RowCount > PptHtmlTableGrid.MaxRows || grid.ColCount > PptHtmlTableGrid.MaxCols)
                {
                    error = "页内表格最多 " + PptHtmlTableGrid.MaxRows + " 行 × "
                        + PptHtmlTableGrid.MaxCols + " 列（含合并占位）";
                    return false;
                }

                PowerPoint.Table table = shape.Table;
                if (!TryAlignDimensions(table, grid.RowCount, grid.ColCount, out error))
                {
                    return false;
                }

                if (!TryClearMerges(table, out error, warnings))
                {
                    return false;
                }

                if (!TryWriteCells(table, grid, out error, warnings))
                {
                    return false;
                }

                if (!TryApplyMerges(table, grid, out error))
                {
                    return false;
                }

                if (style != null)
                {
                    if (style.ColWidthPcts != null || style.RowHeightPcts != null)
                    {
                        TryApplyColRowPcts(shape, table, style.ColWidthPcts, style.RowHeightPcts, warnings);
                    }

                    TryApplyTableStyle(table, style, warnings);
                }

                return true;
            }
            catch (Exception ex)
            {
                error = "写表格失败: " + ex.Message;
                return false;
            }
        }

        public static bool TryAlignDimensions(
            PowerPoint.Table table,
            int wantRows,
            int wantCols,
            out string error)
        {
            error = null;
            try
            {
                while (table.Rows.Count < wantRows)
                {
                    table.Rows.Add(-1);
                }

                while (table.Rows.Count > wantRows)
                {
                    table.Rows[table.Rows.Count].Delete();
                }

                while (table.Columns.Count < wantCols)
                {
                    table.Columns.Add(-1);
                }

                while (table.Columns.Count > wantCols)
                {
                    table.Columns[table.Columns.Count].Delete();
                }

                return true;
            }
            catch (Exception ex)
            {
                error = "调整表格行列失败: " + ex.Message;
                return false;
            }
        }

        private static bool TryClearMerges(
            PowerPoint.Table table,
            out string error,
            List<string> warnings)
        {
            error = null;
            try
            {
                int rows = table.Rows.Count;
                int cols = table.Columns.Count;
                MergeMap map = BuildMergeMap(table, rows, cols, warnings);
                foreach (KeyValuePair<int, MergeSpan> kv in map.Anchors)
                {
                    Unpack(kv.Key, out int r, out int c);
                    if (kv.Value.RowSpan <= 1 && kv.Value.ColSpan <= 1)
                    {
                        continue;
                    }

                    try
                    {
                        table.Cell(r, c).Split(kv.Value.RowSpan, kv.Value.ColSpan);
                    }
                    catch (Exception ex)
                    {
                        warnings.Add("拆分合并格 (" + r + "," + c + ") 失败: " + ex.Message);
                    }
                }

                return true;
            }
            catch (Exception ex)
            {
                error = "清除合并失败: " + ex.Message;
                return false;
            }
        }

        private static bool TryWriteCells(
            PowerPoint.Table table,
            PptHtmlTableGrid grid,
            out string error,
            List<string> warnings)
        {
            error = null;
            try
            {
                for (int r = 0; r < grid.RowCount; r++)
                {
                    for (int c = 0; c < grid.ColCount; c++)
                    {
                        PptHtmlTableCell cell = grid.CellAt(r, c);
                        if (cell == null || cell.IsCovered)
                        {
                            continue;
                        }

                        PowerPoint.Cell pptCell = table.Cell(r + 1, c + 1);
                        try
                        {
                            pptCell.Shape.TextFrame.TextRange.Text = cell.Text ?? "";
                        }
                        catch (Exception ex)
                        {
                            error = "写单元格文字失败 (" + (r + 1) + "," + (c + 1) + "): " + ex.Message;
                            return false;
                        }

                        TryWriteCellSkin(pptCell, cell, warnings);
                    }
                }

                return true;
            }
            catch (Exception ex)
            {
                error = "写单元格失败: " + ex.Message;
                return false;
            }
        }

        private static bool TryApplyMerges(
            PowerPoint.Table table,
            PptHtmlTableGrid grid,
            out string error)
        {
            error = null;
            try
            {
                for (int r = 0; r < grid.RowCount; r++)
                {
                    for (int c = 0; c < grid.ColCount; c++)
                    {
                        PptHtmlTableCell cell = grid.CellAt(r, c);
                        if (cell == null || cell.IsCovered)
                        {
                            continue;
                        }

                        if (cell.RowSpan <= 1 && cell.ColSpan <= 1)
                        {
                            continue;
                        }

                        PowerPoint.Cell start = table.Cell(r + 1, c + 1);
                        PowerPoint.Cell end = table.Cell(r + cell.RowSpan, c + cell.ColSpan);
                        start.Merge(end);
                    }
                }

                return true;
            }
            catch (Exception ex)
            {
                error = "合并单元格失败: " + ex.Message;
                return false;
            }
        }

        private static void TryWriteCellSkin(
            PowerPoint.Cell pptCell,
            PptHtmlTableCell cell,
            List<string> warnings)
        {
            try
            {
                PowerPoint.TextRange tr = pptCell.Shape.TextFrame.TextRange;
                if (!string.IsNullOrEmpty(cell.FontColor) && TryParseRgb(cell.FontColor, out int rgb))
                {
                    tr.Font.Color.RGB = rgb;
                }

                if (cell.FontBold.HasValue)
                {
                    tr.Font.Bold = cell.FontBold.Value
                        ? Office.MsoTriState.msoTrue
                        : Office.MsoTriState.msoFalse;
                }

                if (cell.FontItalic.HasValue)
                {
                    tr.Font.Italic = cell.FontItalic.Value
                        ? Office.MsoTriState.msoTrue
                        : Office.MsoTriState.msoFalse;
                }

                if (!string.IsNullOrEmpty(cell.Fill))
                {
                    if (string.Equals(cell.Fill, "none", StringComparison.OrdinalIgnoreCase))
                    {
                        pptCell.Shape.Fill.Visible = Office.MsoTriState.msoFalse;
                    }
                    else if (TryParseRgb(cell.Fill, out int fillRgb))
                    {
                        pptCell.Shape.Fill.Visible = Office.MsoTriState.msoTrue;
                        pptCell.Shape.Fill.Solid();
                        pptCell.Shape.Fill.ForeColor.RGB = fillRgb;
                    }
                }
            }
            catch (Exception ex)
            {
                warnings.Add("写单元格样式失败: " + ex.Message);
            }
        }

        /// <summary>
        /// 表级 data-fill 写形状 Fill 会刷掉格色；在表 fill 之后把网格里显式格 fill 再盖回去。
        /// </summary>
        public static void TryReapplyExplicitCellFills(
            PowerPoint.Shape shape,
            PptHtmlTableGrid grid,
            List<string> warnings)
        {
            if (shape == null || grid == null || shape.HasTable != Office.MsoTriState.msoTrue)
            {
                return;
            }

            if (warnings == null)
            {
                warnings = new List<string>();
            }

            try
            {
                PowerPoint.Table table = shape.Table;
                for (int r = 0; r < grid.RowCount; r++)
                {
                    for (int c = 0; c < grid.ColCount; c++)
                    {
                        PptHtmlTableCell cell = grid.CellAt(r, c);
                        if (cell == null || cell.IsCovered || string.IsNullOrEmpty(cell.Fill))
                        {
                            continue;
                        }

                        try
                        {
                            PowerPoint.Cell pptCell = table.Cell(r + 1, c + 1);
                            if (string.Equals(cell.Fill, "none", StringComparison.OrdinalIgnoreCase))
                            {
                                pptCell.Shape.Fill.Visible = Office.MsoTriState.msoFalse;
                            }
                            else if (TryParseRgb(cell.Fill, out int fillRgb))
                            {
                                pptCell.Shape.Fill.Visible = Office.MsoTriState.msoTrue;
                                pptCell.Shape.Fill.Solid();
                                pptCell.Shape.Fill.ForeColor.RGB = fillRgb;
                            }
                        }
                        catch (Exception ex)
                        {
                            warnings.Add("回写格 fill 失败 (" + (r + 1) + "," + (c + 1) + "): " + ex.Message);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                warnings.Add("回写格 fill 失败: " + ex.Message);
            }
        }

        private static void TryApplyColRowPcts(
            PowerPoint.Shape shape,
            PowerPoint.Table table,
            float[] colPcts,
            float[] rowPcts,
            List<string> warnings)
        {
            try
            {
                float tableW = shape.Width;
                float tableH = shape.Height;
                if (colPcts != null && colPcts.Length == table.Columns.Count && tableW > 0)
                {
                    float[] norm = NormalizePcts(colPcts);
                    for (int i = 0; i < norm.Length; i++)
                    {
                        table.Columns[i + 1].Width = tableW * norm[i] / 100f;
                    }
                }

                if (rowPcts != null && rowPcts.Length == table.Rows.Count && tableH > 0)
                {
                    float[] norm = NormalizePcts(rowPcts);
                    for (int i = 0; i < norm.Length; i++)
                    {
                        table.Rows[i + 1].Height = tableH * norm[i] / 100f;
                    }
                }
            }
            catch (Exception ex)
            {
                warnings.Add("套列宽/行高失败: " + ex.Message);
            }
        }

        private static void TryApplyTableStyle(
            PowerPoint.Table table,
            PptHtmlTableStyleSnap style,
            List<string> warnings)
        {
            try
            {
                int rows = table.Rows.Count;
                int cols = table.Columns.Count;
                if (!string.IsNullOrEmpty(style.FontName)
                    || style.FontSizePt.HasValue
                    || !string.IsNullOrEmpty(style.FontColor))
                {
                    for (int r = 1; r <= rows; r++)
                    {
                        for (int c = 1; c <= cols; c++)
                        {
                            try
                            {
                                PowerPoint.TextRange tr = table.Cell(r, c).Shape.TextFrame.TextRange;
                                if (!string.IsNullOrEmpty(style.FontName))
                                {
                                    tr.Font.Name = style.FontName;
                                }

                                if (style.FontSizePt.HasValue)
                                {
                                    tr.Font.Size = (float)style.FontSizePt.Value;
                                }

                                if (!string.IsNullOrEmpty(style.FontColor)
                                    && TryParseRgb(style.FontColor, out int rgb))
                                {
                                    tr.Font.Color.RGB = rgb;
                                }
                            }
                            catch
                            {
                            }
                        }
                    }
                }

                if (string.Equals(style.TableStyle, "three-line", StringComparison.OrdinalIgnoreCase))
                {
                    ApplyThreeLine(table, warnings);
                }
            }
            catch (Exception ex)
            {
                warnings.Add("套整表样式失败: " + ex.Message);
            }
        }

        private static void ApplyThreeLine(PowerPoint.Table table, List<string> warnings)
        {
            try
            {
                int rows = table.Rows.Count;
                int cols = table.Columns.Count;
                for (int r = 1; r <= rows; r++)
                {
                    for (int c = 1; c <= cols; c++)
                    {
                        PowerPoint.Cell cell = table.Cell(r, c);
                        SetBorderVisible(cell, "left", false);
                        SetBorderVisible(cell, "right", false);
                        SetBorderVisible(cell, "top", r == 1);
                        SetBorderVisible(cell, "bottom", r == 1 || r == rows);
                    }
                }
            }
            catch (Exception ex)
            {
                warnings.Add("套三线表失败: " + ex.Message);
            }
        }

        private static void SetBorderVisible(PowerPoint.Cell cell, string which, bool visible)
        {
            try
            {
                // Borders index: ppBorderTop=1 Left=2 Bottom=3 Right=4 DiagonalDown=5 DiagonalUp=6
                int idx;
                switch (which)
                {
                    case "top":
                        idx = 1;
                        break;
                    case "left":
                        idx = 2;
                        break;
                    case "bottom":
                        idx = 3;
                        break;
                    case "right":
                        idx = 4;
                        break;
                    default:
                        return;
                }

                PowerPoint.LineFormat line = cell.Borders[(PowerPoint.PpBorderType)idx];
                line.Visible = visible ? Office.MsoTriState.msoTrue : Office.MsoTriState.msoFalse;
                if (visible)
                {
                    line.Weight = 1f;
                    line.ForeColor.RGB = 0x000000;
                }
            }
            catch
            {
            }
        }
    }
}
