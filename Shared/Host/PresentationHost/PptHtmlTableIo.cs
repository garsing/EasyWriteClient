using System;
using System.Collections.Generic;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Text;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace WordAddIn1.PresentationHost
{
    /// <summary>页内 table 形状 COM 读写（PPT Interop；WPP 晚绑定另入口）。</summary>
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

                var mergeMap = BuildMergeMap(table, useRows, useCols, warnings);
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

        public static string BuildInnerHtml(PptHtmlTableGrid grid)
        {
            var sb = new StringBuilder();
            PptHtmlTableParse.AppendInnerHtml(sb, grid, "    ");
            return sb.ToString();
        }

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

        private static float[] NormalizePcts(float[] pcts)
        {
            float sum = 0;
            for (int i = 0; i < pcts.Length; i++)
            {
                sum += pcts[i];
            }

            if (sum <= 0)
            {
                float even = 100f / pcts.Length;
                var evenArr = new float[pcts.Length];
                for (int i = 0; i < evenArr.Length; i++)
                {
                    evenArr[i] = even;
                }

                return evenArr;
            }

            if (Math.Abs(sum - 100f) <= 0.5f)
            {
                return pcts;
            }

            var norm = new float[pcts.Length];
            for (int i = 0; i < pcts.Length; i++)
            {
                norm[i] = pcts[i] * 100f / sum;
            }

            return norm;
        }

        private static bool TryParseRgb(string color, out int rgb)
        {
            rgb = 0;
            if (string.IsNullOrWhiteSpace(color) || color.Length != 7 || color[0] != '#')
            {
                return false;
            }

            if (!int.TryParse(color.Substring(1, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out int r)
                || !int.TryParse(color.Substring(3, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out int g)
                || !int.TryParse(color.Substring(5, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out int b))
            {
                return false;
            }

            rgb = r + (g << 8) + (b << 16);
            return true;
        }

        private static string FormatRgb(int rgb)
        {
            int r = rgb & 0xFF;
            int g = (rgb >> 8) & 0xFF;
            int b = (rgb >> 16) & 0xFF;
            return "#" + r.ToString("X2") + g.ToString("X2") + b.ToString("X2");
        }

        private static int Pack(int r, int c)
        {
            return (r << 16) | c;
        }

        private static void Unpack(int key, out int r, out int c)
        {
            r = key >> 16;
            c = key & 0xFFFF;
        }

        private sealed class MergeSpan
        {
            public int RowSpan;
            public int ColSpan;
        }

        private sealed class MergeMap
        {
            public Dictionary<int, MergeSpan> Anchors = new Dictionary<int, MergeSpan>();
            public HashSet<int> Covered = new HashSet<int>();
        }
    }
}
