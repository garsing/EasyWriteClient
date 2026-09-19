using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using WordAddIn1.OpenFiles;

namespace WordAddIn1.PresentationHost
{
    /// <summary>WPP 晚绑定：与 PPT Interop 同语义的 TryRead / TryApply。</summary>
    internal static partial class PptHtmlTableIo
    {
        public static bool TryRead(
            object shape,
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
                if (shape == null || !IsTruthy(WppCom.GetProperty(shape, "HasTable")))
                {
                    error = "目标不是表格";
                    return false;
                }

                object table = WppCom.GetProperty(shape, "Table");
                object rowsObj = WppCom.GetProperty(table, "Rows");
                object colsObj = WppCom.GetProperty(table, "Columns");
                int rows = Convert.ToInt32(WppCom.GetProperty(rowsObj, "Count"));
                int cols = Convert.ToInt32(WppCom.GetProperty(colsObj, "Count"));
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

                MergeMap mergeMap = ResolveMergeMapWpp(shape, table, useRows, useCols, warnings);
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

                        object cell = WppCom.Invoke(table, "Cell", r, c);
                        var outCell = new PptHtmlTableCell
                        {
                            IsCovered = false,
                            RowSpan = rs,
                            ColSpan = cs,
                            Text = ReadCellTextWpp(cell)
                        };
                        TryReadCellSkinWpp(cell, outCell, warnings);
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
                TryReadColRowPctsWpp(shape, table, useRows, useCols, style, warnings);
                TryInferTableFontWpp(table, useRows, useCols, style, warnings);
                return true;
            }
            catch (Exception ex)
            {
                error = "读表格失败: " + ex.Message;
                return false;
            }
        }

        public static bool TryApply(
            object shape,
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
                if (shape == null || !IsTruthy(WppCom.GetProperty(shape, "HasTable")))
                {
                    error = "目标不是表格";
                    return false;
                }

                object table = WppCom.GetProperty(shape, "Table");
                if (grid == null)
                {
                    if (style == null)
                    {
                        error = "缺少表格网格";
                        return false;
                    }

                    TryApplyTableStyleWpp(table, style, warnings);
                    if (style.ColWidthPcts != null || style.RowHeightPcts != null)
                    {
                        TryApplyColRowPctsWpp(shape, table, style.ColWidthPcts, style.RowHeightPcts, warnings);
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

                if (!TryAlignDimensionsWpp(table, grid.RowCount, grid.ColCount, out error))
                {
                    return false;
                }

                if (!TryClearMergesWpp(table, out error, warnings))
                {
                    return false;
                }

                if (!TryWriteCellsWpp(table, grid, out error, warnings))
                {
                    return false;
                }

                if (!TryApplyMergesWpp(table, grid, out error))
                {
                    return false;
                }

                if (style != null)
                {
                    if (style.ColWidthPcts != null || style.RowHeightPcts != null)
                    {
                        TryApplyColRowPctsWpp(shape, table, style.ColWidthPcts, style.RowHeightPcts, warnings);
                    }

                    TryApplyTableStyleWpp(table, style, warnings);
                }

                return true;
            }
            catch (Exception ex)
            {
                error = "写表格失败: " + ex.Message;
                return false;
            }
        }

        private static bool TryAlignDimensionsWpp(
            object table,
            int wantRows,
            int wantCols,
            out string error)
        {
            error = null;
            try
            {
                object rowsObj = WppCom.GetProperty(table, "Rows");
                object colsObj = WppCom.GetProperty(table, "Columns");
                while (Convert.ToInt32(WppCom.GetProperty(rowsObj, "Count")) < wantRows)
                {
                    WppCom.Invoke(rowsObj, "Add", -1);
                }

                while (Convert.ToInt32(WppCom.GetProperty(rowsObj, "Count")) > wantRows)
                {
                    int n = Convert.ToInt32(WppCom.GetProperty(rowsObj, "Count"));
                    object last = WppCom.GetIndexed(rowsObj, n);
                    WppCom.Invoke(last, "Delete");
                }

                while (Convert.ToInt32(WppCom.GetProperty(colsObj, "Count")) < wantCols)
                {
                    WppCom.Invoke(colsObj, "Add", -1);
                }

                while (Convert.ToInt32(WppCom.GetProperty(colsObj, "Count")) > wantCols)
                {
                    int n = Convert.ToInt32(WppCom.GetProperty(colsObj, "Count"));
                    object last = WppCom.GetIndexed(colsObj, n);
                    WppCom.Invoke(last, "Delete");
                }

                return true;
            }
            catch (Exception ex)
            {
                error = "调整表格行列失败: " + ex.Message;
                return false;
            }
        }

        private static bool TryClearMergesWpp(
            object table,
            out string error,
            List<string> warnings)
        {
            error = null;
            try
            {
                object rowsObj = WppCom.GetProperty(table, "Rows");
                object colsObj = WppCom.GetProperty(table, "Columns");
                int rows = Convert.ToInt32(WppCom.GetProperty(rowsObj, "Count"));
                int cols = Convert.ToInt32(WppCom.GetProperty(colsObj, "Count"));
                MergeMap map = BuildMergeMapWpp(table, rows, cols, warnings);
                foreach (KeyValuePair<int, MergeSpan> kv in map.Anchors)
                {
                    Unpack(kv.Key, out int r, out int c);
                    if (kv.Value.RowSpan <= 1 && kv.Value.ColSpan <= 1)
                    {
                        continue;
                    }

                    try
                    {
                        object cell = WppCom.Invoke(table, "Cell", r, c);
                        WppCom.Invoke(cell, "Split", kv.Value.RowSpan, kv.Value.ColSpan);
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

        private static bool TryWriteCellsWpp(
            object table,
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

                        object pptCell = WppCom.Invoke(table, "Cell", r + 1, c + 1);
                        try
                        {
                            object cellShape = WppCom.GetProperty(pptCell, "Shape");
                            object tf = WppCom.GetProperty(cellShape, "TextFrame");
                            object tr = WppCom.GetProperty(tf, "TextRange");
                            WppCom.TrySetProperty(tr, "Text", cell.Text ?? "");
                        }
                        catch (Exception ex)
                        {
                            error = "写单元格文字失败 (" + (r + 1) + "," + (c + 1) + "): " + ex.Message;
                            return false;
                        }

                        TryWriteCellSkinWpp(pptCell, cell, warnings);
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

        private static bool TryApplyMergesWpp(
            object table,
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

                        object start = WppCom.Invoke(table, "Cell", r + 1, c + 1);
                        object end = WppCom.Invoke(table, "Cell", r + cell.RowSpan, c + cell.ColSpan);
                        WppCom.Invoke(start, "Merge", end);
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

        private static void TryWriteCellSkinWpp(
            object pptCell,
            PptHtmlTableCell cell,
            List<string> warnings)
        {
            try
            {
                object cellShape = WppCom.GetProperty(pptCell, "Shape");
                object tf = WppCom.GetProperty(cellShape, "TextFrame");
                object tr = WppCom.GetProperty(tf, "TextRange");
                object font = WppCom.GetProperty(tr, "Font");
                if (!string.IsNullOrEmpty(cell.FontColor) && TryParseRgb(cell.FontColor, out int rgb))
                {
                    object color = WppCom.GetProperty(font, "Color");
                    WppCom.TrySetProperty(color, "RGB", rgb);
                }

                if (cell.FontBold.HasValue)
                {
                    WppCom.TrySetProperty(font, "Bold", cell.FontBold.Value ? -1 : 0);
                }

                if (cell.FontItalic.HasValue)
                {
                    WppCom.TrySetProperty(font, "Italic", cell.FontItalic.Value ? -1 : 0);
                }

                if (!string.IsNullOrEmpty(cell.Fill))
                {
                    object fill = WppCom.GetProperty(cellShape, "Fill");
                    if (string.Equals(cell.Fill, "none", StringComparison.OrdinalIgnoreCase))
                    {
                        WppCom.TrySetProperty(fill, "Visible", 0);
                    }
                    else if (TryParseRgb(cell.Fill, out int fillRgb))
                    {
                        WppCom.TrySetProperty(fill, "Visible", -1);
                        try
                        {
                            WppCom.Invoke(fill, "Solid");
                        }
                        catch
                        {
                        }

                        object fore = WppCom.GetProperty(fill, "ForeColor");
                        WppCom.TrySetProperty(fore, "RGB", fillRgb);
                    }
                }
            }
            catch (Exception ex)
            {
                warnings.Add("写单元格样式失败: " + ex.Message);
            }
        }

        private static void TryReadCellSkinWpp(
            object pptCell,
            PptHtmlTableCell cell,
            List<string> warnings)
        {
            try
            {
                object cellShape = WppCom.GetProperty(pptCell, "Shape");
                object tf = WppCom.GetProperty(cellShape, "TextFrame");
                object tr = WppCom.GetProperty(tf, "TextRange");
                object font = WppCom.GetProperty(tr, "Font");
                try
                {
                    if (IsTruthy(WppCom.GetProperty(font, "Bold")))
                    {
                        cell.FontBold = true;
                    }
                }
                catch
                {
                }

                try
                {
                    if (IsTruthy(WppCom.GetProperty(font, "Italic")))
                    {
                        cell.FontItalic = true;
                    }
                }
                catch
                {
                }

                try
                {
                    object color = WppCom.GetProperty(font, "Color");
                    object rgbObj = WppCom.GetProperty(color, "RGB");
                    if (rgbObj != null)
                    {
                        cell.FontColor = FormatRgb(Convert.ToInt32(rgbObj));
                    }
                }
                catch
                {
                }

                try
                {
                    object fill = WppCom.GetProperty(cellShape, "Fill");
                    if (IsTruthy(WppCom.GetProperty(fill, "Visible")))
                    {
                        object fore = WppCom.GetProperty(fill, "ForeColor");
                        object rgbObj = WppCom.GetProperty(fore, "RGB");
                        if (rgbObj != null)
                        {
                            cell.Fill = FormatRgb(Convert.ToInt32(rgbObj));
                        }
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

        private static void TryApplyColRowPctsWpp(
            object shape,
            object table,
            float[] colPcts,
            float[] rowPcts,
            List<string> warnings)
        {
            try
            {
                float tableW = Convert.ToSingle(WppCom.GetProperty(shape, "Width") ?? 0f);
                float tableH = Convert.ToSingle(WppCom.GetProperty(shape, "Height") ?? 0f);
                object colsObj = WppCom.GetProperty(table, "Columns");
                object rowsObj = WppCom.GetProperty(table, "Rows");
                int colCount = Convert.ToInt32(WppCom.GetProperty(colsObj, "Count"));
                int rowCount = Convert.ToInt32(WppCom.GetProperty(rowsObj, "Count"));
                if (colPcts != null && colPcts.Length == colCount && tableW > 0)
                {
                    float[] norm = NormalizePcts(colPcts);
                    for (int i = 0; i < norm.Length; i++)
                    {
                        object col = WppCom.GetIndexed(colsObj, i + 1);
                        WppCom.TrySetProperty(col, "Width", tableW * norm[i] / 100f);
                    }
                }

                if (rowPcts != null && rowPcts.Length == rowCount && tableH > 0)
                {
                    float[] norm = NormalizePcts(rowPcts);
                    for (int i = 0; i < norm.Length; i++)
                    {
                        object row = WppCom.GetIndexed(rowsObj, i + 1);
                        WppCom.TrySetProperty(row, "Height", tableH * norm[i] / 100f);
                    }
                }
            }
            catch (Exception ex)
            {
                warnings.Add("套列宽/行高失败: " + ex.Message);
            }
        }

        private static void TryReadColRowPctsWpp(
            object shape,
            object table,
            int rows,
            int cols,
            PptHtmlTableStyleSnap style,
            List<string> warnings)
        {
            try
            {
                float tableW = Convert.ToSingle(WppCom.GetProperty(shape, "Width") ?? 0f);
                float tableH = Convert.ToSingle(WppCom.GetProperty(shape, "Height") ?? 0f);
                object colsObj = WppCom.GetProperty(table, "Columns");
                object rowsObj = WppCom.GetProperty(table, "Rows");
                if (tableW > 0 && cols > 0)
                {
                    var pcts = new float[cols];
                    for (int c = 1; c <= cols; c++)
                    {
                        object col = WppCom.GetIndexed(colsObj, c);
                        float w = Convert.ToSingle(WppCom.GetProperty(col, "Width") ?? 0f);
                        pcts[c - 1] = w / tableW * 100f;
                    }

                    style.ColWidthPcts = pcts;
                }

                if (tableH > 0 && rows > 0)
                {
                    var pcts = new float[rows];
                    for (int r = 1; r <= rows; r++)
                    {
                        object row = WppCom.GetIndexed(rowsObj, r);
                        float h = Convert.ToSingle(WppCom.GetProperty(row, "Height") ?? 0f);
                        pcts[r - 1] = h / tableH * 100f;
                    }

                    style.RowHeightPcts = pcts;
                }
            }
            catch (Exception ex)
            {
                warnings.Add("读列宽/行高失败: " + ex.Message);
            }
        }

        private static void TryInferTableFontWpp(
            object table,
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

                object cell = WppCom.Invoke(table, "Cell", 1, 1);
                object cellShape = WppCom.GetProperty(cell, "Shape");
                object tf = WppCom.GetProperty(cellShape, "TextFrame");
                object tr = WppCom.GetProperty(tf, "TextRange");
                object font = WppCom.GetProperty(tr, "Font");
                try
                {
                    string name = Convert.ToString(WppCom.GetProperty(font, "Name"));
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
                    object sz = WppCom.GetProperty(font, "Size");
                    if (sz != null)
                    {
                        style.FontSizePt = Convert.ToDouble(sz);
                    }
                }
                catch
                {
                }

                try
                {
                    object color = WppCom.GetProperty(font, "Color");
                    object rgbObj = WppCom.GetProperty(color, "RGB");
                    if (rgbObj != null)
                    {
                        style.FontColor = FormatRgb(Convert.ToInt32(rgbObj));
                    }
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

        private static void TryApplyTableStyleWpp(
            object table,
            PptHtmlTableStyleSnap style,
            List<string> warnings)
        {
            try
            {
                object rowsObj = WppCom.GetProperty(table, "Rows");
                object colsObj = WppCom.GetProperty(table, "Columns");
                int rows = Convert.ToInt32(WppCom.GetProperty(rowsObj, "Count"));
                int cols = Convert.ToInt32(WppCom.GetProperty(colsObj, "Count"));
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
                                object cell = WppCom.Invoke(table, "Cell", r, c);
                                object cellShape = WppCom.GetProperty(cell, "Shape");
                                object tf = WppCom.GetProperty(cellShape, "TextFrame");
                                object tr = WppCom.GetProperty(tf, "TextRange");
                                object font = WppCom.GetProperty(tr, "Font");
                                if (!string.IsNullOrEmpty(style.FontName))
                                {
                                    WppCom.TrySetProperty(font, "Name", style.FontName);
                                }

                                if (style.FontSizePt.HasValue)
                                {
                                    WppCom.TrySetProperty(font, "Size", (float)style.FontSizePt.Value);
                                }

                                if (!string.IsNullOrEmpty(style.FontColor)
                                    && TryParseRgb(style.FontColor, out int rgb))
                                {
                                    object color = WppCom.GetProperty(font, "Color");
                                    WppCom.TrySetProperty(color, "RGB", rgb);
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
                    ApplyThreeLineWpp(table, warnings);
                }
            }
            catch (Exception ex)
            {
                warnings.Add("套整表样式失败: " + ex.Message);
            }
        }

        private static void ApplyThreeLineWpp(object table, List<string> warnings)
        {
            try
            {
                object rowsObj = WppCom.GetProperty(table, "Rows");
                object colsObj = WppCom.GetProperty(table, "Columns");
                int rows = Convert.ToInt32(WppCom.GetProperty(rowsObj, "Count"));
                int cols = Convert.ToInt32(WppCom.GetProperty(colsObj, "Count"));
                for (int r = 1; r <= rows; r++)
                {
                    for (int c = 1; c <= cols; c++)
                    {
                        object cell = WppCom.Invoke(table, "Cell", r, c);
                        SetBorderVisibleWpp(cell, "left", false);
                        SetBorderVisibleWpp(cell, "right", false);
                        SetBorderVisibleWpp(cell, "top", r == 1);
                        SetBorderVisibleWpp(cell, "bottom", r == 1 || r == rows);
                    }
                }
            }
            catch (Exception ex)
            {
                warnings.Add("套三线表失败: " + ex.Message);
            }
        }

        private static void SetBorderVisibleWpp(object cell, string which, bool visible)
        {
            try
            {
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

                object borders = WppCom.GetProperty(cell, "Borders");
                object line = WppCom.GetIndexed(borders, idx);
                WppCom.TrySetProperty(line, "Visible", visible ? -1 : 0);
                if (visible)
                {
                    WppCom.TrySetProperty(line, "Weight", 1f);
                    object fore = WppCom.GetProperty(line, "ForeColor");
                    WppCom.TrySetProperty(fore, "RGB", 0x000000);
                }
            }
            catch
            {
            }
        }

        private static string ReadCellTextWpp(object cell)
        {
            try
            {
                object cellShape = WppCom.GetProperty(cell, "Shape");
                object tf = WppCom.GetProperty(cellShape, "TextFrame");
                object tr = WppCom.GetProperty(tf, "TextRange");
                return Convert.ToString(WppCom.GetProperty(tr, "Text")) ?? "";
            }
            catch
            {
                return "";
            }
        }

        private static MergeMap BuildMergeMapWpp(
            object table,
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
                        object cell = WppCom.Invoke(table, "Cell", r, c);
                        object sh = WppCom.GetProperty(cell, "Shape");
                        id = Convert.ToString(WppCom.GetProperty(sh, "Id"), CultureInfo.InvariantCulture)
                            + "@"
                            + Convert.ToSingle(WppCom.GetProperty(sh, "Left") ?? 0f)
                                .ToString("0.###", CultureInfo.InvariantCulture)
                            + ","
                            + Convert.ToSingle(WppCom.GetProperty(sh, "Top") ?? 0f)
                                .ToString("0.###", CultureInfo.InvariantCulture)
                            + ","
                            + Convert.ToSingle(WppCom.GetProperty(sh, "Width") ?? 0f)
                                .ToString("0.###", CultureInfo.InvariantCulture)
                            + ","
                            + Convert.ToSingle(WppCom.GetProperty(sh, "Height") ?? 0f)
                                .ToString("0.###", CultureInfo.InvariantCulture);
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

        private static bool IsTruthy(object value)
        {
            if (value == null)
            {
                return false;
            }

            try
            {
                if (value is bool b)
                {
                    return b;
                }

                int n = Convert.ToInt32(value);
                return n == -1 || n == 1;
            }
            catch
            {
                return false;
            }
        }
    }
}
