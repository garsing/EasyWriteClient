using System;
using System.Collections.Generic;

namespace WordAddIn1
{
    /// <summary>HTML 铺表：可见格 + span → 逻辑网格。建表只认省略写法。</summary>
    public sealed class TableHtmlVisibleCell
    {
        public string Text { get; set; }
        public int RowSpan { get; set; } = 1;
        public int ColSpan { get; set; } = 1;
    }

    public sealed class TableHtmlGridResult
    {
        public bool Ok { get; set; }
        public string Error { get; set; }
        public int RowCount { get; set; }
        public int ColCount { get; set; }
        public List<List<string>> RectangularData { get; set; }
        public List<List<int>> Merge { get; set; }
    }

    public static class TableHtmlGrid
    {
        public static bool TryExpand(IList<IList<TableHtmlVisibleCell>> rows, out TableHtmlGridResult result)
        {
            result = new TableHtmlGridResult
            {
                Ok = false,
                RectangularData = new List<List<string>>(),
                Merge = new List<List<int>>()
            };

            if (rows == null || rows.Count == 0)
            {
                result.Error = "缺少 Data 或没有任何行";
                return false;
            }

            int rowCount = rows.Count;
            var occupied = new List<List<bool>>();
            EnsureRows(occupied, rowCount);
            int colCount = 0;

            for (int r = 0; r < rowCount; r++)
            {
                IList<TableHtmlVisibleCell> row = rows[r];
                if (row == null)
                {
                    result.Error = "第 " + (r + 1) + " 行为空";
                    return false;
                }

                if (HasColspanPlaceholders(row, r + 1, out string placeholderError))
                {
                    result.Error = placeholderError;
                    return false;
                }

                int cursor = 0;
                foreach (TableHtmlVisibleCell cell in row)
                {
                    if (cell == null)
                    {
                        result.Error = "第 " + (r + 1) + " 行存在空格";
                        return false;
                    }

                    int rs = cell.RowSpan < 1 ? 1 : cell.RowSpan;
                    int cs = cell.ColSpan < 1 ? 1 : cell.ColSpan;

                    while (true)
                    {
                        EnsureWidth(occupied, cursor + 1);
                        if (colCount < cursor + 1)
                        {
                            colCount = cursor + 1;
                        }

                        if (cursor < colCount && IsOccupied(occupied, r, cursor))
                        {
                            cursor++;
                            continue;
                        }

                        break;
                    }

                    if (r + rs > rowCount)
                    {
                        result.Error = "第 " + (r + 1) + " 行合并越界（rowspan=" + rs + "）";
                        return false;
                    }

                    if (r > 0 && cursor >= colCount)
                    {
                        result.Error = "第 " + (r + 1) + " 行多写了占用槽";
                        return false;
                    }

                    if (r > 0 && cursor + cs > colCount)
                    {
                        result.Error = "第 " + (r + 1) + " 行合并越界（colspan=" + cs + "）";
                        return false;
                    }

                    EnsureRows(occupied, r + rs);
                    EnsureWidth(occupied, cursor + cs);

                    for (int i = 0; i < rs; i++)
                    {
                        for (int j = 0; j < cs; j++)
                        {
                            if (IsOccupied(occupied, r + i, cursor + j))
                            {
                                result.Error = "第 " + (r + 1) + " 行第 " + (cursor + 1) + " 列与已有合并重叠（rowspan="
                                    + rs + "，colspan=" + cs + "）";
                                return false;
                            }
                        }
                    }

                    for (int i = 0; i < rs; i++)
                    {
                        for (int j = 0; j < cs; j++)
                        {
                            occupied[r + i][cursor + j] = true;
                        }
                    }

                    if (rs > 1 || cs > 1)
                    {
                        result.Merge.Add(new List<int>
                        {
                            r + 1,
                            cursor + 1,
                            r + rs,
                            cursor + cs
                        });
                    }

                    cursor += cs;
                    if (cursor > colCount)
                    {
                        colCount = cursor;
                    }
                }

                EnsureWidth(occupied, colCount);
                for (int c = 0; c < colCount; c++)
                {
                    if (!IsOccupied(occupied, r, c))
                    {
                        result.Error = "第 " + (r + 1) + " 行可见格未铺满逻辑列";
                        return false;
                    }
                }
            }

            EnsureWidth(occupied, colCount);
            for (int r = 0; r < rowCount; r++)
            {
                for (int c = 0; c < colCount; c++)
                {
                    if (!IsOccupied(occupied, r, c))
                    {
                        result.Error = "第 " + (r + 1) + " 行可见格未铺满逻辑列";
                        return false;
                    }
                }
            }

            if (colCount <= 0)
            {
                result.Error = "无法从可见格推断列数";
                return false;
            }

            var data = new List<List<string>>(rowCount);
            for (int r = 0; r < rowCount; r++)
            {
                var line = new List<string>(colCount);
                for (int c = 0; c < colCount; c++)
                {
                    line.Add("");
                }

                data.Add(line);
            }

            var occupiedPlace = NewOccupied(rowCount, colCount);
            for (int r = 0; r < rowCount; r++)
            {
                int cursor = 0;
                foreach (TableHtmlVisibleCell cell in rows[r])
                {
                    int rs = cell.RowSpan < 1 ? 1 : cell.RowSpan;
                    int cs = cell.ColSpan < 1 ? 1 : cell.ColSpan;
                    while (cursor < colCount && occupiedPlace[r, cursor])
                    {
                        cursor++;
                    }

                    data[r][cursor] = cell.Text ?? "";
                    for (int i = 0; i < rs; i++)
                    {
                        for (int j = 0; j < cs; j++)
                        {
                            occupiedPlace[r + i, cursor + j] = true;
                        }
                    }

                    cursor += cs;
                }
            }

            result.Ok = true;
            result.RowCount = rowCount;
            result.ColCount = colCount;
            result.RectangularData = data;
            return true;
        }

        /// <summary>colspan=N 后再写 N-1 个空格，视为占位写法。</summary>
        private static bool HasColspanPlaceholders(
            IList<TableHtmlVisibleCell> row,
            int row1Based,
            out string error)
        {
            error = null;
            for (int i = 0; i < row.Count; i++)
            {
                TableHtmlVisibleCell cell = row[i];
                if (cell == null)
                {
                    continue;
                }

                int cs = cell.ColSpan < 1 ? 1 : cell.ColSpan;
                if (cs <= 1)
                {
                    continue;
                }

                int need = cs - 1;
                if (i + need >= row.Count)
                {
                    continue;
                }

                bool allEmpty = true;
                for (int k = 1; k <= need; k++)
                {
                    TableHtmlVisibleCell next = row[i + k];
                    if (next == null
                        || next.ColSpan > 1
                        || next.RowSpan > 1
                        || !string.IsNullOrWhiteSpace(next.Text))
                    {
                        allEmpty = false;
                        break;
                    }
                }

                if (allEmpty)
                {
                    error = "第 " + row1Based + " 行第 " + (i + 2) + " 列与已有合并重叠（rowspan=1，colspan="
                        + cs + "）";
                    return true;
                }
            }

            return false;
        }

        private static bool IsOccupied(List<List<bool>> occupied, int r, int c)
        {
            if (r < 0 || r >= occupied.Count)
            {
                return false;
            }

            List<bool> row = occupied[r];
            if (c < 0 || c >= row.Count)
            {
                return false;
            }

            return row[c];
        }

        private static void EnsureRows(List<List<bool>> occupied, int rowCount)
        {
            while (occupied.Count < rowCount)
            {
                occupied.Add(new List<bool>());
            }
        }

        private static void EnsureWidth(List<List<bool>> occupied, int width)
        {
            for (int r = 0; r < occupied.Count; r++)
            {
                while (occupied[r].Count < width)
                {
                    occupied[r].Add(false);
                }
            }
        }

        private static bool[,] NewOccupied(int rows, int cols)
        {
            return new bool[rows, cols];
        }
    }
}
