using System;
using System.Collections.Generic;
using System.Linq;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    public static class TableMergeHelper
    {
        /// <summary>
        /// 判断 (row,col) 是否为 merge 延续格（1-based）。
        /// </summary>
        public static bool IsMergeContinuation(int row, int col, List<List<int>> merge)
        {
            if (merge == null || merge.Count == 0)
            {
                return false;
            }

            foreach (var item in merge)
            {
                if (item == null || item.Count < 4)
                {
                    continue;
                }

                int startRow = item[0];
                int startCol = item[1];
                int endRow = item[2];
                int endCol = item[3];

                if (row >= startRow && row <= endRow && col >= startCol && col <= endCol)
                {
                    if (row == startRow && col == startCol)
                    {
                        return false;
                    }

                    return true;
                }
            }

            return false;
        }

        /// <summary>
        /// I22：统计待补普通空格（不含 merge 延续格）。
        /// </summary>
        public static int CountFillableEmptyCells(List<List<string>> data, List<List<int>> merge)
        {
            if (data == null || data.Count == 0)
            {
                return 0;
            }

            int count = 0;
            for (int r = 0; r < data.Count; r++)
            {
                var row = data[r];
                if (row == null)
                {
                    continue;
                }

                for (int c = 0; c < row.Count; c++)
                {
                    if (IsMergeContinuation(r + 1, c + 1, merge))
                    {
                        continue;
                    }

                    if (IsNormalizedCellEmpty(row[c]))
                    {
                        count++;
                    }
                }
            }

            return count;
        }

        public static bool IsNormalizedCellEmpty(string cellContent)
        {
            if (cellContent == null)
            {
                return true;
            }

            string normalized = cellContent
                .Replace("\r", "")
                .Replace("\a", "")
                .Trim();
            return string.IsNullOrEmpty(normalized);
        }

        public static bool IsSameMerge(List<List<int>> a, List<List<int>> b)
        {
            return NormalizeMergeKey(a) == NormalizeMergeKey(b);
        }

        public static bool ValidateStructure(
            Word.Table table,
            int? xmlRows,
            int? xmlCols,
            List<List<int>> xmlMerge,
            out string errorCode,
            out string message)
        {
            errorCode = null;
            message = null;

            GetTableDimensions(table, out int? tableRows, out int? tableCols);
            TableFormatExtractCore.ExtractStructure(table, out int extractRows, out int extractCols, out List<List<int>> tableMerge);

            int expectedRows = xmlRows ?? extractRows;
            int expectedCols = xmlCols ?? extractCols;

            if (tableRows.HasValue && tableRows.Value != expectedRows)
            {
                errorCode = "table_data_structure_mismatch";
                message = $"Rows 不一致：Word={tableRows.Value}，XML={expectedRows}";
                return false;
            }

            if (tableCols.HasValue && tableCols.Value != expectedCols)
            {
                errorCode = "table_data_structure_mismatch";
                message = $"Cols 不一致：Word={tableCols.Value}，XML={expectedCols}";
                return false;
            }

            if (xmlMerge != null && xmlMerge.Count > 0 && !IsSameMerge(xmlMerge, tableMerge))
            {
                errorCode = "table_data_structure_mismatch";
                message = "Merge 与现表不一致";
                return false;
            }

            return true;
        }

        public static bool PreflightMergeContinuationWrites(
            List<List<string>> data,
            List<List<int>> merge,
            out int row,
            out int col,
            out string message)
        {
            row = 0;
            col = 0;
            message = null;

            if (data == null)
            {
                return true;
            }

            for (int r = 0; r < data.Count; r++)
            {
                var dataRow = data[r];
                if (dataRow == null)
                {
                    continue;
                }

                for (int c = 0; c < dataRow.Count; c++)
                {
                    if (!IsMergeContinuation(r + 1, c + 1, merge))
                    {
                        continue;
                    }

                    if (!IsNormalizedCellEmpty(dataRow[c]))
                    {
                        row = r + 1;
                        col = c + 1;
                        message = $"invalid_merge_cell: 不能对 merge 延续格 ({row},{col}) 写入";
                        return false;
                    }
                }
            }

            return true;
        }

        public static void GetTableDimensions(Word.Table table, out int? rows, out int? cols)
        {
            rows = null;
            cols = null;
            if (table == null)
            {
                return;
            }

            try
            {
                int maxRow = 0;
                int maxCol = 0;
                foreach (Word.Cell cell in table.Range.Cells)
                {
                    maxRow = Math.Max(maxRow, cell.RowIndex);
                    maxCol = Math.Max(maxCol, cell.ColumnIndex);
                }

                if (maxRow > 0)
                {
                    rows = maxRow;
                }

                if (maxCol > 0)
                {
                    cols = maxCol;
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[TableMergeHelper] 读取行列失败: {ex.Message}");
            }
        }

        private static string NormalizeMergeKey(List<List<int>> merge)
        {
            if (merge == null || merge.Count == 0)
            {
                return "";
            }

            var normalized = merge
                .Where(m => m != null && m.Count >= 4)
                .Select(m => $"{m[0]}:{m[1]}:{m[2]}:{m[3]}")
                .OrderBy(s => s, StringComparer.Ordinal)
                .ToList();
            return string.Join("|", normalized);
        }
    }
}
