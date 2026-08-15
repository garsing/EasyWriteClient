using System;
using System.Collections.Generic;
using System.Globalization;

namespace WordAddIn1.SpreadsheetHost
{
    internal static class SpreadsheetWritePlanner
    {
        public static bool TryNormalizeGrid(
            IReadOnlyList<IReadOnlyList<string>> raw,
            out List<List<string>> grid,
            out int rowCount,
            out int colCount,
            out string error)
        {
            grid = null;
            rowCount = 0;
            colCount = 0;
            error = null;
            if (raw == null || raw.Count == 0)
            {
                error = "CSV 为空或解析失败";
                return false;
            }

            int maxCols = 0;
            foreach (IReadOnlyList<string> row in raw)
            {
                if (row != null && row.Count > maxCols)
                {
                    maxCols = row.Count;
                }
            }

            if (maxCols <= 0)
            {
                error = "CSV 为空或解析失败";
                return false;
            }

            grid = new List<List<string>>(raw.Count);
            foreach (IReadOnlyList<string> row in raw)
            {
                var line = new List<string>(maxCols);
                int n = row == null ? 0 : row.Count;
                for (int c = 0; c < maxCols; c++)
                {
                    line.Add(c < n ? (row[c] ?? "") : "");
                }

                grid.Add(line);
            }

            rowCount = grid.Count;
            colCount = maxCols;
            return true;
        }

        public static bool TryResolveWriteRect(
            string rangeA1,
            int csvRows,
            int csvCols,
            out int firstRow,
            out int firstCol,
            out int lastRow,
            out int lastCol,
            out string actualRange,
            out string error)
        {
            firstRow = firstCol = lastRow = lastCol = 0;
            actualRange = "";
            error = null;
            if (!A1Address.TryParseRange(
                    rangeA1,
                    out firstRow,
                    out firstCol,
                    out int reqLastRow,
                    out int reqLastCol,
                    out error))
            {
                return false;
            }

            bool single = firstRow == reqLastRow && firstCol == reqLastCol;
            if (single)
            {
                lastRow = firstRow + csvRows - 1;
                lastCol = firstCol + csvCols - 1;
            }
            else
            {
                int reqRows = reqLastRow - firstRow + 1;
                int reqCols = reqLastCol - firstCol + 1;
                if (reqRows != csvRows || reqCols != csvCols)
                {
                    error = "CSV 尺寸与 range 不一致（CSV "
                        + csvRows + "×" + csvCols
                        + "，range " + reqRows + "×" + reqCols + "）";
                    return false;
                }

                lastRow = reqLastRow;
                lastCol = reqLastCol;
            }

            if (!TryCheckLimits(csvRows, csvCols, out error))
            {
                return false;
            }

            actualRange = A1Address.Range(firstRow, firstCol, lastRow, lastCol);
            return true;
        }

        public static bool TryCheckLimits(int rowCount, int colCount, out string error)
        {
            return TryCheckLimits(rowCount, colCount, "写入", out error);
        }

        public static bool TryCheckLimits(int rowCount, int colCount, string actionLabel, out string error)
        {
            error = null;
            if (rowCount > SpreadsheetRangeLimits.MaxRows
                || colCount > SpreadsheetRangeLimits.MaxCols
                || (long)rowCount * colCount > SpreadsheetRangeLimits.MaxCells)
            {
                string label = string.IsNullOrEmpty(actionLabel) ? "操作" : actionLabel;
                error = label + "区域过大（最多 "
                    + SpreadsheetRangeLimits.MaxRows + "×"
                    + SpreadsheetRangeLimits.MaxCols + " / "
                    + SpreadsheetRangeLimits.MaxCells + " 格）";
                return false;
            }

            return true;
        }

        /// <summary>
        /// 判定格写入方式。
        /// </summary>
        public static void ClassifyCell(
            string raw,
            out bool clearOnly,
            out bool isFormula,
            out object value)
        {
            clearOnly = false;
            isFormula = false;
            value = null;
            if (raw == null || string.IsNullOrWhiteSpace(raw))
            {
                clearOnly = true;
                return;
            }

            if (raw.StartsWith("'", StringComparison.Ordinal))
            {
                value = raw.Substring(1);
                return;
            }

            if (raw.StartsWith("=", StringComparison.Ordinal))
            {
                isFormula = true;
                value = raw;
                return;
            }

            if (bool.TryParse(raw, out bool b))
            {
                value = b;
                return;
            }

            if (double.TryParse(
                    raw,
                    NumberStyles.Float,
                    CultureInfo.InvariantCulture,
                    out double d)
                || double.TryParse(
                    raw,
                    NumberStyles.Float,
                    CultureInfo.CurrentCulture,
                    out d))
            {
                value = d;
                return;
            }

            value = raw;
        }

        /// <summary>
        /// 合并区非左上角：CSV 必须为空（空白视为空）。
        /// </summary>
        public static bool IsEmptyForMergeCheck(string raw)
        {
            return string.IsNullOrWhiteSpace(raw);
        }

        public static string FormatMergeCsvConflictError(string cellAddr, string mergeAreaA1)
        {
            return "CSV 与合并结构不符："
                + (cellAddr ?? "")
                + " 属于合并区 "
                + (mergeAreaA1 ?? "")
                + "，仅左上角可有值，其余格须为空（请改 CSV 或先取消合并）";
        }
    }
}
