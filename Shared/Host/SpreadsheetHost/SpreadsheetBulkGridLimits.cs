namespace WordAddIn1.SpreadsheetHost
{
    /// <summary>
    /// 写 / 导 CSV / 套格式 / 条件格式 / 透视源 / 图源：失败不截断。
    /// 不带 path 的读截断、截图仍用 SpreadsheetRangeLimits。
    /// </summary>
    internal static class SpreadsheetBulkGridLimits
    {
        public const int MaxRows = 50000;

        public const int MaxCols = 50;

        public const int MaxCells = 2500000;

        public const int WriteChunkRows = 2000;

        public static string CapText()
        {
            return MaxRows + "×" + MaxCols + " / " + MaxCells + " 格";
        }

        public static bool TryCheck(int rowCount, int colCount, string actionLabel, out string error)
        {
            error = null;
            if (rowCount > MaxRows
                || colCount > MaxCols
                || (long)rowCount * colCount > MaxCells)
            {
                string label = string.IsNullOrEmpty(actionLabel) ? "操作" : actionLabel;
                error = label + "区域过大（最多 " + CapText() + "）";
                return false;
            }

            return true;
        }

        public static bool TryCheckExport(int rowCount, int colCount, out string error)
        {
            if (TryCheck(rowCount, colCount, "导出", out error))
            {
                return true;
            }

            error = "导出区域过大（最多 " + CapText()
                + "）。请把 range 拆成多次 F_read_excel_range（每次不超过该上限）再导出。";
            return false;
        }
    }
}
