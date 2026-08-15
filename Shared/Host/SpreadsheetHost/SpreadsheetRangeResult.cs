using System.Collections.Generic;

namespace WordAddIn1.SpreadsheetHost
{
    internal sealed class SpreadsheetRangeResult
    {
        public string ChannelId { get; set; }

        public string Kind { get; set; }

        public string Sheet { get; set; }

        public string RequestedRange { get; set; }

        public string ActualRange { get; set; }

        public bool Truncated { get; set; }

        public string TruncatedReason { get; set; }

        public bool IncludeFormulas { get; set; }

        public List<PreviewRow> Rows { get; set; }
    }

    internal static class SpreadsheetRangeLimits
    {
        public const int MaxRows = 100;
        public const int MaxCols = 40;
        public const int MaxCells = 4000;

        public static void Apply(
            int firstRow,
            int firstCol,
            int lastRow,
            int lastCol,
            out int actualLastRow,
            out int actualLastCol,
            out bool truncated,
            out string reason)
        {
            actualLastRow = lastRow;
            actualLastCol = lastCol;
            truncated = false;
            reason = "";

            int rowCount = lastRow - firstRow + 1;
            int colCount = lastCol - firstCol + 1;
            if (rowCount <= 0 || colCount <= 0)
            {
                actualLastRow = firstRow - 1;
                actualLastCol = firstCol - 1;
                return;
            }

            if (rowCount > MaxRows)
            {
                actualLastRow = firstRow + MaxRows - 1;
                truncated = true;
                reason = "max_rows";
                rowCount = MaxRows;
            }

            if (colCount > MaxCols)
            {
                actualLastCol = firstCol + MaxCols - 1;
                truncated = true;
                reason = string.IsNullOrEmpty(reason) ? "max_cols" : reason + ",max_cols";
                colCount = MaxCols;
            }

            long cells = (long)rowCount * colCount;
            if (cells > MaxCells)
            {
                int maxRowsForCells = MaxCells / colCount;
                if (maxRowsForCells < 1)
                {
                    maxRowsForCells = 1;
                    actualLastCol = firstCol;
                    colCount = 1;
                }

                actualLastRow = firstRow + maxRowsForCells - 1;
                truncated = true;
                reason = string.IsNullOrEmpty(reason) ? "max_cells" : reason + ",max_cells";
            }
        }
    }
}
