using System.Collections.Generic;

namespace WordAddIn1.SpreadsheetHost
{
    internal sealed class SpreadsheetGetFormatRequest
    {
        public string SheetName { get; set; }

        public string RangeA1 { get; set; }
    }

    internal sealed class SpreadsheetGetFormatResult
    {
        public string ChannelId { get; set; }

        public string Kind { get; set; }

        public string Sheet { get; set; }

        public string RequestedRange { get; set; }

        public string ActualRange { get; set; }

        /// <summary>cells | column | columns</summary>
        public string Mode { get; set; }

        public List<SpreadsheetFormatItem> Items { get; set; }
    }

    internal sealed class SpreadsheetFormatItem
    {
        public string Addr { get; set; }

        public string Col { get; set; }

        public string RangeA1 { get; set; }

        public bool? Mixed { get; set; }

        public List<string> MixedFields { get; set; }

        public SpreadsheetFormatPatch Format { get; set; }
    }

    internal static class SpreadsheetGetFormatLimits
    {
        public const int MaxCellsForCellsMode = 40;

        public static string ResolveMode(int rowCount, int colCount)
        {
            long cells = (long)rowCount * colCount;
            if (cells <= MaxCellsForCellsMode)
            {
                return "cells";
            }

            if (colCount == 1)
            {
                return "column";
            }

            return "columns";
        }
    }
}
