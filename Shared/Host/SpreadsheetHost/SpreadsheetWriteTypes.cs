using System.Collections.Generic;

namespace WordAddIn1.SpreadsheetHost
{
    internal enum SpreadsheetWriteMode
    {
        Csv = 0,
        Clear = 1
    }

    internal sealed class SpreadsheetWriteRequest
    {
        public string SheetName { get; set; }

        public SpreadsheetWriteMode Mode { get; set; }

        /// <summary>单格锚点或矩形（纯 A1）。</summary>
        public string RangeA1 { get; set; }

        /// <summary>mode=Csv 时已解析网格。</summary>
        public IReadOnlyList<IReadOnlyList<string>> CsvGrid { get; set; }
    }

    internal sealed class SpreadsheetWriteResult
    {
        public string ChannelId { get; set; }

        public string Kind { get; set; }

        public string Sheet { get; set; }

        public string Mode { get; set; }

        public int WrittenCount { get; set; }

        public string ActualRange { get; set; }

        public string CsvFilename { get; set; }
    }
}
