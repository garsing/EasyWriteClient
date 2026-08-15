using System.Collections.Generic;

namespace WordAddIn1.SpreadsheetHost
{
    public sealed class WorkbookSheetInfo
    {
        public string Name { get; set; }

        public bool Hidden { get; set; }

        public string SheetType { get; set; }

        public string UsedRange { get; set; }

        public int LastRow { get; set; }

        public string LastCol { get; set; }

        public string Preview { get; set; }

        public bool PreviewTruncated { get; set; }

        public string Note { get; set; }
    }

    public sealed class WorkbookContentResult
    {
        public string ChannelId { get; set; }

        public string Kind { get; set; }

        public string Name { get; set; }

        public string Path { get; set; }

        public List<WorkbookSheetInfo> Sheets { get; set; }
    }
}
