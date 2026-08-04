using System.Collections.Generic;

namespace WordAddIn1
{
    /// <summary>
    /// TableConfig XML 格式 apply 侧模型（General=可迁移格式；Table=结构/数据，apply 时忽略）。
    /// </summary>
    public class TableGeneralFormat
    {
        public string TableFontName { get; set; }
        public float? TableFontSize { get; set; }
        public string TableFontColor { get; set; }
        public TableStyleConfig StyleConfig { get; set; }
    }

    public class TableStyleConfig
    {
        public string TableStyle { get; set; }
        public int HeaderRows { get; set; } = 1;
    }

    public class TablePropertiesFormat
    {
        public int? Rows { get; set; }
        public int? Cols { get; set; }
    }

    public class TableApplyConfig
    {
        public TableGeneralFormat General { get; set; }
        public TablePropertiesFormat Properties { get; set; }
        /// <summary>XML 位于 Table/Properties/ColWidths；仅提取/建表用，精迁 apply 不写。</summary>
        public List<float> ColWidths { get; set; }
        public List<List<int>> Merge { get; set; }
    }

    public class TableConfigApplyResult
    {
        public bool Success { get; set; }
        public string Error { get; set; }
        public List<string> Warnings { get; set; } = new List<string>();
        /// <summary>来自 Range.Cells 锚点，兼容纵向合并表。</summary>
        public int? RowCount { get; set; }
        public int? ColumnCount { get; set; }
    }

    public class TableDataApplyConfig
    {
        public TablePropertiesFormat Properties { get; set; } = new TablePropertiesFormat();
        public List<float> ColWidths { get; set; }
        public List<List<int>> Merge { get; set; } = new List<List<int>>();
        public List<List<string>> Data { get; set; } = new List<List<string>>();
    }

    public class TableDataApplyResult
    {
        public bool Success { get; set; }
        public string Error { get; set; }
        public string Mode { get; set; }
        public int CellsWritten { get; set; }
        public int CellsSkipped { get; set; }
        public List<string> Warnings { get; set; } = new List<string>();
    }
}
