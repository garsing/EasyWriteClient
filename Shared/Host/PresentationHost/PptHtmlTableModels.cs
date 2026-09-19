using System.Collections.Generic;

namespace WordAddIn1.PresentationHost
{
    /// <summary>页内 table 形状网格（含合并与格皮）。与 chart 内嵌灌数表分离。</summary>
    internal sealed class PptHtmlTableGrid
    {
        public const int MaxRows = 40;

        public const int MaxCols = 16;

        public int RowCount { get; set; }

        public int ColCount { get; set; }

        public List<PptHtmlTableCell> Cells { get; set; }

        public PptHtmlTableCell CellAt(int row, int col)
        {
            if (Cells == null || row < 0 || col < 0 || row >= RowCount || col >= ColCount)
            {
                return null;
            }

            int i = row * ColCount + col;
            if (i < 0 || i >= Cells.Count)
            {
                return null;
            }

            return Cells[i];
        }

        public List<List<string>> ToPlainTextMatrix()
        {
            var rows = new List<List<string>>();
            for (int r = 0; r < RowCount; r++)
            {
                var row = new List<string>();
                for (int c = 0; c < ColCount; c++)
                {
                    PptHtmlTableCell cell = CellAt(r, c);
                    row.Add(cell == null || cell.IsCovered ? "" : (cell.Text ?? ""));
                }

                rows.Add(row);
            }

            return rows;
        }
    }

    internal sealed class PptHtmlTableCell
    {
        public string Text { get; set; }

        public int RowSpan { get; set; } = 1;

        public int ColSpan { get; set; } = 1;

        /// <summary>被合并占位的格（HTML 空 td）。</summary>
        public bool IsCovered { get; set; }

        public bool IsHeader { get; set; }

        /// <summary>null=不改；none=无填充；#RRGGBB</summary>
        public string Fill { get; set; }

        /// <summary>null=不改；#RRGGBB</summary>
        public string FontColor { get; set; }

        public bool? FontBold { get; set; }

        public bool? FontItalic { get; set; }
    }

    /// <summary>表级皮；字段 null=不改。</summary>
    internal sealed class PptHtmlTableStyleSnap
    {
        public string FontName { get; set; }

        public double? FontSizePt { get; set; }

        public string FontColor { get; set; }

        /// <summary>本期仅 three-line；null=不改</summary>
        public string TableStyle { get; set; }

        /// <summary>相对表宽 %，长度=列数；null=不改</summary>
        public float[] ColWidthPcts { get; set; }

        /// <summary>相对表高 %，长度=行数；null=不改</summary>
        public float[] RowHeightPcts { get; set; }

        public bool Truncated { get; set; }
    }
}
