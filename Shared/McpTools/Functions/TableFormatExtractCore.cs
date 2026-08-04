using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    public class TableExtractDto
    {
        public int Rows { get; set; }
        public int Cols { get; set; }
        public List<float> ColWidths { get; set; } = new List<float>();
        public string Style { get; set; }
        public TableStyleConfig StyleConfig { get; set; }
        public TableFontStatsResult FontStats { get; set; }
        /// <summary>Data 矩阵 span 默认色（来自 merge 提取），非 General 三字体。</summary>
        public string DefaultSpanColor { get; set; }
        public List<List<int>> Merge { get; set; } = new List<List<int>>();
        public List<List<string>> Data { get; set; } = new List<List<string>>();
        public int RangeStart { get; set; }
        public int RangeEnd { get; set; }
    }

    public static class TableFormatExtractCore
    {
        private static readonly XNamespace W =
            "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

        public static TableExtractDto Extract(Word.Table table, Word.Document document)
        {
            if (table == null)
            {
                throw new ArgumentNullException(nameof(table));
            }

            ExtractStructureFromOpenXml(table, out XElement tableElement, out int actualRows, out int actualCols);

            string tableStyle = TableFormatExtractor.ExtractTableStyleFromXml(tableElement, W, table);
            List<float> colWidths = TableFormatExtractor.ExtractColWidthsFromXml(tableElement, W, actualCols);

            var merge = new List<List<int>>();
            var data = new List<List<string>>();
            string defaultSpanColor = TableFormatExtractor.ExtractMergeInfoAndDataFromXml(
                tableElement, W, actualRows, actualCols, merge, data);

            var fontStats = TableFontStatsHelper.ComputeModeTriplet(table, document);

            return new TableExtractDto
            {
                Rows = actualRows,
                Cols = actualCols,
                ColWidths = colWidths ?? new List<float>(),
                Style = tableStyle,
                StyleConfig = new TableStyleConfig
                {
                    TableStyle = tableStyle,
                    HeaderRows = 1,
                },
                FontStats = fontStats,
                DefaultSpanColor = defaultSpanColor,
                Merge = merge ?? new List<List<int>>(),
                Data = data ?? new List<List<string>>(),
                RangeStart = table.Range.Start,
                RangeEnd = table.Range.End,
            };
        }

        /// <summary>仅提取行列与 Merge（供 apply data 结构校验）。</summary>
        public static void ExtractStructure(
            Word.Table table,
            out int rows,
            out int cols,
            out List<List<int>> merge)
        {
            if (table == null)
            {
                throw new ArgumentNullException(nameof(table));
            }

            ExtractStructureFromOpenXml(table, out XElement tableElement, out rows, out cols);
            merge = new List<List<int>>();
            var data = new List<List<string>>();
            TableFormatExtractor.ExtractMergeInfoAndDataFromXml(
                tableElement, W, rows, cols, merge, data);
            cols = ResolveLogicalColumnCount(tableElement, W, table, merge, cols);
        }

        /// <summary>
        /// OpenXML tblGrid 逻辑列数（与 merge / EnumerateRowSlots 坐标一致）。
        /// </summary>
        public static int GetLogicalGridColumnCount(XElement tableElement, XNamespace w)
        {
            if (tableElement == null)
            {
                return 0;
            }

            XElement tblGrid = tableElement.Element(w + "tblGrid");
            if (tblGrid == null)
            {
                return 0;
            }

            return tblGrid.Elements(w + "gridCol").Count();
        }

        /// <summary>
        /// 逻辑列上界 = max(tblGrid 列数, Word.Columns.Count, merge 最大 endCol)。
        /// </summary>
        public static int ResolveLogicalColumnCount(
            XElement tableElement,
            XNamespace w,
            Word.Table table,
            List<List<int>> merge,
            int preliminaryCols)
        {
            int cols = preliminaryCols;
            int gridCols = GetLogicalGridColumnCount(tableElement, w);
            if (gridCols > cols)
            {
                cols = gridCols;
            }

            try
            {
                if (table != null && table.Columns != null)
                {
                    cols = Math.Max(cols, table.Columns.Count);
                }
            }
            catch (Exception)
            {
            }

            if (merge != null)
            {
                foreach (List<int> item in merge)
                {
                    if (item == null || item.Count < 4)
                    {
                        continue;
                    }

                    cols = Math.Max(cols, item[3]);
                }
            }

            return cols;
        }

        private static void ExtractStructureFromOpenXml(
            Word.Table table,
            out XElement tableElement,
            out int actualRows,
            out int actualCols)
        {
            string tableXml = table.Range.WordOpenXML;
            if (string.IsNullOrEmpty(tableXml))
            {
                throw new Exception("无法获取表格的 XML 内容");
            }

            XDocument xmlDoc = XDocument.Parse(tableXml);
            tableElement = xmlDoc.Descendants(W + "tbl").FirstOrDefault();
            if (tableElement == null)
            {
                throw new Exception("无法找到表格元素");
            }

            actualRows = tableElement.Elements(W + "tr").Count();
            int wordCols = table.Columns.Count;
            int gridCols = GetLogicalGridColumnCount(tableElement, W);
            actualCols = gridCols > 0 ? Math.Max(gridCols, wordCols) : wordCols;
        }
    }
}
