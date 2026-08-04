using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    /// <summary>
    /// 将 OpenXML 逻辑列 (1-based，与 merge / EnumerateRowSlots 一致) 映射到 Word.Cell。
    /// 复杂 merge 表中 Word COM 的 ColumnIndex 与逻辑列不对齐，不能直接用 table.Cell(row,col)。
    /// </summary>
    public static class TableLogicalCellMap
    {
        private static readonly XNamespace W =
            "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

        public static Dictionary<(int Row, int Col), Word.Cell> Build(Word.Table table)
        {
            var map = new Dictionary<(int Row, int Col), Word.Cell>();
            if (table == null)
            {
                return map;
            }

            string tableXml;
            try
            {
                tableXml = table.Range.WordOpenXML;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[TableLogicalCellMap] 无法读取表格 XML: {ex.Message}");
                return map;
            }

            if (string.IsNullOrEmpty(tableXml))
            {
                return map;
            }

            XElement tableElement;
            List<XElement> xmlRows;
            try
            {
                XDocument doc = XDocument.Parse(tableXml);
                tableElement = doc.Descendants(W + "tbl").FirstOrDefault();
                if (tableElement == null)
                {
                    return map;
                }

                xmlRows = tableElement.Elements(W + "tr").ToList();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[TableLogicalCellMap] 解析表格 XML 失败: {ex.Message}");
                return map;
            }

            Dictionary<int, List<Word.Cell>> cellsByRow = CollectPhysicalCellsByRow(table);
            var rowspanMap = TableFormatExtractor.CalculateAllRowspansFromXml(xmlRows, W);
            var colorCounts = new Dictionary<string, int>();
            int totalRuns = 0;

            for (int rowIdx = 0; rowIdx < xmlRows.Count; rowIdx++)
            {
                int wordRow = rowIdx + 1;
                if (!cellsByRow.TryGetValue(wordRow, out List<Word.Cell> rowCells))
                {
                    rowCells = new List<Word.Cell>();
                }

                int physicalIdx = 0;
                int logicalCol = 0;
                foreach (XElement tc in xmlRows[rowIdx].Elements(W + "tc"))
                {
                    TableFormatExtractor.CellInfo cellInfo = TableFormatExtractor.ProcessCellFromXml(
                        tc,
                        W,
                        rowIdx,
                        logicalCol,
                        rowspanMap,
                        colorCounts,
                        ref totalRuns);

                    if (cellInfo.IsVerticalMergeContinue)
                    {
                        logicalCol += cellInfo.Colspan;
                        continue;
                    }

                    if (physicalIdx < rowCells.Count)
                    {
                        map[(wordRow, logicalCol + 1)] = rowCells[physicalIdx];
                        physicalIdx++;
                    }
                    else
                    {
                        EasyWriteDiagnostics.LogTableRowValues(
                            $"[TableLogicalCellMap] row={wordRow} logicalCol={logicalCol + 1} 缺少物理单元格 "
                            + $"(physicalIdx={physicalIdx}, rowCells={rowCells.Count})",
                            verboseOnly: true);
                    }

                    logicalCol += cellInfo.Colspan;
                }

                if (physicalIdx != rowCells.Count)
                {
                    EasyWriteDiagnostics.LogTableRowValues(
                        $"[TableLogicalCellMap] row={wordRow} XML/物理单元格数量不一致 "
                        + $"(used={physicalIdx}, physical={rowCells.Count})",
                        verboseOnly: true);
                }
            }

            return map;
        }

        private static Dictionary<int, List<Word.Cell>> CollectPhysicalCellsByRow(Word.Table table)
        {
            var byRow = new Dictionary<int, List<Word.Cell>>();
            try
            {
                foreach (Word.Cell cell in table.Range.Cells)
                {
                    try
                    {
                        int row = cell.RowIndex;
                        if (!byRow.TryGetValue(row, out List<Word.Cell> list))
                        {
                            list = new List<Word.Cell>();
                            byRow[row] = list;
                        }

                        list.Add(cell);
                    }
                    catch (System.Runtime.InteropServices.COMException)
                    {
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[TableLogicalCellMap] Range.Cells 异常: {ex.Message}");
            }

            if (byRow.Count == 0)
            {
                try
                {
                    foreach (Word.Row tableRow in table.Rows)
                    {
                        foreach (Word.Cell cell in tableRow.Cells)
                        {
                            try
                            {
                                int row = cell.RowIndex;
                                if (!byRow.TryGetValue(row, out List<Word.Cell> list))
                                {
                                    list = new List<Word.Cell>();
                                    byRow[row] = list;
                                }

                                list.Add(cell);
                            }
                            catch (System.Runtime.InteropServices.COMException)
                            {
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"[TableLogicalCellMap] Rows 遍历异常: {ex.Message}");
                }
            }

            foreach (List<Word.Cell> list in byRow.Values)
            {
                list.Sort((a, b) => a.ColumnIndex.CompareTo(b.ColumnIndex));
            }

            return byRow;
        }
    }
}
