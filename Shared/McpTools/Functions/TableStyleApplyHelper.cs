using System;
using System.Collections.Generic;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    /// <summary>
    /// 表格 Style apply（三线表等），自 F_CreateTableFromXmlTool 抽出。
    /// </summary>
    public static class TableStyleApplyHelper
    {
        public class ApplyResult
        {
            public bool Success { get; set; }
            public string Error { get; set; }
        }

        public static ApplyResult Apply(Word.Table table, TableStyleConfig styleConfig)
        {
            if (styleConfig == null || string.IsNullOrEmpty(styleConfig.TableStyle))
            {
                return new ApplyResult { Success = true, Error = null };
            }

            return ApplyViaCells(table, styleConfig);
        }

        /// <summary>
        /// 通过 Range.Cells 应用样式，兼容纵向合并表（table.Rows.Count 可能抛 COMException）。
        /// </summary>
        public static ApplyResult ApplyViaCells(Word.Table table, TableStyleConfig styleConfig)
        {
            if (styleConfig == null || string.IsNullOrEmpty(styleConfig.TableStyle))
            {
                return new ApplyResult { Success = true, Error = null };
            }

            return ApplyViaCells(table, styleConfig.TableStyle, styleConfig.HeaderRows);
        }

        public static ApplyResult ApplyViaCells(Word.Table table, string tableStyle, int headerRows = 1)
        {
            if (string.IsNullOrEmpty(tableStyle))
            {
                return new ApplyResult { Success = true, Error = null };
            }

            if (!string.Equals(tableStyle, "ThreeLine", StringComparison.OrdinalIgnoreCase))
            {
                System.Diagnostics.Debug.WriteLine($"[TableStyleApplyHelper] 未知表格样式: {tableStyle}，跳过");
                return new ApplyResult { Success = true, Error = null };
            }

            try
            {
                var rowCells = new Dictionary<int, List<Word.Cell>>();
                int maxRow = 0;
                foreach (Word.Cell cell in table.Range.Cells)
                {
                    int rowIndex = cell.RowIndex;
                    maxRow = Math.Max(maxRow, rowIndex);
                    if (!rowCells.TryGetValue(rowIndex, out List<Word.Cell> cells))
                    {
                        cells = new List<Word.Cell>();
                        rowCells[rowIndex] = cells;
                    }

                    cells.Add(cell);
                }

                if (maxRow <= 0)
                {
                    return new ApplyResult { Success = false, Error = "表格无可用单元格" };
                }

                int effectiveHeaderRows = Math.Max(1, Math.Min(headerRows, maxRow));

                ClearTableBordersBulk(table);

                foreach (List<Word.Cell> cells in rowCells.Values)
                {
                    foreach (Word.Cell cell in cells)
                    {
                        ClearCellBorders(cell);
                    }
                }

                SetCellsBorder(
                    rowCells,
                    1,
                    Word.WdBorderType.wdBorderTop,
                    Word.WdLineStyle.wdLineStyleSingle,
                    Word.WdLineWidth.wdLineWidth150pt);
                SetCellsBorder(
                    rowCells,
                    effectiveHeaderRows,
                    Word.WdBorderType.wdBorderBottom,
                    Word.WdLineStyle.wdLineStyleSingle,
                    Word.WdLineWidth.wdLineWidth150pt);
                SetCellsBorder(
                    rowCells,
                    maxRow,
                    Word.WdBorderType.wdBorderBottom,
                    Word.WdLineStyle.wdLineStyleSingle,
                    Word.WdLineWidth.wdLineWidth150pt);

                System.Diagnostics.Debug.WriteLine(
                    $"[TableStyleApplyHelper] 三线表(Cells)应用成功，表头行数={effectiveHeaderRows}，末行={maxRow}");
                return new ApplyResult { Success = true, Error = null };
            }
            catch (Exception ex)
            {
                return new ApplyResult { Success = false, Error = $"表格样式应用异常: {ex.Message}" };
            }
        }

        private static void ClearTableBordersBulk(Word.Table table)
        {
            if (table == null)
            {
                return;
            }

            Word.WdBorderType[] types =
            {
                Word.WdBorderType.wdBorderTop,
                Word.WdBorderType.wdBorderBottom,
                Word.WdBorderType.wdBorderLeft,
                Word.WdBorderType.wdBorderRight,
                Word.WdBorderType.wdBorderHorizontal,
                Word.WdBorderType.wdBorderVertical,
            };

            foreach (Word.WdBorderType borderType in types)
            {
                try
                {
                    table.Borders[borderType].LineStyle = Word.WdLineStyle.wdLineStyleNone;
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine(
                        $"[TableStyleApplyHelper] 清除表格{borderType}边框失败: {ex.Message}");
                }
            }
        }

        private static void ClearCellBorders(Word.Cell cell)
        {
            try
            {
                cell.Borders[Word.WdBorderType.wdBorderTop].LineStyle = Word.WdLineStyle.wdLineStyleNone;
                cell.Borders[Word.WdBorderType.wdBorderBottom].LineStyle = Word.WdLineStyle.wdLineStyleNone;
                cell.Borders[Word.WdBorderType.wdBorderLeft].LineStyle = Word.WdLineStyle.wdLineStyleNone;
                cell.Borders[Word.WdBorderType.wdBorderRight].LineStyle = Word.WdLineStyle.wdLineStyleNone;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[TableStyleApplyHelper] 清除单元格边框失败: {ex.Message}");
            }
        }

        private static void SetCellsBorder(
            Dictionary<int, List<Word.Cell>> rowCells,
            int rowIndex,
            Word.WdBorderType borderType,
            Word.WdLineStyle lineStyle,
            Word.WdLineWidth lineWidth)
        {
            if (!rowCells.TryGetValue(rowIndex, out List<Word.Cell> cells))
            {
                return;
            }

            foreach (Word.Cell cell in cells)
            {
                try
                {
                    var border = cell.Borders[borderType];
                    border.LineStyle = lineStyle;
                    border.LineWidth = lineWidth;
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine(
                        $"[TableStyleApplyHelper] 设置行{rowIndex}边框失败: {ex.Message}");
                }
            }
        }

        public static ApplyResult Apply(Word.Table table, string tableStyle, int headerRows = 1)
        {
            return ApplyViaCells(table, tableStyle, headerRows);
        }
    }
}
