using System;
using System.Collections.Generic;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    /// <summary>
    /// T_ → COM Table：只查 TableIdOrder + 文档序第 i 张。不按 Title 查找，不包 ReadWord。
    /// </summary>
    public static class TableResolveHelper
    {
        public static Word.Table TryResolveTableById(Word.Document document, string tableId, out string error)
        {
            error = null;
            if (document == null || string.IsNullOrEmpty(tableId))
            {
                error = $"找不到表格编号 '{tableId}'，请确认表格编号是否正确";
                return null;
            }

            int tableOrderIndex = DocumentState.GetTableIndex(tableId);
            if (tableOrderIndex < 0)
            {
                error = $"找不到表格编号 '{tableId}'，请确认表格编号是否正确";
                return null;
            }

            List<Word.Table> allTables = GetAllTablesInOrder(document);
            if (tableOrderIndex >= allTables.Count)
            {
                error =
                    $"表格编号 '{tableId}' 对应的索引{tableOrderIndex}无效，文档中只有{allTables.Count}个表格（包括嵌套表格）";
                return null;
            }

            return allTables[tableOrderIndex];
        }

        public static List<Word.Table> GetAllTablesInOrder(Word.Document document)
        {
            var allTables = new List<Word.Table>();
            if (document == null)
            {
                return allTables;
            }

            try
            {
                foreach (Word.Table table in document.Tables)
                {
                    CollectTablesRecursive(table, allTables);
                }

                allTables.Sort((t1, t2) => t1.Range.Start.CompareTo(t2.Range.Start));
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[TableResolve] 收集表格失败: {ex.Message}");
            }

            return allTables;
        }

        private static void CollectTablesRecursive(Word.Table table, List<Word.Table> allTables)
        {
            if (table == null)
            {
                return;
            }

            try
            {
                allTables.Add(table);
                foreach (Word.Cell cell in table.Range.Cells)
                {
                    foreach (Word.Table nestedTable in cell.Tables)
                    {
                        CollectTablesRecursive(nestedTable, allTables);
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[TableResolve] 递归收集表格失败: {ex.Message}");
            }
        }
    }
}
