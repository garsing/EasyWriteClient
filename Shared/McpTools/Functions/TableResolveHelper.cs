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

            List<Word.Table> allTables = GetAllTablesInOrder(document);
            Word.Table stamped = FindTableByTitleStamp(allTables, tableId);
            if (stamped != null)
            {
                return stamped;
            }

            int tableOrderIndex = DocumentState.GetTableIndex(tableId);
            if (tableOrderIndex < 0)
            {
                error = $"找不到表格编号 '{tableId}'，请确认表格编号是否正确";
                return null;
            }

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

        private static Word.Table FindTableByTitleStamp(List<Word.Table> tables, string tableId)
        {
            if (tables == null || string.IsNullOrEmpty(tableId))
            {
                return null;
            }

            for (int i = 0; i < tables.Count; i++)
            {
                if (!TableTitleStampHelper.TryReadTitle(tables[i], out string title)
                    || !TableTitleStampHelper.TryParse(title, out string stampedId, out _))
                {
                    continue;
                }

                if (string.Equals(stampedId, tableId, StringComparison.Ordinal))
                {
                    return tables[i];
                }
            }

            return null;
        }

        private static void CollectTablesRecursive(Word.Table table, List<Word.Table> allTables)
        {
            if (table == null || ContainsSameTable(allTables, table))
            {
                return;
            }

            allTables.Add(table);

            foreach (Word.Table nested in ListNestedTables(table))
            {
                CollectTablesRecursive(nested, allTables);
            }
        }

        private static List<Word.Table> ListNestedTables(Word.Table table)
        {
            var nested = new List<Word.Table>();
            CollectNestedFromRows(table, nested);
            if (nested.Count == 0)
            {
                CollectNestedFromRangeCells(table, nested);
            }

            if (nested.Count == 0)
            {
                CollectNestedFromRangeTables(table, nested);
            }

            return nested;
        }

        private static void CollectNestedFromRows(Word.Table table, List<Word.Table> nested)
        {
            try
            {
                foreach (Word.Row row in table.Rows)
                {
                    foreach (Word.Cell cell in row.Cells)
                    {
                        AddCellTables(cell, nested);
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[TableResolve] Rows 收集嵌套失败: {ex.Message}");
            }
        }

        private static void CollectNestedFromRangeCells(Word.Table table, List<Word.Table> nested)
        {
            try
            {
                foreach (Word.Cell cell in table.Range.Cells)
                {
                    AddCellTables(cell, nested);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[TableResolve] Range.Cells 收集嵌套失败: {ex.Message}");
            }
        }

        private static void CollectNestedFromRangeTables(Word.Table table, List<Word.Table> nested)
        {
            try
            {
                int parentLevel = 1;
                try
                {
                    parentLevel = table.NestingLevel;
                }
                catch (Exception)
                {
                }

                foreach (Word.Table inner in table.Range.Tables)
                {
                    if (ContainsSameTable(nested, inner) || IsSameTable(table, inner))
                    {
                        continue;
                    }

                    try
                    {
                        if (inner.NestingLevel <= parentLevel)
                        {
                            continue;
                        }
                    }
                    catch (Exception)
                    {
                    }

                    nested.Add(inner);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[TableResolve] Range.Tables 收集嵌套失败: {ex.Message}");
            }
        }

        private static void AddCellTables(Word.Cell cell, List<Word.Table> nested)
        {
            if (cell == null || cell.Tables == null)
            {
                return;
            }

            try
            {
                foreach (Word.Table inner in cell.Tables)
                {
                    if (!ContainsSameTable(nested, inner))
                    {
                        nested.Add(inner);
                    }
                }
            }
            catch (Exception)
            {
            }
        }

        private static bool ContainsSameTable(List<Word.Table> tables, Word.Table table)
        {
            if (tables == null || table == null)
            {
                return false;
            }

            for (int i = 0; i < tables.Count; i++)
            {
                if (IsSameTable(tables[i], table))
                {
                    return true;
                }
            }

            return false;
        }

        private static bool IsSameTable(Word.Table a, Word.Table b)
        {
            if (a == null || b == null)
            {
                return false;
            }

            try
            {
                return a.Range.Start == b.Range.Start && a.Range.End == b.Range.End;
            }
            catch (Exception)
            {
                return false;
            }
        }
    }
}
