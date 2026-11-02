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
            return TryResolveFromList(allTables, tableId, out error);
        }

        /// <summary>
        /// 在已经收好的扁平表列表里按 <paramref name="tableId"/> 取表。不再枚举 COM。
        /// </summary>
        public static Word.Table TryResolveFromList(List<Word.Table> allTables, string tableId, out string error)
        {
            error = null;
            if (allTables == null || string.IsNullOrEmpty(tableId))
            {
                error = $"找不到表格编号 '{tableId}'，请确认表格编号是否正确";
                return null;
            }

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
                foreach (Word.Table top in document.Tables)
                {
                    try
                    {
                        if (top == null || ContainsSameTable(allTables, top))
                        {
                            continue;
                        }

                        allTables.Add(top);
                        ExpandRangeTables(top, allTables);
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"[TableResolve] 收集表格失败: {ex.Message}");
                    }
                }

                allTables.Sort(CompareByStart);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[TableResolve] 收集表格失败: {ex.Message}");
            }

            return allTables;
        }

        /// <summary>
        /// 从种子表的 Range.Tables 收内表和孙表。Range.Tables 含表自己，跳过。
        /// 某宿主只给直接子表时，对刚收进的表再展开，直到没有新表。
        /// </summary>
        private static void ExpandRangeTables(Word.Table seed, List<Word.Table> allTables)
        {
            var pending = new Queue<Word.Table>();
            pending.Enqueue(seed);
            while (pending.Count > 0)
            {
                Word.Table current = pending.Dequeue();
                Word.Tables inners;
                try
                {
                    inners = current.Range.Tables;
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"[TableResolve] Range.Tables 收集失败: {ex.Message}");
                    continue;
                }

                if (inners == null)
                {
                    continue;
                }

                foreach (Word.Table inner in inners)
                {
                    try
                    {
                        if (inner == null || IsSameTable(current, inner) || ContainsSameTable(allTables, inner))
                        {
                            continue;
                        }

                        allTables.Add(inner);
                        pending.Enqueue(inner);
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"[TableResolve] Range.Tables 收集失败: {ex.Message}");
                    }
                }
            }
        }

        private static int CompareByStart(Word.Table left, Word.Table right)
        {
            try
            {
                return left.Range.Start.CompareTo(right.Range.Start);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[TableResolve] 排序失败: {ex.Message}");
                return 0;
            }
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
