using System;
using System.Collections.Generic;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    /// <summary>
    /// T_ → COM Table。书签通道按 ewT_ 取 Range；Title 通道用留下的编号骨架。
    /// </summary>
    public static class TableResolveHelper
    {
        public static Word.Table TryResolveTableById(Word.Document document, string tableId, out string error)
        {
            return TryResolveTableById(document, tableId, refreshDocument: false, out error);
        }

        public static Word.Table TryResolveTableById(
            Word.Document document,
            string tableId,
            bool refreshDocument,
            out string error)
        {
            error = null;
            if (document == null || string.IsNullOrEmpty(tableId))
            {
                error = $"找不到表格编号 '{tableId}'，请确认表格编号是否正确";
                return null;
            }

            if (refreshDocument)
            {
                try
                {
                    WordDocumentExtractor.ProcessDocument(
                        document,
                        ProcessDocumentOptions.ForGetDocumentContent("table_resolve"));
                }
                catch (Exception ex)
                {
                    error = $"找不到表格编号 '{tableId}'，请确认表格编号是否正确";
                    System.Diagnostics.Debug.WriteLine("[TableResolve] ProcessDocument 失败: " + ex.Message);
                    return null;
                }
            }

            string channel = DocumentState.GetTableStampChannel();
            if (channel == TableTitleStampHelper.ChannelBookmark)
            {
                Word.Range mark = TableTitleStampHelper.TryGetBookmarkRange(document, tableId);
                if (mark == null)
                {
                    error = $"找不到表格编号 '{tableId}'，请确认表格编号是否正确";
                    return null;
                }

                Word.Table fromBookmark = TableFromBookmarkRange(mark);
                if (fromBookmark == null)
                {
                    error = $"找不到表格编号 '{tableId}'，请确认表格编号是否正确";
                    return null;
                }

                return fromBookmark;
            }

            string markup = DocumentState.GetNumberedTableMarkup();
            if (string.IsNullOrEmpty(markup))
            {
                markup = TableMarkupSkeleton.FromDocument(document);
            }

            List<Word.Table> allTables = GetAllTablesInOrder(document, markup);
            return TryResolveFromList(allTables, tableId, markup, out error);
        }

        /// <summary>
        /// 在已经收好的扁平表列表里按 <paramref name="tableId"/> 取表。不再枚举 COM。
        /// </summary>
        public static Word.Table TryResolveFromList(List<Word.Table> allTables, string tableId, out string error)
        {
            return TryResolveFromList(allTables, tableId, DocumentState.GetNumberedTableMarkup(), out error);
        }

        public static Word.Table TryResolveFromList(
            List<Word.Table> allTables,
            string tableId,
            string numberedMarkup,
            out string error)
        {
            error = null;
            if (allTables == null || string.IsNullOrEmpty(tableId))
            {
                error = $"找不到表格编号 '{tableId}'，请确认表格编号是否正确";
                return null;
            }

            int tableOrderIndex = IndexOfTableId(numberedMarkup, tableId);
            if (tableOrderIndex < 0)
            {
                tableOrderIndex = DocumentState.GetTableIndex(tableId);
            }

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

        private static int IndexOfTableId(string markup, string tableId)
        {
            if (string.IsNullOrEmpty(markup) || string.IsNullOrEmpty(tableId))
            {
                return -1;
            }

            List<string> ids = NestedTableParentFinder.ListOpenIds(markup);
            return ids.IndexOf(tableId);
        }

        private static Word.Table TableFromBookmarkRange(Word.Range mark)
        {
            if (mark == null)
            {
                return null;
            }

            try
            {
                if (mark.Tables != null && mark.Tables.Count > 0)
                {
                    return mark.Tables[1];
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("[TableResolve] 书签 Range.Tables 失败: " + ex.Message);
            }

            return null;
        }

        public static List<Word.Table> GetAllTablesInOrder(Word.Document document)
        {
            return GetAllTablesInOrder(document, display: null);
        }

        /// <summary>
        /// 顶层用 <c>document.Tables</c>。内表用 display 找到父格，只进那一格的 <c>cell.Tables</c>。
        /// <paramref name="display"/> 空则从 OOXML 抽表骨架。
        /// </summary>
        public static List<Word.Table> GetAllTablesInOrder(Word.Document document, string display)
        {
            var allTables = new List<Word.Table>();
            if (document == null)
            {
                return allTables;
            }

            try
            {
                var tops = new List<Word.Table>();
                foreach (Word.Table top in document.Tables)
                {
                    if (top == null || ContainsSameTable(tops, top))
                    {
                        continue;
                    }

                    tops.Add(top);
                }

                tops.Sort(CompareByStart);
                string markup = display;
                if (string.IsNullOrEmpty(markup))
                {
                    markup = TableMarkupSkeleton.FromDocument(document);
                }

                if (string.IsNullOrEmpty(markup))
                {
                    return tops;
                }

                CollectByDisplayParentCells(markup, tops, allTables);
                allTables.Sort(CompareByStart);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[TableResolve] 收集表格失败: {ex.Message}");
            }

            return allTables;
        }

        /// <summary>
        /// display 出现序：顶层对 tops 的起点序；内表只解析父格，取尚未收入的 <c>cell.Tables</c>。
        /// </summary>
        private static void CollectByDisplayParentCells(
            string display,
            List<Word.Table> tops,
            List<Word.Table> allTables)
        {
            var byId = new Dictionary<string, Word.Table>(StringComparer.Ordinal);
            int topIndex = 0;
            List<string> ids = NestedTableParentFinder.ListOpenIds(display);
            for (int i = 0; i < ids.Count; i++)
            {
                string id = ids[i];
                NestedTableParentFinder.Result loc = NestedTableParentFinder.Find(display, id);
                if (loc.Kind == NestedTableParentFinder.Kind.TopLevel)
                {
                    if (topIndex >= tops.Count)
                    {
                        System.Diagnostics.Debug.WriteLine(
                            $"[TableResolve] display 顶层表多于 document.Tables，跳过 {id}");
                        continue;
                    }

                    Word.Table top = tops[topIndex];
                    topIndex++;
                    allTables.Add(top);
                    byId[id] = top;
                    continue;
                }

                if (loc.Kind != NestedTableParentFinder.Kind.Nested)
                {
                    System.Diagnostics.Debug.WriteLine(
                        $"[TableResolve] {id} 父格失败: {loc.Error}");
                    continue;
                }

                if (!byId.TryGetValue(loc.ParentTableId, out Word.Table parent) || parent == null)
                {
                    System.Diagnostics.Debug.WriteLine(
                        $"[TableResolve] {id} 父表 {loc.ParentTableId} 尚未收入");
                    continue;
                }

                Word.Table inner = TakeNextInnerFromParentCell(parent, loc.Row, loc.Col, allTables);
                if (inner == null)
                {
                    System.Diagnostics.Debug.WriteLine(
                        $"[TableResolve] {id} 父格 {loc.GeometricCode} 没有新的内表");
                    continue;
                }

                allTables.Add(inner);
                byId[id] = inner;
            }
        }

        private static Word.Table TakeNextInnerFromParentCell(
            Word.Table parent,
            int row,
            int col,
            List<Word.Table> already)
        {
            Word.Cell cell;
            try
            {
                cell = TableRowValuesHelper.TryResolveWritableCell(parent, row, col);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[TableResolve] 父格定位失败: {ex.Message}");
                return null;
            }

            if (cell == null)
            {
                return null;
            }

            var candidates = new List<Word.Table>();
            AddCellTables(cell.Tables, parent, already, candidates);
            if (candidates.Count == 0)
            {
                try
                {
                    AddCellTables(cell.Range.Tables, parent, already, candidates);
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"[TableResolve] 父格 Range.Tables 失败: {ex.Message}");
                }
            }

            if (candidates.Count == 0)
            {
                return null;
            }

            candidates.Sort(CompareByStart);
            return candidates[0];
        }

        private static void AddCellTables(
            Word.Tables source,
            Word.Table parent,
            List<Word.Table> already,
            List<Word.Table> into)
        {
            if (source == null)
            {
                return;
            }

            foreach (Word.Table inner in source)
            {
                try
                {
                    if (inner == null
                        || IsSameTable(parent, inner)
                        || ContainsSameTable(already, inner)
                        || ContainsSameTable(into, inner))
                    {
                        continue;
                    }

                    into.Add(inner);
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"[TableResolve] cell.Tables 失败: {ex.Message}");
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
