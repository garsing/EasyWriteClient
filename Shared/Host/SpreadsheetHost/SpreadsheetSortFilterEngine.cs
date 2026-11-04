using System;
using System.Collections.Generic;

namespace WordAddIn1.SpreadsheetHost
{
    internal static class SpreadsheetSortFilterEngine
    {
        public static bool Run(
            ISortFilterSheet sheet,
            SpreadsheetSortFilterRequest request,
            out SpreadsheetSortFilterResult result,
            out string error)
        {
            result = new SpreadsheetSortFilterResult
            {
                Action = request.Action,
                Sheet = request.Sheet,
                Panels = new List<SortFilterPanel>()
            };
            error = null;
            try
            {
                if (request.Action == "list")
                {
                    return List(sheet, result, out error);
                }

                if (request.Action == "clear")
                {
                    return Clear(sheet, request, result, out error);
                }

                return Write(sheet, request, result, out error);
            }
            catch (Exception ex)
            {
                while (ex.InnerException != null
                    && ex.GetType().Name.IndexOf("TargetInvocation", StringComparison.Ordinal) >= 0)
                {
                    ex = ex.InnerException;
                }

                error = "排序和筛选失败: " + ex.Message;
                return false;
            }
        }

        private static bool List(
            ISortFilterSheet sheet,
            SpreadsheetSortFilterResult result,
            out string error)
        {
            if (!sheet.TryList(out List<SortFilterLivePanel> live, out error))
            {
                return false;
            }

            for (int i = 0; i < live.Count; i++)
            {
                if (!TryDescribe(sheet, live[i], i + 1, out SortFilterPanel panel, out error))
                {
                    return false;
                }

                result.Panels.Add(panel);
            }

            return true;
        }

        private static bool Clear(
            ISortFilterSheet sheet,
            SpreadsheetSortFilterRequest request,
            SpreadsheetSortFilterResult result,
            out string error)
        {
            if (!sheet.TryList(out List<SortFilterLivePanel> live, out error))
            {
                return false;
            }

            if (request.Index < 1 || request.Index > live.Count)
            {
                error = "序号不存在";
                return false;
            }

            SortFilterLivePanel target = live[request.Index - 1];
            result.ClearedIndex = request.Index;
            result.ClearedRange = A1Address.Range(
                target.HeaderRow,
                target.FirstCol,
                target.HeaderRow,
                target.LastCol);
            ClearOne(sheet, target);
            return true;
        }

        private static bool Write(
            ISortFilterSheet sheet,
            SpreadsheetSortFilterRequest request,
            SpreadsheetSortFilterResult result,
            out string error)
        {
            error = null;
            if (!sheet.TryList(out List<SortFilterLivePanel> live, out error))
            {
                return false;
            }

            SortFilterLivePanel indexed = null;
            if (request.Action == "update")
            {
                if (request.Index < 1 || request.Index > live.Count)
                {
                    error = "序号不存在";
                    return false;
                }

                indexed = live[request.Index - 1];
            }

            if (!sheet.TryTableBounds(
                    request.HeaderRow,
                    request.FirstCol,
                    request.LastCol,
                    out bool inside,
                    out bool fullHeader,
                    out int tableLast,
                    out string tableName,
                    out error))
            {
                return false;
            }

            if (inside && !fullHeader)
            {
                error = "插入表的筛选区域必须是该表表头整行";
                return false;
            }

            if (!TryHeaderMap(
                    sheet,
                    request.HeaderRow,
                    request.FirstCol,
                    request.LastCol,
                    out Dictionary<string, int> columns,
                    out error))
            {
                return false;
            }

            if (!ResolveColumns(request, columns, out error))
            {
                return false;
            }

            int dataLast = inside
                ? tableLast
                : ExpandDown(sheet, request.HeaderRow, request.FirstCol, request.LastCol);
            if (dataLast <= request.HeaderRow)
            {
                error = "这一行下面没有数据";
                return false;
            }

            int rows = dataLast - request.HeaderRow + 1;
            int cols = request.LastCol - request.FirstCol + 1;
            if (!SpreadsheetBulkGridLimits.TryCheck(rows, cols, "筛选", out error))
            {
                return false;
            }

            if (indexed != null)
            {
                ClearOne(sheet, indexed);
            }

            if (!sheet.TryList(out live, out error))
            {
                return false;
            }

            SortFilterLivePanel occupying = FindCovering(live, request.HeaderRow, request.FirstCol, request.LastCol, inside, tableName);
            if (occupying != null)
            {
                ClearOne(sheet, occupying);
            }

            if (request.Sorts != null && request.Sorts.Count > 0)
            {
                sheet.Sort(
                    request.HeaderRow,
                    request.FirstCol,
                    dataLast,
                    request.LastCol,
                    request.Sorts,
                    columns);
            }

            sheet.ApplyFilters(
                request.HeaderRow,
                request.FirstCol,
                dataLast,
                request.LastCol,
                request.Filters,
                columns,
                inside,
                tableName);

            if (!sheet.TryList(out live, out error))
            {
                return false;
            }

            int found = -1;
            for (int i = 0; i < live.Count; i++)
            {
                SortFilterLivePanel panel = live[i];
                if (panel.HeaderRow == request.HeaderRow
                    && panel.FirstCol == request.FirstCol
                    && panel.LastCol == request.LastCol)
                {
                    found = i;
                    break;
                }
            }

            if (found < 0)
            {
                error = "写入后面板没有出现在 list 中";
                return false;
            }

            if (!TryDescribe(sheet, live[found], found + 1, out SortFilterPanel written, out error))
            {
                return false;
            }

            result.Panel = written;
            return true;
        }

        private static bool TryDescribe(
            ISortFilterSheet sheet,
            SortFilterLivePanel live,
            int index,
            out SortFilterPanel panel,
            out string error)
        {
            panel = null;
            int dataLast = live.DataLastRow;
            if (!live.IsTable || dataLast <= 0)
            {
                dataLast = ExpandDown(sheet, live.HeaderRow, live.FirstCol, live.LastCol);
            }

            if (!sheet.TryReadFilters(live, out List<SortFilterColumn> filters, out error))
            {
                return false;
            }

            if (!sheet.TryReadSorts(live, out List<SortFilterSortKey> sorts, out error))
            {
                return false;
            }

            int dataRows = Math.Max(0, dataLast - live.HeaderRow);
            panel = new SortFilterPanel
            {
                Index = index,
                Range = A1Address.Range(live.HeaderRow, live.FirstCol, live.HeaderRow, live.LastCol),
                Sorts = sorts ?? new List<SortFilterSortKey>(),
                Filters = filters ?? new List<SortFilterColumn>(),
                DataRows = dataRows,
                VisibleRows = sheet.CountVisible(live.HeaderRow, live.FirstCol, dataLast)
            };
            return true;
        }

        private static void ClearOne(ISortFilterSheet sheet, SortFilterLivePanel panel)
        {
            if (panel.IsTable)
            {
                sheet.ClearTable(panel.TableName);
                return;
            }

            sheet.ClearWorksheet();
        }

        private static SortFilterLivePanel FindCovering(
            List<SortFilterLivePanel> live,
            int headerRow,
            int firstCol,
            int lastCol,
            bool inside,
            string tableName)
        {
            foreach (SortFilterLivePanel panel in live)
            {
                if (inside && panel.IsTable
                    && string.Equals(panel.TableName, tableName, StringComparison.OrdinalIgnoreCase))
                {
                    return panel;
                }

                if (!inside && !panel.IsTable)
                {
                    return panel;
                }

                if (panel.HeaderRow == headerRow
                    && panel.FirstCol == firstCol
                    && panel.LastCol == lastCol)
                {
                    return panel;
                }
            }

            return null;
        }

        private static bool TryHeaderMap(
            ISortFilterSheet sheet,
            int headerRow,
            int firstCol,
            int lastCol,
            out Dictionary<string, int> columns,
            out string error)
        {
            columns = new Dictionary<string, int>(StringComparer.Ordinal);
            error = null;
            for (int col = firstCol; col <= lastCol; col++)
            {
                string name = (sheet.ReadCell(headerRow, col) ?? "").Trim();
                if (string.IsNullOrEmpty(name))
                {
                    error = "表头有空列名";
                    return false;
                }

                if (columns.ContainsKey(name))
                {
                    error = "列名重复: " + name;
                    return false;
                }

                columns[name] = col;
            }

            return true;
        }

        private static bool ResolveColumns(
            SpreadsheetSortFilterRequest request,
            Dictionary<string, int> columns,
            out string error)
        {
            error = null;
            if (request.Sorts != null)
            {
                foreach (SortFilterSortKey sort in request.Sorts)
                {
                    if (!columns.ContainsKey(sort.Column))
                    {
                        error = "列名不在这一行: " + sort.Column;
                        return false;
                    }
                }
            }

            if (request.Filters != null)
            {
                foreach (SortFilterColumn filter in request.Filters)
                {
                    if (!columns.ContainsKey(filter.Column))
                    {
                        error = "列名不在这一行: " + filter.Column;
                        return false;
                    }
                }
            }

            return true;
        }

        private static int ExpandDown(ISortFilterSheet sheet, int headerRow, int firstCol, int lastCol)
        {
            int row = headerRow + 1;
            int last = headerRow;
            int guard = headerRow + SpreadsheetBulkGridLimits.MaxRows + 2;
            while (row < guard)
            {
                bool any = false;
                for (int col = firstCol; col <= lastCol; col++)
                {
                    if (!sheet.CellEmpty(row, col))
                    {
                        any = true;
                        break;
                    }
                }

                if (!any)
                {
                    break;
                }

                last = row;
                row++;
            }

            return last;
        }
    }
}
