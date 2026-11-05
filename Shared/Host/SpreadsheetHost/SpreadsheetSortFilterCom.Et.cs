using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using WordAddIn1.OpenFiles;

namespace WordAddIn1.SpreadsheetHost
{
    internal static partial class SpreadsheetSortFilterCom
    {
        public static bool TryExecuteOnEtWorkbook(
            object book,
            SpreadsheetSortFilterRequest request,
            out SpreadsheetSortFilterResult result,
            out string error)
        {
            result = null;
            error = null;
            object sheet = null;
            foreach (object candidate in EtCom.EnumerateSheets(book))
            {
                string name = Convert.ToString(EtCom.GetProperty(candidate, "Name")) ?? "";
                if (string.Equals(name, request.Sheet, StringComparison.OrdinalIgnoreCase))
                {
                    sheet = candidate;
                    break;
                }
            }

            if (sheet == null)
            {
                error = "未找到工作表: " + request.Sheet;
                return false;
            }

            var surface = new EtSortFilterSheet(sheet);
            return SpreadsheetSortFilterEngine.Run(surface, request, out result, out error);
        }

        private sealed class EtSortFilterSheet : ISortFilterSheet
        {
            private readonly object _sheet;

            public EtSortFilterSheet(object sheet)
            {
                _sheet = sheet;
            }

            public bool TryList(out List<SortFilterLivePanel> panels, out string error)
            {
                panels = new List<SortFilterLivePanel>();
                error = null;
                var seen = new List<SortFilterLivePanel>();
                try
                {
                    object tables = EtCom.GetProperty(_sheet, "ListObjects");
                    int count = tables == null ? 0 : Convert.ToInt32(EtCom.GetProperty(tables, "Count"), CultureInfo.InvariantCulture);
                    for (int i = 1; i <= count; i++)
                    {
                        object table = EtCom.GetIndexed(tables, i);
                        if (table == null || !ReadBool(table, "ShowAutoFilter"))
                        {
                            continue;
                        }

                        object header = EtCom.GetProperty(table, "HeaderRowRange");
                        if (header == null)
                        {
                            continue;
                        }

                        seen.Add(new SortFilterLivePanel
                        {
                            HeaderRow = ReadInt(header, "Row"),
                            FirstCol = ReadInt(header, "Column"),
                            LastCol = ReadInt(header, "Column") + ReadInt(EtCom.GetProperty(header, "Columns"), "Count") - 1,
                            IsTable = true,
                            TableName = Convert.ToString(EtCom.GetProperty(table, "Name")),
                            DataLastRow = TableDataLast(table)
                        });
                    }
                }
                catch (Exception ex)
                {
                    error = "读取插入表筛选失败: " + ex.Message;
                    return false;
                }

                try
                {
                    if (ReadBool(_sheet, "AutoFilterMode"))
                    {
                        object auto = EtCom.GetProperty(_sheet, "AutoFilter");
                        object range = auto == null ? null : EtCom.GetProperty(auto, "Range");
                        object rows = range == null ? null : EtCom.GetProperty(range, "Rows");
                        object header = rows == null ? null : EtCom.GetIndexed(rows, 1);
                        if (header != null)
                        {
                            var panel = new SortFilterLivePanel
                            {
                                HeaderRow = ReadInt(header, "Row"),
                                FirstCol = ReadInt(header, "Column"),
                                LastCol = ReadInt(header, "Column") + ReadInt(EtCom.GetProperty(header, "Columns"), "Count") - 1,
                                IsTable = false
                            };
                            if (!Covered(seen, panel))
                            {
                                seen.Add(panel);
                            }
                        }
                    }
                }
                catch (Exception)
                {
                }

                seen.Sort((a, b) =>
                {
                    int row = a.HeaderRow.CompareTo(b.HeaderRow);
                    return row != 0 ? row : a.FirstCol.CompareTo(b.FirstCol);
                });
                panels = seen;
                return true;
            }

            public string ReadCell(int row, int col)
            {
                object cells = EtCom.GetProperty(_sheet, "Cells");
                object cell = EtCom.GetIndexed2(cells, row, col);
                object value = cell == null ? null : EtCom.GetProperty(cell, "Value2");
                return value == null ? "" : (Convert.ToString(value, CultureInfo.InvariantCulture) ?? "");
            }

            public bool CellEmpty(int row, int col)
            {
                return string.IsNullOrEmpty(ReadCell(row, col));
            }

            public bool TryTableBounds(
                int headerRow,
                int firstCol,
                int lastCol,
                out bool inside,
                out bool fullHeader,
                out int dataLastRow,
                out string tableName,
                out string error)
            {
                inside = false;
                fullHeader = false;
                dataLastRow = headerRow;
                tableName = null;
                error = null;
                object tables = EtCom.GetProperty(_sheet, "ListObjects");
                int count = tables == null ? 0 : Convert.ToInt32(EtCom.GetProperty(tables, "Count"), CultureInfo.InvariantCulture);
                for (int i = 1; i <= count; i++)
                {
                    object table = EtCom.GetIndexed(tables, i);
                    object header = table == null ? null : EtCom.GetProperty(table, "HeaderRowRange");
                    if (header == null)
                    {
                        continue;
                    }

                    int row = ReadInt(header, "Row");
                    int c1 = ReadInt(header, "Column");
                    int c2 = c1 + ReadInt(EtCom.GetProperty(header, "Columns"), "Count") - 1;
                    if (headerRow != row || firstCol > c2 || lastCol < c1)
                    {
                        continue;
                    }

                    inside = true;
                    tableName = Convert.ToString(EtCom.GetProperty(table, "Name"));
                    fullHeader = firstCol == c1 && lastCol == c2;
                    dataLastRow = TableDataLast(table);
                    if (!fullHeader)
                    {
                        error = "插入表的筛选区域必须是该表表头整行";
                        return false;
                    }

                    return true;
                }

                return true;
            }

            public void ClearTable(string name)
            {
                object table = FindTable(name);
                if (table != null)
                {
                    EtCom.TrySetProperty(table, "ShowAutoFilter", false);
                }
            }

            public void ClearWorksheet()
            {
                try
                {
                    object auto = EtCom.GetProperty(_sheet, "AutoFilter");
                    object range = auto == null ? null : EtCom.GetProperty(auto, "Range");
                    if (range != null)
                    {
                        EtCom.Invoke(range, "AutoFilter");
                    }
                }
                catch (Exception)
                {
                }

                try
                {
                    if (ReadBool(_sheet, "AutoFilterMode"))
                    {
                        EtCom.TrySetProperty(_sheet, "AutoFilterMode", false);
                    }
                }
                catch (Exception)
                {
                }
            }

            public void Sort(
                int headerRow,
                int firstCol,
                int lastRow,
                int lastCol,
                IList<SortFilterSortKey> sorts,
                IDictionary<string, int> columns)
            {
                object sort = EtCom.GetProperty(_sheet, "Sort");
                if (sort == null)
                {
                    throw new InvalidOperationException("打开排序失败");
                }

                object fields = EtCom.GetProperty(sort, "SortFields");
                EtCom.Invoke(fields, "Clear");
                foreach (SortFilterSortKey key in sorts)
                {
                    int col = columns[key.Column];
                    object keyRange = RangeOf(headerRow, col, lastRow, col);
                    EtCom.Invoke(
                        fields,
                        "Add",
                        keyRange,
                        0,
                        key.Order == "desc" ? 2 : 1);
                }

                EtCom.Invoke(sort, "SetRange", RangeOf(headerRow, firstCol, lastRow, lastCol));
                EtCom.TrySetProperty(sort, "Header", 1);
                EtCom.TrySetProperty(sort, "MatchCase", false);
                EtCom.TrySetProperty(sort, "Orientation", 1);
                EtCom.Invoke(sort, "Apply");
            }

            public void ApplyFilters(
                int headerRow,
                int firstCol,
                int lastRow,
                int lastCol,
                IList<SortFilterColumn> filters,
                IDictionary<string, int> columns,
                bool table,
                string tableName)
            {
                object range = table
                    ? EtCom.GetProperty(FindTable(tableName), "Range")
                    : RangeOf(headerRow, firstCol, lastRow, lastCol);
                if (table)
                {
                    EtCom.TrySetProperty(FindTable(tableName), "ShowAutoFilter", true);
                }
                else
                {
                    EtCom.InvokeFlex(range, "AutoFilter");
                }

                foreach (SortFilterColumn filter in filters)
                {
                    int field = columns[filter.Column] - firstCol + 1;
                    if (filter.IsIn)
                    {
                        EtCom.InvokeFlex(range, "AutoFilter", field, CheckedValues(filter.InValues), 7, Type.Missing, Type.Missing);
                    }
                    else if (string.IsNullOrEmpty(filter.Criteria2))
                    {
                        EtCom.InvokeFlex(range, "AutoFilter", field, filter.Criteria1, 1, Type.Missing, true);
                    }
                    else
                    {
                        EtCom.InvokeFlex(
                            range,
                            "AutoFilter",
                            field,
                            filter.Criteria1,
                            filter.OrJoin ? 2 : 1,
                            filter.Criteria2,
                            Type.Missing);
                    }
                }
            }

            public bool TryReadFilters(SortFilterLivePanel panel, out List<SortFilterColumn> filters, out string error)
            {
                filters = new List<SortFilterColumn>();
                error = null;
                object auto = panel.IsTable
                    ? EtCom.GetProperty(FindTable(panel.TableName), "AutoFilter")
                    : EtCom.GetProperty(_sheet, "AutoFilter");
                object collection = auto == null ? null : EtCom.GetProperty(auto, "Filters");
                if (collection == null)
                {
                    return true;
                }

                int count = Convert.ToInt32(EtCom.GetProperty(collection, "Count"), CultureInfo.InvariantCulture);
                for (int i = 1; i <= count; i++)
                {
                    object filter = EtCom.GetIndexed(collection, i);
                    if (filter == null || !ReadBool(filter, "On"))
                    {
                        continue;
                    }

                    string column = (ReadCell(panel.HeaderRow, panel.FirstCol + i - 1) ?? "").Trim();
                    int op = 0;
                    try
                    {
                        op = Convert.ToInt32(EtCom.GetProperty(filter, "Operator"), CultureInfo.InvariantCulture);
                    }
                    catch (Exception)
                    {
                        op = 0;
                    }

                    if (op == 3 || op == 4 || op == 5 || op == 6 || op >= 8)
                    {
                        error = "筛选列「" + column + "」不是勾选或比较式，本工具不能读回";
                        return false;
                    }

                    object criteria1 = EtCom.InvokeFlex(filter, "Criteria1");
                    if (op == 7)
                    {
                        object second = null;
                        try
                        {
                            second = EtCom.InvokeFlex(filter, "Criteria2");
                        }
                        catch (Exception)
                        {
                            second = null;
                        }

                        filters.Add(new SortFilterColumn
                        {
                            Column = column,
                            IsIn = true,
                            InValues = SpreadsheetSortFilterParse.ReadCheckedValues(criteria1, second)
                        });
                        continue;
                    }

                    string c1 = Convert.ToString(criteria1, CultureInfo.InvariantCulture) ?? "";
                    string c2 = null;
                    if (op == 1 || op == 2)
                    {
                        try
                        {
                            c2 = Convert.ToString(EtCom.GetProperty(filter, "Criteria2"), CultureInfo.InvariantCulture);
                        }
                        catch (Exception)
                        {
                            c2 = null;
                        }
                    }

                    if (HasWildcard(c1) || HasWildcard(c2))
                    {
                        error = "筛选列「" + column + "」不是勾选或比较式，本工具不能读回";
                        return false;
                    }

                    if (string.IsNullOrEmpty(c2)
                        && c1.StartsWith("=", StringComparison.Ordinal)
                        && !c1.StartsWith(">=", StringComparison.Ordinal))
                    {
                        string value = c1.Substring(1);
                        filters.Add(new SortFilterColumn
                        {
                            Column = column,
                            IsIn = true,
                            InValues = new List<string> { value }
                        });
                        continue;
                    }

                    if (op == 2
                        && SpreadsheetSortFilterParse.TryTwoChecked(c1, c2, out List<string> checkedValues))
                    {
                        filters.Add(new SortFilterColumn
                        {
                            Column = column,
                            IsIn = true,
                            InValues = checkedValues
                        });
                        continue;
                    }

                    filters.Add(new SortFilterColumn
                    {
                        Column = column,
                        IsIn = false,
                        Expr = SpreadsheetSortFilterParse.CanonicalExpr(c1, c2, op == 2)
                    });
                }

                return true;
            }

            public bool TryReadSorts(SortFilterLivePanel panel, out List<SortFilterSortKey> sorts, out string error)
            {
                sorts = new List<SortFilterSortKey>();
                error = null;
                object sort = null;
                if (panel.IsTable)
                {
                    object table = FindTable(panel.TableName);
                    sort = table == null ? null : EtCom.GetProperty(table, "Sort");
                }

                if (!HasSortFields(sort))
                {
                    sort = EtCom.GetProperty(_sheet, "Sort");
                }

                if (!HasSortFields(sort) || !SortRangeCoversPanel(sort, panel))
                {
                    return true;
                }

                object fields = EtCom.GetProperty(sort, "SortFields");
                int count = Convert.ToInt32(EtCom.GetProperty(fields, "Count"), CultureInfo.InvariantCulture);
                for (int i = 1; i <= count; i++)
                {
                    object field = EtCom.GetIndexed(fields, i);
                    object key = field == null ? null : EtCom.GetProperty(field, "Key");
                    if (key == null)
                    {
                        continue;
                    }

                    int col = ReadInt(key, "Column");
                    if (col < panel.FirstCol || col > panel.LastCol)
                    {
                        continue;
                    }

                    int order = Convert.ToInt32(EtCom.GetProperty(field, "Order"), CultureInfo.InvariantCulture);
                    sorts.Add(new SortFilterSortKey
                    {
                        Column = (ReadCell(panel.HeaderRow, col) ?? "").Trim(),
                        Order = order == 2 ? "desc" : "asc"
                    });
                }

                return true;
            }

            private static bool HasSortFields(object sort)
            {
                if (sort == null)
                {
                    return false;
                }

                try
                {
                    object fields = EtCom.GetProperty(sort, "SortFields");
                    if (fields == null)
                    {
                        return false;
                    }

                    return Convert.ToInt32(EtCom.GetProperty(fields, "Count"), CultureInfo.InvariantCulture) > 0;
                }
                catch (Exception)
                {
                    return false;
                }
            }

            private static bool SortRangeCoversPanel(object sort, SortFilterLivePanel panel)
            {
                object range = null;
                try
                {
                    range = EtCom.GetProperty(sort, "Rng");
                }
                catch (Exception)
                {
                    return false;
                }

                if (range == null)
                {
                    return false;
                }

                int row = ReadInt(range, "Row");
                int first = ReadInt(range, "Column");
                int last = first + ReadInt(EtCom.GetProperty(range, "Columns"), "Count") - 1;
                return row == panel.HeaderRow && first <= panel.FirstCol && last >= panel.LastCol;
            }

            public int CountVisible(int headerRow, int firstCol, int lastRow)
            {
                if (lastRow <= headerRow)
                {
                    return 0;
                }

                try
                {
                    object body = RangeOf(headerRow + 1, firstCol, lastRow, firstCol);
                    object visible = EtCom.Invoke(body, "SpecialCells", 12);
                    if (visible == null)
                    {
                        return 0;
                    }

                    return Convert.ToInt32(EtCom.GetProperty(visible, "Count"), CultureInfo.InvariantCulture);
                }
                catch (Exception)
                {
                    return 0;
                }
            }

            private object FindTable(string name)
            {
                object tables = EtCom.GetProperty(_sheet, "ListObjects");
                int count = tables == null ? 0 : Convert.ToInt32(EtCom.GetProperty(tables, "Count"), CultureInfo.InvariantCulture);
                for (int i = 1; i <= count; i++)
                {
                    object table = EtCom.GetIndexed(tables, i);
                    string tableName = table == null ? "" : Convert.ToString(EtCom.GetProperty(table, "Name"));
                    if (string.Equals(tableName, name, StringComparison.OrdinalIgnoreCase))
                    {
                        return table;
                    }
                }

                return null;
            }

            private object RangeOf(int r1, int c1, int r2, int c2)
            {
                return EtCom.InvokeFlex(_sheet, "Range", A1Address.Range(r1, c1, r2, c2));
            }

            private static int TableDataLast(object table)
            {
                object range = EtCom.GetProperty(table, "Range");
                int last = ReadInt(range, "Row") + ReadInt(EtCom.GetProperty(range, "Rows"), "Count") - 1;
                try
                {
                    if (ReadBool(table, "ShowTotals"))
                    {
                        object totals = EtCom.GetProperty(table, "TotalsRowRange");
                        if (totals != null)
                        {
                            last = ReadInt(totals, "Row") - 1;
                        }
                    }
                }
                catch (Exception)
                {
                }

                return last;
            }

            private static bool Covered(List<SortFilterLivePanel> panels, SortFilterLivePanel candidate)
            {
                foreach (SortFilterLivePanel panel in panels)
                {
                    if (panel.HeaderRow == candidate.HeaderRow
                        && panel.FirstCol == candidate.FirstCol
                        && panel.LastCol == candidate.LastCol)
                    {
                        return true;
                    }
                }

                return false;
            }

            private static int ReadInt(object target, string name)
            {
                return Convert.ToInt32(EtCom.GetProperty(target, name), CultureInfo.InvariantCulture);
            }

            private static bool ReadBool(object target, string name)
            {
                object value = EtCom.GetProperty(target, name);
                if (value == null)
                {
                    return false;
                }

                if (value is bool flag)
                {
                    return flag;
                }

                return Convert.ToInt32(value, CultureInfo.InvariantCulture) != 0;
            }

            private static object[] CheckedValues(IList<string> values)
            {
                var boxed = new object[values.Count];
                for (int i = 0; i < values.Count; i++)
                {
                    boxed[i] = values[i];
                }

                return boxed;
            }

            private static bool HasWildcard(string criteria)
            {
                return !string.IsNullOrEmpty(criteria)
                    && (criteria.IndexOf('*') >= 0 || criteria.IndexOf('?') >= 0);
            }
        }
    }
}
