using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace WordAddIn1.SpreadsheetHost
{
    internal static partial class SpreadsheetSortFilterCom
    {
        private const int XlAnd = 1;
        private const int XlOr = 2;
        private const int XlFilterValues = 7;
        private const int XlTop10Items = 3;
        private const int XlSortOnValues = 0;
        private const int XlAscending = 1;
        private const int XlDescending = 2;
        private const int XlYes = 1;
        private const int XlSortColumns = 1;
        private const int XlCellTypeVisible = 12;

        public static bool TryExecuteOnExcelWorkbook(
            Excel.Workbook book,
            SpreadsheetSortFilterRequest request,
            out SpreadsheetSortFilterResult result,
            out string error)
        {
            result = null;
            error = null;
            if (!TryFindExcelSheet(book, request.Sheet, out Excel.Worksheet sheet, out error))
            {
                return false;
            }

            var surface = new ExcelSortFilterSheet(sheet);
            return SpreadsheetSortFilterEngine.Run(surface, request, out result, out error);
        }

        private static bool TryFindExcelSheet(
            Excel.Workbook book,
            string sheetName,
            out Excel.Worksheet sheet,
            out string error)
        {
            sheet = null;
            error = null;
            foreach (Excel.Worksheet candidate in book.Worksheets)
            {
                if (string.Equals(candidate.Name, sheetName, StringComparison.OrdinalIgnoreCase))
                {
                    sheet = candidate;
                    return true;
                }
            }

            error = "未找到工作表: " + sheetName;
            return false;
        }

        private sealed class ExcelSortFilterSheet : ISortFilterSheet
        {
            private readonly Excel.Worksheet _sheet;

            public ExcelSortFilterSheet(Excel.Worksheet sheet)
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
                    foreach (Excel.ListObject table in _sheet.ListObjects)
                    {
                        if (!table.ShowAutoFilter)
                        {
                            continue;
                        }

                        Excel.Range header = table.HeaderRowRange;
                        if (header == null)
                        {
                            continue;
                        }

                        seen.Add(PanelFromRange(header, true, table.Name, TableDataLast(table)));
                    }
                }
                catch (Exception ex)
                {
                    error = "读取插入表筛选失败: " + ex.Message;
                    return false;
                }

                try
                {
                    if (_sheet.AutoFilterMode && _sheet.AutoFilter != null && _sheet.AutoFilter.Range != null)
                    {
                        Excel.Range header = (Excel.Range)_sheet.AutoFilter.Range.Rows[1];
                        SortFilterLivePanel worksheetPanel = PanelFromRange(header, false, null, 0);
                        if (!CoveredBy(seen, worksheetPanel))
                        {
                            seen.Add(worksheetPanel);
                        }
                    }
                }
                catch (Exception)
                {
                    // 只有插入表筛选时，工作表 AutoFilter 可能不可读。
                }

                seen.Sort(ComparePanels);
                panels = seen;
                return true;
            }

            public string ReadCell(int row, int col)
            {
                object value = ((Excel.Range)_sheet.Cells[row, col]).Value2;
                if (value == null)
                {
                    return "";
                }

                return Convert.ToString(value, CultureInfo.InvariantCulture) ?? "";
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
                foreach (Excel.ListObject table in _sheet.ListObjects)
                {
                    Excel.Range header = table.HeaderRowRange;
                    if (header == null)
                    {
                        continue;
                    }

                    int row = header.Row;
                    int c1 = header.Column;
                    int c2 = c1 + header.Columns.Count - 1;
                    bool overlaps = headerRow == row && firstCol <= c2 && lastCol >= c1;
                    if (!overlaps)
                    {
                        continue;
                    }

                    inside = true;
                    tableName = table.Name;
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

            public void ClearTable(string tableName)
            {
                foreach (Excel.ListObject table in _sheet.ListObjects)
                {
                    if (string.Equals(table.Name, tableName, StringComparison.OrdinalIgnoreCase))
                    {
                        table.ShowAutoFilter = false;
                        return;
                    }
                }
            }

            public void ClearWorksheet()
            {
                try
                {
                    if (_sheet.AutoFilterMode && _sheet.AutoFilter != null && _sheet.AutoFilter.Range != null)
                    {
                        _sheet.AutoFilter.Range.AutoFilter();
                    }
                }
                catch (Exception)
                {
                }

                try
                {
                    if (_sheet.AutoFilterMode)
                    {
                        _sheet.AutoFilterMode = false;
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
                Excel.Range range = Rect(headerRow, firstCol, lastRow, lastCol);
                Excel.Sort sort = _sheet.Sort;
                sort.SortFields.Clear();
                foreach (SortFilterSortKey key in sorts)
                {
                    int col = columns[key.Column];
                    Excel.Range keyRange = Rect(headerRow, col, lastRow, col);
                    sort.SortFields.Add(
                        keyRange,
                        (Excel.XlSortOn)XlSortOnValues,
                        key.Order == "desc"
                            ? Excel.XlSortOrder.xlDescending
                            : Excel.XlSortOrder.xlAscending,
                        Type.Missing,
                        Excel.XlSortDataOption.xlSortNormal);
                }

                sort.SetRange(range);
                sort.Header = Excel.XlYesNoGuess.xlYes;
                sort.MatchCase = false;
                sort.Orientation = Excel.XlSortOrientation.xlSortColumns;
                sort.Apply();
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
                Excel.Range range = table
                    ? TableRange(tableName)
                    : Rect(headerRow, firstCol, lastRow, lastCol);
                if (table)
                {
                    TableByName(tableName).ShowAutoFilter = true;
                }
                else
                {
                    range.AutoFilter();
                }

                foreach (SortFilterColumn filter in filters)
                {
                    int field = columns[filter.Column] - firstCol + 1;
                    if (filter.IsIn)
                    {
                        range.AutoFilter(
                            field,
                            filter.InValues.ToArray(),
                            Excel.XlAutoFilterOperator.xlFilterValues,
                            Type.Missing,
                            Type.Missing);
                    }
                    else if (string.IsNullOrEmpty(filter.Criteria2))
                    {
                        range.AutoFilter(field, filter.Criteria1);
                    }
                    else
                    {
                        range.AutoFilter(
                            field,
                            filter.Criteria1,
                            filter.OrJoin
                                ? Excel.XlAutoFilterOperator.xlOr
                                : Excel.XlAutoFilterOperator.xlAnd,
                            filter.Criteria2,
                            Type.Missing);
                    }
                }
            }

            public bool TryReadFilters(SortFilterLivePanel panel, out List<SortFilterColumn> filters, out string error)
            {
                filters = new List<SortFilterColumn>();
                error = null;
                Excel.AutoFilter auto = panel.IsTable
                    ? TableByName(panel.TableName).AutoFilter
                    : _sheet.AutoFilter;
                if (auto == null || auto.Filters == null)
                {
                    return true;
                }

                Excel.Filters collection = auto.Filters;
                for (int i = 1; i <= collection.Count; i++)
                {
                    Excel.Filter filter = collection[i];
                    if (filter == null || !filter.On)
                    {
                        continue;
                    }

                    string column = ReadCell(panel.HeaderRow, panel.FirstCol + i - 1);
                    int op = 0;
                    try
                    {
                        op = Convert.ToInt32(filter.Operator, CultureInfo.InvariantCulture);
                    }
                    catch (Exception)
                    {
                        op = 0;
                    }

                    if (op == XlTop10Items || op == 4 || op == 5 || op == 6 || op >= 8)
                    {
                        error = "筛选列「" + column + "」不是勾选或比较式，本工具不能读回";
                        return false;
                    }

                    if (op == XlFilterValues)
                    {
                        filters.Add(new SortFilterColumn
                        {
                            Column = column,
                            IsIn = true,
                            InValues = ReadValueList(filter.Criteria1)
                        });
                        continue;
                    }

                    string c1 = CriteriaText(filter.Criteria1);
                    string c2 = null;
                    if (op == XlAnd || op == XlOr)
                    {
                        try
                        {
                            c2 = CriteriaText(filter.Criteria2);
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

                    filters.Add(new SortFilterColumn
                    {
                        Column = column,
                        IsIn = false,
                        Expr = SpreadsheetSortFilterParse.CanonicalExpr(c1, c2, op == XlOr)
                    });
                }

                return true;
            }

            public bool TryReadSorts(SortFilterLivePanel panel, out List<SortFilterSortKey> sorts, out string error)
            {
                sorts = new List<SortFilterSortKey>();
                error = null;
                if (!panel.IsTable)
                {
                    return true;
                }

                Excel.ListObject table = TableByName(panel.TableName);
                Excel.Sort sort = table.Sort;
                if (sort == null || sort.SortFields == null)
                {
                    return true;
                }

                foreach (Excel.SortField field in sort.SortFields)
                {
                    Excel.Range key = field.Key;
                    if (key == null)
                    {
                        continue;
                    }

                    int col = key.Column;
                    if (col < panel.FirstCol || col > panel.LastCol)
                    {
                        continue;
                    }

                    string order = Convert.ToInt32(field.Order, CultureInfo.InvariantCulture) == XlDescending
                        ? "desc"
                        : "asc";
                    sorts.Add(new SortFilterSortKey
                    {
                        Column = ReadCell(panel.HeaderRow, col),
                        Order = order
                    });
                }

                return true;
            }

            public int CountVisible(int headerRow, int firstCol, int lastRow)
            {
                if (lastRow <= headerRow)
                {
                    return 0;
                }

                try
                {
                    Excel.Range body = Rect(headerRow + 1, firstCol, lastRow, firstCol);
                    Excel.Range visible = body.SpecialCells((Excel.XlCellType)XlCellTypeVisible);
                    return visible == null ? 0 : visible.Count;
                }
                catch (COMException)
                {
                    return 0;
                }
                catch (Exception)
                {
                    return 0;
                }
            }

            private Excel.ListObject TableByName(string tableName)
            {
                foreach (Excel.ListObject table in _sheet.ListObjects)
                {
                    if (string.Equals(table.Name, tableName, StringComparison.OrdinalIgnoreCase))
                    {
                        return table;
                    }
                }

                return null;
            }

            private Excel.Range TableRange(string tableName)
            {
                return TableByName(tableName).Range;
            }

            private Excel.Range Rect(int r1, int c1, int r2, int c2)
            {
                return _sheet.Range[A1Address.Range(r1, c1, r2, c2)];
            }

            private static int TableDataLast(Excel.ListObject table)
            {
                Excel.Range range = table.Range;
                int last = range.Row + range.Rows.Count - 1;
                try
                {
                    if (table.ShowTotals && table.TotalsRowRange != null)
                    {
                        last = table.TotalsRowRange.Row - 1;
                    }
                }
                catch (Exception)
                {
                }

                return last;
            }

            private static SortFilterLivePanel PanelFromRange(
                Excel.Range header,
                bool isTable,
                string tableName,
                int dataLast)
            {
                int row = header.Row;
                int c1 = header.Column;
                int c2 = c1 + header.Columns.Count - 1;
                return new SortFilterLivePanel
                {
                    HeaderRow = row,
                    FirstCol = c1,
                    LastCol = c2,
                    IsTable = isTable,
                    TableName = tableName,
                    DataLastRow = dataLast
                };
            }

            private static bool CoveredBy(List<SortFilterLivePanel> panels, SortFilterLivePanel candidate)
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

            private static int ComparePanels(SortFilterLivePanel a, SortFilterLivePanel b)
            {
                int row = a.HeaderRow.CompareTo(b.HeaderRow);
                if (row != 0)
                {
                    return row;
                }

                return a.FirstCol.CompareTo(b.FirstCol);
            }

            private static List<string> ReadValueList(object criteria)
            {
                var values = new List<string>();
                if (criteria is IEnumerable enumerable && !(criteria is string))
                {
                    foreach (object item in enumerable)
                    {
                        string text = StripEquals(CriteriaText(item));
                        if (!string.IsNullOrEmpty(text))
                        {
                            values.Add(text);
                        }
                    }

                    return values;
                }

                string single = StripEquals(CriteriaText(criteria));
                if (!string.IsNullOrEmpty(single))
                {
                    values.Add(single);
                }

                return values;
            }

            private static string CriteriaText(object criteria)
            {
                if (criteria == null)
                {
                    return "";
                }

                return Convert.ToString(criteria, CultureInfo.InvariantCulture) ?? "";
            }

            private static string StripEquals(string text)
            {
                if (text != null && text.StartsWith("=", StringComparison.Ordinal) && !text.StartsWith(">=", StringComparison.Ordinal))
                {
                    return text.Substring(1);
                }

                return text ?? "";
            }

            private static bool HasWildcard(string criteria)
            {
                if (string.IsNullOrEmpty(criteria))
                {
                    return false;
                }

                return criteria.IndexOf('*') >= 0 || criteria.IndexOf('?') >= 0;
            }
        }
    }

    internal interface ISortFilterSheet
    {
        bool TryList(out List<SortFilterLivePanel> panels, out string error);

        string ReadCell(int row, int col);

        bool CellEmpty(int row, int col);

        bool TryTableBounds(
            int headerRow,
            int firstCol,
            int lastCol,
            out bool inside,
            out bool fullHeader,
            out int dataLastRow,
            out string tableName,
            out string error);

        void ClearTable(string tableName);

        void ClearWorksheet();

        void Sort(
            int headerRow,
            int firstCol,
            int lastRow,
            int lastCol,
            IList<SortFilterSortKey> sorts,
            IDictionary<string, int> columns);

        void ApplyFilters(
            int headerRow,
            int firstCol,
            int lastRow,
            int lastCol,
            IList<SortFilterColumn> filters,
            IDictionary<string, int> columns,
            bool table,
            string tableName);

        bool TryReadFilters(SortFilterLivePanel panel, out List<SortFilterColumn> filters, out string error);

        bool TryReadSorts(SortFilterLivePanel panel, out List<SortFilterSortKey> sorts, out string error);

        int CountVisible(int headerRow, int firstCol, int lastRow);
    }
}
