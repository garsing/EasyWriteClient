using System;
using System.Collections.Generic;
using System.Reflection;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using WordAddIn1.OpenFiles;

namespace WordAddIn1.SpreadsheetHost
{
    internal static class SpreadsheetPivotCom
    {
        private static string FormatPivotError(Exception ex)
        {
            Exception cur = ex;
            while (cur is TargetInvocationException tie && tie.InnerException != null)
            {
                cur = tie.InnerException;
            }

            if (cur is AggregateException ae && ae.InnerExceptions != null && ae.InnerExceptions.Count > 0)
            {
                cur = ae.InnerExceptions[0];
                while (cur is TargetInvocationException tie2 && tie2.InnerException != null)
                {
                    cur = tie2.InnerException;
                }
            }

            string msg = cur?.Message ?? "未知错误";
            if (cur is COMException com)
            {
                msg += " (HRESULT=0x" + unchecked((uint)com.ErrorCode).ToString("X8") + ")";
            }

            return msg;
        }

        public static bool TryExecuteOnExcelWorkbook(
            Excel.Workbook book,
            SpreadsheetPivotRequest request,
            out SpreadsheetPivotResult partial,
            out string error)
        {
            partial = new SpreadsheetPivotResult
            {
                Action = request.Action,
                Pivots = new List<SpreadsheetPivotInfo>()
            };
            error = null;

            try
            {
                if (request.Action == "list")
                {
                    return TryListExcel(book, request.ListSheetFilter, partial, out error);
                }

                if (request.Action == "create")
                {
                    return TryCreateExcel(book, request, partial, out error);
                }

                if (request.Action == "refresh")
                {
                    return TryRefreshOrDeleteExcel(book, request, delete: false, partial, out error);
                }

                if (request.Action == "delete")
                {
                    return TryRefreshOrDeleteExcel(book, request, delete: true, partial, out error);
                }

                error = "未知 action";
                return false;
            }
            catch (Exception ex)
            {
                error = "透视表操作失败: " + FormatPivotError(ex);
                return false;
            }
        }

        public static bool TryExecuteOnEtWorkbook(
            object book,
            SpreadsheetPivotRequest request,
            out SpreadsheetPivotResult partial,
            out string error)
        {
            partial = new SpreadsheetPivotResult
            {
                Action = request.Action,
                Pivots = new List<SpreadsheetPivotInfo>()
            };
            error = null;

            try
            {
                if (request.Action == "list")
                {
                    return TryListEt(book, request.ListSheetFilter, partial, out error);
                }

                if (request.Action == "create")
                {
                    return TryCreateEt(book, request, partial, out error);
                }

                if (request.Action == "refresh")
                {
                    return TryRefreshOrDeleteEt(book, request, delete: false, partial, out error);
                }

                if (request.Action == "delete")
                {
                    return TryRefreshOrDeleteEt(book, request, delete: true, partial, out error);
                }

                error = "未知 action";
                return false;
            }
            catch (Exception ex)
            {
                error = "透视表操作失败: " + FormatPivotError(ex);
                return false;
            }
        }

        private static bool TryListExcel(
            Excel.Workbook book,
            string sheetFilter,
            SpreadsheetPivotResult partial,
            out string error)
        {
            error = null;
            foreach (object raw in book.Sheets)
            {
                var sheet = raw as Excel.Worksheet;
                if (sheet == null)
                {
                    continue;
                }

                string sheetName;
                try
                {
                    sheetName = sheet.Name;
                }
                catch (Exception)
                {
                    continue;
                }

                if (!string.IsNullOrEmpty(sheetFilter)
                    && !string.Equals(sheetName, sheetFilter, StringComparison.Ordinal))
                {
                    continue;
                }

                Excel.PivotTables tables;
                try
                {
                    tables = sheet.PivotTables() as Excel.PivotTables;
                }
                catch (Exception)
                {
                    continue;
                }

                if (tables == null)
                {
                    continue;
                }

                int count;
                try
                {
                    count = tables.Count;
                }
                catch (Exception)
                {
                    continue;
                }

                for (int i = 1; i <= count; i++)
                {
                    Excel.PivotTable pt;
                    try
                    {
                        pt = tables.Item(i);
                    }
                    catch (Exception)
                    {
                        continue;
                    }

                    SpreadsheetPivotInfo info = ReadExcelPivotInfo(sheetName, pt);
                    if (info != null)
                    {
                        partial.Pivots.Add(info);
                    }
                }
            }

            if (!string.IsNullOrEmpty(sheetFilter) && partial.Pivots.Count == 0)
            {
                // 过滤表存在但无透视也算成功；表不存在则失败
                bool sheetFound = false;
                foreach (object raw in book.Sheets)
                {
                    var ws = raw as Excel.Worksheet;
                    if (ws != null && string.Equals(ws.Name, sheetFilter, StringComparison.Ordinal))
                    {
                        sheetFound = true;
                        break;
                    }

                    var chart = raw as Excel.Chart;
                    if (chart != null && string.Equals(chart.Name, sheetFilter, StringComparison.Ordinal))
                    {
                        error = "sheet_type=chart，无法列出透视表";
                        return false;
                    }
                }

                if (!sheetFound)
                {
                    error = "工作表不存在: " + sheetFilter;
                    return false;
                }
            }

            partial.PivotsCount = partial.Pivots.Count;
            return true;
        }

        private static SpreadsheetPivotInfo ReadExcelPivotInfo(string sheetName, Excel.PivotTable pt)
        {
            if (pt == null)
            {
                return null;
            }

            var info = new SpreadsheetPivotInfo { Sheet = sheetName };
            try
            {
                info.Name = pt.Name;
            }
            catch (Exception)
            {
                info.Name = "";
            }

            try
            {
                Excel.Range tr = pt.TableRange1;
                if (tr != null)
                {
                    int r = tr.Row;
                    int c = tr.Column;
                    int lr = r + tr.Rows.Count - 1;
                    int lc = c + tr.Columns.Count - 1;
                    info.DestCell = A1Address.Cell(r, c);
                    info.TableRange = A1Address.Range(r, c, lr, lc);
                }
            }
            catch (Exception)
            {
            }

            try
            {
                object src = pt.SourceData;
                TryParseSourceData(Convert.ToString(src), out string srcSheet, out string srcRange);
                info.SourceSheet = srcSheet;
                info.SourceRange = srcRange;
            }
            catch (Exception)
            {
            }

            return info;
        }

        private static bool TryCreateExcel(
            Excel.Workbook book,
            SpreadsheetPivotRequest request,
            SpreadsheetPivotResult partial,
            out string error)
        {
            if (!TryGetExcelWorksheet(book, request.SourceSheet, out Excel.Worksheet sourceSheet, out error))
            {
                return false;
            }

            if (!TryGetExcelWorksheet(book, request.DestSheet, out Excel.Worksheet destSheet, out error))
            {
                return false;
            }

            string sourceA1 = request.SourceRangeA1.Trim();
            if (sourceA1.IndexOf('!') >= 0)
            {
                error = "source.range 须为纯 A1";
                return false;
            }

            if (!A1Address.TryParseRange(
                    sourceA1,
                    out int firstRow,
                    out int firstCol,
                    out int lastRow,
                    out int lastCol,
                    out error))
            {
                return false;
            }

            int rowCount = lastRow - firstRow + 1;
            int colCount = lastCol - firstCol + 1;
            if (!SpreadsheetPivotLimits.TryCheckSourceLimits(rowCount, colCount, out error))
            {
                return false;
            }

            string actualSource = A1Address.Range(firstRow, firstCol, lastRow, lastCol);
            Excel.Range sourceRange = sourceSheet.Range[actualSource];

            if (!TryValidateSourceExcel(sourceRange, firstRow, firstCol, lastRow, lastCol, request, out error))
            {
                return false;
            }

            string destCell = request.DestCellA1.Trim();
            if (!A1Address.TryParseCell(destCell, out int destRow, out int destCol))
            {
                error = "dest.cell 须为单格 A1（如 H1）";
                return false;
            }

            destCell = A1Address.Cell(destRow, destCol);
            Excel.Range destRange = destSheet.Range[destCell];

            if (!string.IsNullOrEmpty(request.Name)
                && TryFindExcelPivotByName(book, request.Name, out _, out _))
            {
                error = "已存在同名透视表: " + request.Name + "，请先 delete 或换 name";
                return false;
            }

            if (TryFindExcelPivotCoveringCell(destSheet, destRow, destCol, out Excel.PivotTable existing, out _))
            {
                string existingName = "";
                try
                {
                    existingName = existing.Name;
                }
                catch (Exception)
                {
                }

                error = "dest 已存在透视表"
                    + (string.IsNullOrEmpty(existingName) ? "" : "（" + existingName + "）")
                    + "，请先 delete";
                return false;
            }

            Excel.PivotCache cache = book.PivotCaches().Create(
                Excel.XlPivotTableSourceType.xlDatabase,
                sourceRange,
                Excel.XlPivotTableVersionList.xlPivotTableVersion15);
            object tableNameArg = string.IsNullOrEmpty(request.Name)
                ? Type.Missing
                : (object)request.Name;
            Excel.PivotTable table = cache.CreatePivotTable(
                destRange,
                tableNameArg,
                Type.Missing,
                Excel.XlPivotTableVersionList.xlPivotTableVersion15);

            try
            {
                try
                {
                    table.ManualUpdate = true;
                }
                catch (Exception)
                {
                }

                foreach (string rowField in request.Rows)
                {
                    if (!TryResolveExcelPivotField(table, rowField, out Excel.PivotField pf, out error))
                    {
                        TryClearExcelPivotOrphan(table);
                        return false;
                    }

                    pf.Orientation = Excel.XlPivotFieldOrientation.xlRowField;
                }

                foreach (string colField in request.Columns)
                {
                    if (!TryResolveExcelPivotField(table, colField, out Excel.PivotField pf, out error))
                    {
                        TryClearExcelPivotOrphan(table);
                        return false;
                    }

                    pf.Orientation = Excel.XlPivotFieldOrientation.xlColumnField;
                }

                foreach (SpreadsheetPivotValueField value in request.Values)
                {
                    if (!SpreadsheetPivotParse.TryMapAgg(value.Agg, out int xlFn, out error))
                    {
                        TryClearExcelPivotOrphan(table);
                        return false;
                    }

                    if (!TryResolveExcelPivotField(table, value.Field, out Excel.PivotField pf, out error))
                    {
                        TryClearExcelPivotOrphan(table);
                        return false;
                    }

                    table.AddDataField(
                        pf,
                        Type.Missing,
                        (Excel.XlConsolidationFunction)xlFn);
                }

                try
                {
                    table.ManualUpdate = false;
                }
                catch (Exception)
                {
                }
            }
            catch (Exception ex)
            {
                TryClearExcelPivotOrphan(table);
                error = "配置透视字段失败: " + ex.Message;
                return false;
            }

            SpreadsheetPivotInfo info = ReadExcelPivotInfo(destSheet.Name, table);
            FillResultFromInfo(partial, info, request.SourceSheet, actualSource);
            return true;
        }

        private static bool TryResolveExcelPivotField(
            Excel.PivotTable table,
            string requestedName,
            out Excel.PivotField field,
            out string error)
        {
            field = null;
            error = null;
            string want = (requestedName ?? "").Trim();
            if (string.IsNullOrEmpty(want))
            {
                error = "字段名不能为空";
                return false;
            }

            try
            {
                field = table.PivotFields(want) as Excel.PivotField;
                if (field != null)
                {
                    return true;
                }
            }
            catch (Exception)
            {
            }

            var available = new List<string>();
            try
            {
                Excel.PivotFields all = table.PivotFields(Type.Missing) as Excel.PivotFields;
                if (all != null)
                {
                    for (int i = 1; i <= all.Count; i++)
                    {
                        Excel.PivotField candidate;
                        try
                        {
                            candidate = all.Item(i) as Excel.PivotField;
                        }
                        catch (Exception)
                        {
                            continue;
                        }

                        if (candidate == null)
                        {
                            continue;
                        }

                        string name = SafePivotName(candidate);
                        string caption = "";
                        string sourceName = "";
                        try
                        {
                            caption = (candidate.Caption ?? "").Trim();
                        }
                        catch (Exception)
                        {
                        }

                        try
                        {
                            sourceName = Convert.ToString(candidate.SourceName)?.Trim() ?? "";
                        }
                        catch (Exception)
                        {
                        }

                        if (!string.IsNullOrEmpty(name))
                        {
                            available.Add(name);
                        }

                        if (NamesMatch(want, name)
                            || NamesMatch(want, caption)
                            || NamesMatch(want, sourceName))
                        {
                            field = candidate;
                            return true;
                        }
                    }
                }
            }
            catch (Exception)
            {
            }

            error = "无法匹配透视字段: " + want;
            if (available.Count > 0)
            {
                error += "（可用: " + string.Join(", ", available) + "）";
            }
            else
            {
                error += "（PivotFields 方法无效，请确认源表头与字段名一致）";
            }

            return false;
        }

        private static bool NamesMatch(string a, string b)
        {
            if (string.IsNullOrEmpty(a) || string.IsNullOrEmpty(b))
            {
                return false;
            }

            return string.Equals(a.Trim(), b.Trim(), StringComparison.OrdinalIgnoreCase);
        }

        private static string SafePivotName(Excel.PivotField field)
        {
            try
            {
                return (field.Name ?? "").Trim();
            }
            catch (Exception)
            {
                return "";
            }
        }

        private static void TryClearExcelPivotOrphan(Excel.PivotTable table)
        {
            if (table == null)
            {
                return;
            }

            try
            {
                table.TableRange2.Clear();
            }
            catch (Exception)
            {
                try
                {
                    table.TableRange1.Clear();
                }
                catch (Exception)
                {
                }
            }
        }

        private static bool TryValidateSourceExcel(
            Excel.Range sourceRange,
            int firstRow,
            int firstCol,
            int lastRow,
            int lastCol,
            SpreadsheetPivotRequest request,
            out string error)
        {
            error = null;
            try
            {
                object merge = sourceRange.MergeCells;
                if (merge is bool merged)
                {
                    if (merged)
                    {
                        error = "源区域含合并格，首期不支持";
                        return false;
                    }
                }
                else if (merge != null && !(merge is DBNull))
                {
                    // 非 bool 且非空：按部分合并处理；DBNull 也视为部分合并
                    error = "源区域含合并格，首期不支持";
                    return false;
                }
                else if (merge is DBNull)
                {
                    error = "源区域含合并格，首期不支持";
                    return false;
                }
            }
            catch (Exception)
            {
            }

            // Excel 透视字段名通常跟 Value2 一致；同时收 Text 作别名，避免显示文本与底层值不一致
            var headers = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            for (int c = firstCol; c <= lastCol; c++)
            {
                Excel.Range cell = sourceRange.Worksheet.Cells[firstRow, c];
                string valueName = ReadCellPivotFieldName(cell);
                if (string.IsNullOrWhiteSpace(valueName))
                {
                    error = "源第一行须为非空字段名（列 " + A1Address.ColumnLetter(c) + "）";
                    return false;
                }

                valueName = valueName.Trim();
                if (!headers.Add(valueName))
                {
                    error = "源表头字段名重复: " + valueName;
                    return false;
                }

                string textName = ReadCellText(cell);
                if (!string.IsNullOrWhiteSpace(textName))
                {
                    headers.Add(textName.Trim());
                }
            }

            if (lastRow <= firstRow)
            {
                error = "源区域除表头外至少需要一行数据";
                return false;
            }

            foreach (string f in request.Rows)
            {
                if (!headers.Contains(f))
                {
                    error = "rows 字段不在源表头: " + f;
                    return false;
                }
            }

            foreach (string f in request.Columns)
            {
                if (!headers.Contains(f))
                {
                    error = "columns 字段不在源表头: " + f;
                    return false;
                }
            }

            foreach (SpreadsheetPivotValueField v in request.Values)
            {
                if (!headers.Contains(v.Field))
                {
                    error = "values 字段不在源表头: " + v.Field;
                    return false;
                }
            }

            return true;
        }

        private static string ReadCellPivotFieldName(Excel.Range cell)
        {
            try
            {
                object v = cell.Value2;
                if (v != null && !(v is DBNull))
                {
                    string s = Convert.ToString(v);
                    if (!string.IsNullOrWhiteSpace(s))
                    {
                        return s;
                    }
                }
            }
            catch (Exception)
            {
            }

            return ReadCellText(cell);
        }

        private static string ReadCellText(Excel.Range cell)
        {
            try
            {
                object t = cell.Text;
                if (t != null)
                {
                    string s = Convert.ToString(t);
                    if (!string.IsNullOrWhiteSpace(s))
                    {
                        return s;
                    }
                }
            }
            catch (Exception)
            {
            }

            try
            {
                object v = cell.Value2;
                return v == null ? "" : Convert.ToString(v);
            }
            catch (Exception)
            {
                return "";
            }
        }

        private static bool TryRefreshOrDeleteExcel(
            Excel.Workbook book,
            SpreadsheetPivotRequest request,
            bool delete,
            SpreadsheetPivotResult partial,
            out string error)
        {
            if (!TryLocateExcelPivot(book, request, out Excel.Worksheet sheet, out Excel.PivotTable pt, out error))
            {
                return false;
            }

            SpreadsheetPivotInfo info = ReadExcelPivotInfo(sheet.Name, pt);
            if (delete)
            {
                try
                {
                    pt.TableRange2.Clear();
                }
                catch (Exception ex)
                {
                    error = "删除透视表失败: " + ex.Message;
                    return false;
                }

                FillResultFromInfo(partial, info, info?.SourceSheet, info?.SourceRange);
                return true;
            }

            try
            {
                pt.PivotCache().Refresh();
            }
            catch (Exception)
            {
                try
                {
                    pt.RefreshTable();
                }
                catch (Exception ex)
                {
                    error = "刷新透视表失败: " + ex.Message;
                    return false;
                }
            }

            info = ReadExcelPivotInfo(sheet.Name, pt);
            FillResultFromInfo(partial, info, info?.SourceSheet, info?.SourceRange);
            return true;
        }

        private static bool TryLocateExcelPivot(
            Excel.Workbook book,
            SpreadsheetPivotRequest request,
            out Excel.Worksheet sheet,
            out Excel.PivotTable pt,
            out string error)
        {
            sheet = null;
            pt = null;
            error = null;

            if (!string.IsNullOrEmpty(request.Name))
            {
                if (TryFindExcelPivotByName(book, request.Name, out sheet, out pt))
                {
                    return true;
                }

                error = "未找到透视表: " + request.Name;
                return false;
            }

            if (!TryGetExcelWorksheet(book, request.DestSheet, out sheet, out error))
            {
                return false;
            }

            if (!A1Address.TryParseCell(request.DestCellA1.Trim(), out int row, out int col))
            {
                error = "dest.cell 须为单格 A1（如 H1）";
                return false;
            }

            if (TryFindExcelPivotCoveringCell(sheet, row, col, out pt, out string pivotName))
            {
                return true;
            }

            error = "未找到透视表（位置 " + request.DestSheet + "!" + A1Address.Cell(row, col) + "）";
            return false;
        }

        private static bool TryFindExcelPivotByName(
            Excel.Workbook book,
            string name,
            out Excel.Worksheet sheet,
            out Excel.PivotTable pt)
        {
            sheet = null;
            pt = null;
            foreach (object raw in book.Sheets)
            {
                var ws = raw as Excel.Worksheet;
                if (ws == null)
                {
                    continue;
                }

                try
                {
                    Excel.PivotTables tables = ws.PivotTables() as Excel.PivotTables;
                    if (tables == null)
                    {
                        continue;
                    }

                    for (int i = 1; i <= tables.Count; i++)
                    {
                        Excel.PivotTable candidate = tables.Item(i);
                        if (string.Equals(candidate.Name, name, StringComparison.Ordinal))
                        {
                            sheet = ws;
                            pt = candidate;
                            return true;
                        }
                    }
                }
                catch (Exception)
                {
                }
            }

            return false;
        }

        private static bool TryFindExcelPivotCoveringCell(
            Excel.Worksheet sheet,
            int row,
            int col,
            out Excel.PivotTable pt,
            out string pivotName)
        {
            pt = null;
            pivotName = null;
            Excel.PivotTables tables;
            try
            {
                tables = sheet.PivotTables() as Excel.PivotTables;
            }
            catch (Exception)
            {
                return false;
            }

            if (tables == null)
            {
                return false;
            }

            for (int i = 1; i <= tables.Count; i++)
            {
                Excel.PivotTable candidate;
                try
                {
                    candidate = tables.Item(i);
                    Excel.Range tr = candidate.TableRange1;
                    int r = tr.Row;
                    int c = tr.Column;
                    int lr = r + tr.Rows.Count - 1;
                    int lc = c + tr.Columns.Count - 1;
                    if (row >= r && row <= lr && col >= c && col <= lc)
                    {
                        pt = candidate;
                        try
                        {
                            pivotName = candidate.Name;
                        }
                        catch (Exception)
                        {
                            pivotName = "";
                        }

                        return true;
                    }
                }
                catch (Exception)
                {
                }
            }

            return false;
        }

        private static bool TryGetExcelWorksheet(
            Excel.Workbook book,
            string sheetName,
            out Excel.Worksheet sheet,
            out string error)
        {
            sheet = null;
            error = null;
            const int xlChart = -4109;
            try
            {
                foreach (object raw in book.Sheets)
                {
                    var ws = raw as Excel.Worksheet;
                    if (ws != null)
                    {
                        if (!string.Equals(ws.Name, sheetName, StringComparison.Ordinal))
                        {
                            continue;
                        }

                        try
                        {
                            if (Convert.ToInt32(ws.Type) == xlChart)
                            {
                                error = "sheet_type=chart，无法操作透视表";
                                return false;
                            }
                        }
                        catch (Exception)
                        {
                        }

                        sheet = ws;
                        return true;
                    }

                    var chart = raw as Excel.Chart;
                    if (chart != null && string.Equals(chart.Name, sheetName, StringComparison.Ordinal))
                    {
                        error = "sheet_type=chart，无法操作透视表";
                        return false;
                    }
                }
            }
            catch (Exception ex)
            {
                error = "COM 不可用: " + ex.Message;
                return false;
            }

            error = "工作表不存在: " + sheetName;
            return false;
        }

        // ----- et late bind -----

        private static bool TryListEt(
            object book,
            string sheetFilter,
            SpreadsheetPivotResult partial,
            out string error)
        {
            error = null;
            bool filterChecked = string.IsNullOrEmpty(sheetFilter);
            bool sheetFound = filterChecked;

            foreach (object sheet in EtCom.EnumerateSheets(book))
            {
                string sheetName = EtCom.TryReadName(sheet) ?? "";
                if (!string.IsNullOrEmpty(sheetFilter)
                    && !string.Equals(sheetName, sheetFilter, StringComparison.Ordinal))
                {
                    continue;
                }

                if (!string.IsNullOrEmpty(sheetFilter))
                {
                    sheetFound = true;
                    try
                    {
                        object type = EtCom.GetProperty(sheet, "Type");
                        if (type != null && Convert.ToInt32(type) == -4109)
                        {
                            error = "sheet_type=chart，无法列出透视表";
                            return false;
                        }
                    }
                    catch (Exception)
                    {
                    }
                }

                object tables = TryGetPivotTables(sheet);
                if (tables == null)
                {
                    continue;
                }

                int count;
                try
                {
                    count = Convert.ToInt32(EtCom.GetProperty(tables, "Count"));
                }
                catch (Exception)
                {
                    continue;
                }

                for (int i = 1; i <= count; i++)
                {
                    object pt = TryGetPivotItem(tables, i);
                    SpreadsheetPivotInfo info = ReadEtPivotInfo(sheetName, pt);
                    if (info != null)
                    {
                        partial.Pivots.Add(info);
                    }
                }
            }

            if (!string.IsNullOrEmpty(sheetFilter) && !sheetFound)
            {
                error = "工作表不存在: " + sheetFilter;
                return false;
            }

            partial.PivotsCount = partial.Pivots.Count;
            return true;
        }

        private static SpreadsheetPivotInfo ReadEtPivotInfo(string sheetName, object pt)
        {
            if (pt == null)
            {
                return null;
            }

            var info = new SpreadsheetPivotInfo { Sheet = sheetName };
            try
            {
                info.Name = Convert.ToString(EtCom.GetProperty(pt, "Name")) ?? "";
            }
            catch (Exception)
            {
                info.Name = "";
            }

            try
            {
                object tr = EtCom.GetProperty(pt, "TableRange1");
                if (tr != null)
                {
                    int r = Convert.ToInt32(EtCom.GetProperty(tr, "Row"));
                    int c = Convert.ToInt32(EtCom.GetProperty(tr, "Column"));
                    int rows = Convert.ToInt32(EtCom.GetProperty(EtCom.GetProperty(tr, "Rows"), "Count"));
                    int cols = Convert.ToInt32(EtCom.GetProperty(EtCom.GetProperty(tr, "Columns"), "Count"));
                    info.DestCell = A1Address.Cell(r, c);
                    info.TableRange = A1Address.Range(r, c, r + rows - 1, c + cols - 1);
                }
            }
            catch (Exception)
            {
            }

            try
            {
                object src = EtCom.GetProperty(pt, "SourceData");
                TryParseSourceData(Convert.ToString(src), out string srcSheet, out string srcRange);
                info.SourceSheet = srcSheet;
                info.SourceRange = srcRange;
            }
            catch (Exception)
            {
            }

            return info;
        }

        private static bool TryCreateEt(
            object book,
            SpreadsheetPivotRequest request,
            SpreadsheetPivotResult partial,
            out string error)
        {
            try
            {
                return TryCreateEtCore(book, request, partial, out error);
            }
            catch (Exception ex)
            {
                error = "et 创建透视未捕获异常: " + FormatPivotError(ex);
                try
                {
                    string path = System.IO.Path.Combine(
                        EasyWriteLog.LogDirectory,
                        "pivot_et_error.txt");
                    System.IO.File.AppendAllText(
                        path,
                        DateTime.Now.ToString("HH:mm:ss.fff") + " " + error + Environment.NewLine + ex + Environment.NewLine);
                }
                catch (Exception)
                {
                }

                System.Diagnostics.Debug.WriteLine("[F_excel_pivot] " + error);
                return false;
            }
        }

        private static bool TryCreateEtCore(
            object book,
            SpreadsheetPivotRequest request,
            SpreadsheetPivotResult partial,
            out string error)
        {
            if (!TryGetEtWorksheet(book, request.SourceSheet, out object sourceSheet, out error))
            {
                return false;
            }

            if (!TryGetEtWorksheet(book, request.DestSheet, out object destSheet, out error))
            {
                return false;
            }

            string sourceA1 = request.SourceRangeA1.Trim();
            if (sourceA1.IndexOf('!') >= 0)
            {
                error = "source.range 须为纯 A1";
                return false;
            }

            if (!A1Address.TryParseRange(
                    sourceA1,
                    out int firstRow,
                    out int firstCol,
                    out int lastRow,
                    out int lastCol,
                    out error))
            {
                return false;
            }

            int rowCount = lastRow - firstRow + 1;
            int colCount = lastCol - firstCol + 1;
            if (!SpreadsheetPivotLimits.TryCheckSourceLimits(rowCount, colCount, out error))
            {
                return false;
            }

            string actualSource = A1Address.Range(firstRow, firstCol, lastRow, lastCol);
            object sourceRange = EtGetSheetRange(sourceSheet, actualSource);
            if (sourceRange == null)
            {
                error = "无法解析源区域";
                return false;
            }

            if (!TryValidateSourceEt(sourceSheet, sourceRange, firstRow, firstCol, lastRow, lastCol, request, out error))
            {
                return false;
            }

            if (!A1Address.TryParseCell(request.DestCellA1.Trim(), out int destRow, out int destCol))
            {
                error = "dest.cell 须为单格 A1（如 H1）";
                return false;
            }

            string destCell = A1Address.Cell(destRow, destCol);
            object destRange = EtGetSheetRange(destSheet, destCell);
            if (destRange == null)
            {
                error = "无法解析 dest.cell";
                return false;
            }

            if (!string.IsNullOrEmpty(request.Name)
                && TryFindEtPivotByName(book, request.Name, out _, out _))
            {
                error = "已存在同名透视表: " + request.Name + "，请先 delete 或换 name";
                return false;
            }

            if (TryFindEtPivotCoveringCell(destSheet, destRow, destCol, out object existing, out _))
            {
                string existingName = "";
                try
                {
                    existingName = Convert.ToString(EtCom.GetProperty(existing, "Name")) ?? "";
                }
                catch (Exception)
                {
                }

                error = "dest 已存在透视表"
                    + (string.IsNullOrEmpty(existingName) ? "" : "（" + existingName + "）")
                    + "，请先 delete";
                return false;
            }

            object caches = TryGetEtPivotCaches(book);
            string sourceSheetName = EtCom.TryReadName(sourceSheet) ?? request.SourceSheet;
            string destSheetNameForAddr = EtCom.TryReadName(destSheet) ?? request.DestSheet;
            string sourceAddrA1 = QuoteSheetForAddress(sourceSheetName) + "!" + actualSource;
            string sourceAddrR1C1 = QuoteSheetForAddress(sourceSheetName) + "!"
                + "R" + firstRow + "C" + firstCol + ":R" + lastRow + "C" + lastCol;
            string destAddrA1 = QuoteSheetForAddress(destSheetNameForAddr) + "!" + destCell;
            string cacheProbe = ProbeEtPivotCaches(caches);
            string flexError = null;

            // 优先 dynamic/IDispatch（WPS 对 Type.InvokeMember 常报 0x80020003）
            if (TryCreateEtPivotDynamic(
                    book,
                    destSheet,
                    sourceRange,
                    destRange,
                    sourceAddrA1,
                    sourceAddrR1C1,
                    destAddrA1,
                    destCell,
                    request.Name,
                    out object dynTable,
                    out string dynError))
            {
                return TryConfigureEtPivotFields(
                    dynTable,
                    destSheet,
                    request,
                    actualSource,
                    partial,
                    out error);
            }

            // InvokeFlex 再试 Create/Add
            if (caches != null
                && TryCreateEtPivotCacheFlex(
                    caches,
                    sourceRange,
                    sourceAddrA1,
                    sourceAddrR1C1,
                    out object flexCache,
                    out flexError)
                && TryCreateEtPivotTableFlex(
                    flexCache,
                    destRange,
                    destAddrA1,
                    request.Name,
                    out object flexTable,
                    out string flexTableError))
            {
                return TryConfigureEtPivotFields(
                    flexTable,
                    destSheet,
                    request,
                    actualSource,
                    partial,
                    out error);
            }

            if (!string.IsNullOrEmpty(flexError))
            {
                // keep
            }

            const int xlPivotTableVersion12 = 3;
            const int xlPivotTableVersion15 = 5;
            string reflectError = null;

            if (caches != null
                && TryCreateEtPivotCache(
                    caches,
                    sourceRange,
                    sourceAddrA1,
                    sourceAddrR1C1,
                    xlPivotTableVersion12,
                    xlPivotTableVersion15,
                    out object cache,
                    out reflectError))
            {
                if (TryCreateEtPivotTable(
                        cache,
                        destRange,
                        destAddrA1,
                        request.Name,
                        out object table,
                        out reflectError))
                {
                    return TryConfigureEtPivotFields(
                        table,
                        destSheet,
                        request,
                        actualSource,
                        partial,
                        out error);
                }
            }

            if (TryCreateEtPivotViaWizard(
                    destSheet,
                    sourceAddrR1C1,
                    sourceAddrA1,
                    destRange,
                    destAddrA1,
                    destCell,
                    request.Name,
                    out object wizardTable,
                    out string wizardError))
            {
                return TryConfigureEtPivotFields(
                    wizardTable,
                    destSheet,
                    request,
                    actualSource,
                    partial,
                    out error);
            }

            error = "et 创建透视失败。probe=[" + cacheProbe + "]"
                + " dynamic=[" + (dynError ?? "") + "]"
                + " flex=[" + (flexError ?? "") + "]"
                + " reflect=[" + (reflectError ?? (caches == null ? "PivotCaches=null" : "")) + "]"
                + " wizard=[" + (wizardError ?? "") + "]";
            try
            {
                string path = System.IO.Path.Combine(EasyWriteLog.LogDirectory, "pivot_et_error.txt");
                System.IO.File.AppendAllText(
                    path,
                    DateTime.Now.ToString("HH:mm:ss.fff") + " " + error + Environment.NewLine);
            }
            catch (Exception)
            {
            }

            System.Diagnostics.Debug.WriteLine("[F_excel_pivot] " + error);
            return false;
        }

        private static string ProbeEtPivotCaches(object caches)
        {
            if (caches == null)
            {
                return "null";
            }

            var parts = new List<string>();
            try
            {
                parts.Add("type=" + caches.GetType().FullName);
            }
            catch (Exception)
            {
            }

            foreach (string name in new[] { "Count", "Create", "Add", "Item", "_Default", "Parent" })
            {
                try
                {
                    object v = EtCom.GetProperty(caches, name);
                    parts.Add(name + ".get=" + (v == null ? "null" : v.GetType().Name));
                }
                catch (Exception ex)
                {
                    parts.Add(name + ".get!" + FormatPivotError(ex));
                }

                try
                {
                    object v = EtCom.InvokeFlex(caches, name);
                    parts.Add(name + ".call0=" + (v == null ? "null" : v.GetType().Name));
                }
                catch (Exception ex)
                {
                    parts.Add(name + ".call0!" + FormatPivotError(ex));
                }
            }

            return string.Join("; ", parts);
        }

        private static bool TryCreateEtPivotCacheFlex(
            object caches,
            object sourceRange,
            string sourceAddrA1,
            string sourceAddrR1C1,
            out object cache,
            out string error)
        {
            cache = null;
            error = null;
            var attempts = new List<(string label, Func<object> run)>
            {
                ("Flex.Create(Range)", () => EtCom.InvokeFlex(caches, "Create", SpreadsheetPivotParse.XlDatabase, sourceRange)),
                ("Flex.Create(R1C1)", () => EtCom.InvokeFlex(caches, "Create", SpreadsheetPivotParse.XlDatabase, sourceAddrR1C1)),
                ("Flex.Create(A1)", () => EtCom.InvokeFlex(caches, "Create", SpreadsheetPivotParse.XlDatabase, sourceAddrA1)),
                ("Flex.Add(R1C1)", () => EtCom.InvokeFlex(caches, "Add", SpreadsheetPivotParse.XlDatabase, sourceAddrR1C1)),
                ("Flex.Add(Range)", () => EtCom.InvokeFlex(caches, "Add", SpreadsheetPivotParse.XlDatabase, sourceRange)),
            };
            var errors = new List<string>();
            foreach (var attempt in attempts)
            {
                try
                {
                    object result = attempt.run();
                    if (result != null)
                    {
                        cache = result;
                        return true;
                    }

                    errors.Add(attempt.label + "=null");
                }
                catch (Exception ex)
                {
                    errors.Add(attempt.label + ": " + FormatPivotError(ex));
                }
            }

            error = string.Join(" | ", errors);
            return false;
        }

        private static bool TryCreateEtPivotTableFlex(
            object cache,
            object destRange,
            string destAddrA1,
            string name,
            out object table,
            out string error)
        {
            table = null;
            error = null;
            var attempts = new List<(string label, Func<object> run)>();
            if (string.IsNullOrEmpty(name))
            {
                attempts.Add(("Flex.CreatePivotTable(Range)", () => EtCom.InvokeFlex(cache, "CreatePivotTable", destRange)));
                attempts.Add(("Flex.CreatePivotTable(A1)", () => EtCom.InvokeFlex(cache, "CreatePivotTable", destAddrA1)));
            }
            else
            {
                attempts.Add(("Flex.CreatePivotTable(Range,name)", () => EtCom.InvokeFlex(cache, "CreatePivotTable", destRange, name)));
                attempts.Add(("Flex.CreatePivotTable(A1,name)", () => EtCom.InvokeFlex(cache, "CreatePivotTable", destAddrA1, name)));
            }

            var errors = new List<string>();
            foreach (var attempt in attempts)
            {
                try
                {
                    object result = attempt.run();
                    if (result != null)
                    {
                        table = result;
                        return true;
                    }

                    errors.Add(attempt.label + "=null");
                }
                catch (Exception ex)
                {
                    errors.Add(attempt.label + ": " + FormatPivotError(ex));
                }
            }

            error = string.Join(" | ", errors);
            return false;
        }

        private static bool TryCreateEtPivotDynamic(
            object book,
            object destSheet,
            object sourceRange,
            object destRange,
            string sourceAddrA1,
            string sourceAddrR1C1,
            string destAddrA1,
            string destCell,
            string name,
            out object table,
            out string error)
        {
            table = null;
            error = null;
            var errors = new List<string>();
            object created = null;

            void TryOne(string label, Func<object> run)
            {
                if (created != null)
                {
                    return;
                }

                try
                {
                    object result = run();
                    if (result != null)
                    {
                        created = result;
                        return;
                    }

                    if (A1Address.TryParseCell(destCell, out int row, out int col)
                        && TryFindEtPivotCoveringCell(destSheet, row, col, out object found, out _))
                    {
                        created = found;
                        return;
                    }

                    errors.Add(label + "=null");
                }
                catch (Exception ex)
                {
                    if (A1Address.TryParseCell(destCell, out int row, out int col)
                        && TryFindEtPivotCoveringCell(destSheet, row, col, out object found, out _))
                    {
                        created = found;
                        return;
                    }

                    errors.Add(label + ": " + FormatPivotError(ex));
                }
            }

            // PivotCaches() 方法 + Create(Range)
            TryOne("Caches().Create(Range)", () =>
            {
                dynamic dBook = book;
                dynamic caches = dBook.PivotCaches();
                dynamic cache = caches.Create(SpreadsheetPivotParse.XlDatabase, sourceRange);
                if (string.IsNullOrEmpty(name))
                {
                    return cache.CreatePivotTable(destRange);
                }

                return cache.CreatePivotTable(destRange, name);
            });

            TryOne("Caches().Create(R1C1)", () =>
            {
                dynamic dBook = book;
                dynamic caches = dBook.PivotCaches();
                dynamic cache = caches.Create(SpreadsheetPivotParse.XlDatabase, sourceAddrR1C1);
                if (string.IsNullOrEmpty(name))
                {
                    return cache.CreatePivotTable(destRange);
                }

                return cache.CreatePivotTable(destRange, name);
            });

            TryOne("Caches().Create(A1)", () =>
            {
                dynamic dBook = book;
                dynamic caches = dBook.PivotCaches();
                dynamic cache = caches.Create(SpreadsheetPivotParse.XlDatabase, sourceAddrA1);
                if (string.IsNullOrEmpty(name))
                {
                    return cache.CreatePivotTable(destRange);
                }

                return cache.CreatePivotTable(destRange, name);
            });

            TryOne("Caches().Create(Range,v12)", () =>
            {
                dynamic dBook = book;
                dynamic caches = dBook.PivotCaches();
                dynamic cache = caches.Create(SpreadsheetPivotParse.XlDatabase, sourceRange, 3);
                if (string.IsNullOrEmpty(name))
                {
                    return cache.CreatePivotTable(destRange);
                }

                return cache.CreatePivotTable(destRange, name);
            });

            TryOne("Caches.Create(Range)", () =>
            {
                dynamic dBook = book;
                dynamic caches = dBook.PivotCaches;
                dynamic cache = caches.Create(SpreadsheetPivotParse.XlDatabase, sourceRange);
                if (string.IsNullOrEmpty(name))
                {
                    return cache.CreatePivotTable(destRange);
                }

                return cache.CreatePivotTable(destRange, name);
            });

            TryOne("Caches().Add(R1C1)", () =>
            {
                dynamic dBook = book;
                dynamic caches = dBook.PivotCaches();
                dynamic cache = caches.Add(SpreadsheetPivotParse.XlDatabase, sourceAddrR1C1);
                if (string.IsNullOrEmpty(name))
                {
                    return cache.CreatePivotTable(destRange);
                }

                return cache.CreatePivotTable(destRange, name);
            });

            TryOne("Wizard(R1C1)", () =>
            {
                dynamic dSheet = destSheet;
                if (string.IsNullOrEmpty(name))
                {
                    dSheet.PivotTableWizard(
                        SpreadsheetPivotParse.XlDatabase,
                        sourceAddrR1C1,
                        destRange);
                }
                else
                {
                    dSheet.PivotTableWizard(
                        SpreadsheetPivotParse.XlDatabase,
                        sourceAddrR1C1,
                        destRange,
                        name);
                }

                return null; // 由 dest 回查
            });

            TryOne("Wizard(A1)", () =>
            {
                dynamic dSheet = destSheet;
                if (string.IsNullOrEmpty(name))
                {
                    dSheet.PivotTableWizard(
                        SpreadsheetPivotParse.XlDatabase,
                        sourceAddrA1,
                        destRange);
                }
                else
                {
                    dSheet.PivotTableWizard(
                        SpreadsheetPivotParse.XlDatabase,
                        sourceAddrA1,
                        destRange,
                        name);
                }

                return null;
            });

            if (created != null)
            {
                table = created;
                return true;
            }

            error = string.Join(" | ", errors);
            return false;
        }

        private static object EtGetSheetRange(object sheet, string a1)
        {
            if (sheet == null || string.IsNullOrWhiteSpace(a1))
            {
                return null;
            }

            try
            {
                return sheet.GetType().InvokeMember(
                    "Range",
                    System.Reflection.BindingFlags.GetProperty
                        | System.Reflection.BindingFlags.InvokeMethod
                        | System.Reflection.BindingFlags.Instance
                        | System.Reflection.BindingFlags.Public,
                    null,
                    sheet,
                    new object[] { a1 });
            }
            catch (Exception)
            {
            }

            try
            {
                return EtCom.InvokeFlex(sheet, "Range", a1);
            }
            catch (Exception)
            {
            }

            try
            {
                dynamic dSheet = sheet;
                return dSheet.Range[a1];
            }
            catch (Exception)
            {
            }

            try
            {
                dynamic dSheet = sheet;
                return dSheet.Range(a1);
            }
            catch (Exception)
            {
                return null;
            }
        }

        private static object TryGetEtPivotCaches(object book)
        {
            // WPS/Excel：PivotCaches() 多为无参方法，不能只用 GetProperty
            try
            {
                dynamic dBook = book;
                object caches = dBook.PivotCaches();
                if (caches != null)
                {
                    return caches;
                }
            }
            catch (Exception)
            {
            }

            try
            {
                object caches = EtCom.Invoke(book, "PivotCaches");
                if (caches != null)
                {
                    return caches;
                }
            }
            catch (Exception)
            {
            }

            try
            {
                return EtCom.GetProperty(book, "PivotCaches");
            }
            catch (Exception)
            {
                return null;
            }
        }

        private static bool TryConfigureEtPivotFields(
            object table,
            object destSheet,
            SpreadsheetPivotRequest request,
            string actualSource,
            SpreadsheetPivotResult partial,
            out string error)
        {
            error = null;
            if (table == null)
            {
                error = "创建透视表失败";
                return false;
            }

            try
            {
                EtCom.TrySetProperty(table, "ManualUpdate", true);

                foreach (string rowField in request.Rows)
                {
                    if (!TryResolveEtPivotField(table, rowField, out object pf, out error))
                    {
                        TryClearEtPivotOrphan(table);
                        return false;
                    }

                    EtCom.TrySetProperty(pf, "Orientation", SpreadsheetPivotParse.XlRowField);
                }

                foreach (string colField in request.Columns)
                {
                    if (!TryResolveEtPivotField(table, colField, out object pf, out error))
                    {
                        TryClearEtPivotOrphan(table);
                        return false;
                    }

                    EtCom.TrySetProperty(pf, "Orientation", SpreadsheetPivotParse.XlColumnField);
                }

                foreach (SpreadsheetPivotValueField value in request.Values)
                {
                    if (!SpreadsheetPivotParse.TryMapAgg(value.Agg, out int xlFn, out error))
                    {
                        TryClearEtPivotOrphan(table);
                        return false;
                    }

                    if (!TryResolveEtPivotField(table, value.Field, out object pf, out error))
                    {
                        TryClearEtPivotOrphan(table);
                        return false;
                    }

                    try
                    {
                        EtCom.Invoke(table, "AddDataField", pf, Type.Missing, xlFn);
                    }
                    catch (Exception)
                    {
                        try
                        {
                            EtCom.Invoke(table, "AddDataField", pf);
                            EtCom.TrySetProperty(pf, "Function", xlFn);
                        }
                        catch (Exception)
                        {
                            EtCom.TrySetProperty(pf, "Orientation", SpreadsheetPivotParse.XlDataField);
                            EtCom.TrySetProperty(pf, "Function", xlFn);
                        }
                    }
                }

                EtCom.TrySetProperty(table, "ManualUpdate", false);
            }
            catch (Exception ex)
            {
                TryClearEtPivotOrphan(table);
                error = "配置透视字段失败: " + FormatPivotError(ex);
                return false;
            }

            string destSheetName = EtCom.TryReadName(destSheet) ?? request.DestSheet;
            SpreadsheetPivotInfo info = ReadEtPivotInfo(destSheetName, table);
            FillResultFromInfo(partial, info, request.SourceSheet, actualSource);
            return true;
        }

        private static bool TryCreateEtPivotViaWizard(
            object destSheet,
            string sourceAddrR1C1,
            string sourceAddrA1,
            object destRange,
            string destAddrA1,
            string destCell,
            string name,
            out object table,
            out string error)
        {
            table = null;
            error = null;
            object nameArg = string.IsNullOrEmpty(name) ? Type.Missing : (object)name;
            var attempts = new List<(string label, Func<object> run)>
            {
                ("Wizard(R1C1,Range)", () => EtCom.Invoke(
                    destSheet, "PivotTableWizard",
                    SpreadsheetPivotParse.XlDatabase, sourceAddrR1C1, destRange, nameArg)),
                ("Wizard(A1,Range)", () => EtCom.Invoke(
                    destSheet, "PivotTableWizard",
                    SpreadsheetPivotParse.XlDatabase, sourceAddrA1, destRange, nameArg)),
                ("Wizard(R1C1,A1)", () => EtCom.Invoke(
                    destSheet, "PivotTableWizard",
                    SpreadsheetPivotParse.XlDatabase, sourceAddrR1C1, destAddrA1, nameArg)),
            };

            // PivotTables.Add(Cache/Source, Destination, Name) — 部分版本
            try
            {
                object pivots = EtCom.Invoke(destSheet, "PivotTables");
                if (pivots == null)
                {
                    pivots = EtCom.GetProperty(destSheet, "PivotTables");
                }

                if (pivots != null)
                {
                    object pivotsRef = pivots;
                    attempts.Add(("PivotTables.Add(R1C1)", () => EtCom.Invoke(
                        pivotsRef, "Add", sourceAddrR1C1, destRange, nameArg)));
                    attempts.Add(("PivotTables.Add(A1)", () => EtCom.Invoke(
                        pivotsRef, "Add", sourceAddrA1, destRange, nameArg)));
                }
            }
            catch (Exception)
            {
            }

            var errors = new List<string>();
            foreach (var attempt in attempts)
            {
                try
                {
                    object result = attempt.run();
                    // Wizard 可能返回 null 但仍创建了表：按 dest 单元格查找
                    if (result != null)
                    {
                        table = result;
                        return true;
                    }

                    if (A1Address.TryParseCell(destCell, out int row, out int col)
                        && TryFindEtPivotCoveringCell(destSheet, row, col, out object found, out _))
                    {
                        table = found;
                        return true;
                    }

                    errors.Add(attempt.label + "=null");
                }
                catch (Exception ex)
                {
                    // Wizard 有时抛错但仍建表
                    if (A1Address.TryParseCell(destCell, out int row, out int col)
                        && TryFindEtPivotCoveringCell(destSheet, row, col, out object found, out _))
                    {
                        table = found;
                        return true;
                    }

                    errors.Add(attempt.label + ": " + FormatPivotError(ex));
                }
            }

            error = string.Join(" | ", errors);
            return false;
        }

        private static string QuoteSheetForAddress(string sheetName)
        {
            string name = sheetName ?? "";
            if (name.IndexOfAny(new[] { ' ', '\'', '!', ':', '（', '）', '(', ')' }) >= 0
                || !IsSimpleSheetName(name))
            {
                return "'" + name.Replace("'", "''") + "'";
            }

            return name;
        }

        private static bool IsSimpleSheetName(string name)
        {
            if (string.IsNullOrEmpty(name))
            {
                return false;
            }

            foreach (char ch in name)
            {
                if (!(char.IsLetterOrDigit(ch) || ch == '_' || ch > 127))
                {
                    return false;
                }
            }

            return true;
        }

        private static bool TryCreateEtPivotCache(
            object caches,
            object sourceRange,
            string sourceAddrA1,
            string sourceAddrR1C1,
            int version12,
            int version15,
            out object cache,
            out string error)
        {
            cache = null;
            error = null;
            var attempts = new List<(string label, Func<object> run)>
            {
                ("Create(Range)", () => EtCom.Invoke(caches, "Create", SpreadsheetPivotParse.XlDatabase, sourceRange)),
                ("Create(Range,v12)", () => EtCom.Invoke(caches, "Create", SpreadsheetPivotParse.XlDatabase, sourceRange, version12)),
                ("Create(Range,v15)", () => EtCom.Invoke(caches, "Create", SpreadsheetPivotParse.XlDatabase, sourceRange, version15)),
                ("Create(A1)", () => EtCom.Invoke(caches, "Create", SpreadsheetPivotParse.XlDatabase, sourceAddrA1)),
                ("Create(A1,v12)", () => EtCom.Invoke(caches, "Create", SpreadsheetPivotParse.XlDatabase, sourceAddrA1, version12)),
                ("Create(R1C1)", () => EtCom.Invoke(caches, "Create", SpreadsheetPivotParse.XlDatabase, sourceAddrR1C1)),
                ("Create(R1C1,v12)", () => EtCom.Invoke(caches, "Create", SpreadsheetPivotParse.XlDatabase, sourceAddrR1C1, version12)),
                ("Add(Range)", () => EtCom.Invoke(caches, "Add", SpreadsheetPivotParse.XlDatabase, sourceRange)),
                ("Add(A1)", () => EtCom.Invoke(caches, "Add", SpreadsheetPivotParse.XlDatabase, sourceAddrA1)),
                ("Add(R1C1)", () => EtCom.Invoke(caches, "Add", SpreadsheetPivotParse.XlDatabase, sourceAddrR1C1)),
            };

            var errors = new List<string>();
            foreach (var attempt in attempts)
            {
                try
                {
                    object result = attempt.run();
                    if (result != null)
                    {
                        cache = result;
                        return true;
                    }

                    errors.Add(attempt.label + "=null");
                }
                catch (Exception ex)
                {
                    errors.Add(attempt.label + ": " + FormatPivotError(ex));
                }
            }

            error = "创建 PivotCache 失败（et）。" + string.Join(" | ", errors);
            return false;
        }

        private static bool TryCreateEtPivotTable(
            object cache,
            object destRange,
            string destAddrA1,
            string name,
            out object table,
            out string error)
        {
            table = null;
            error = null;
            var attempts = new List<(string label, Func<object> run)>();

            if (string.IsNullOrEmpty(name))
            {
                attempts.Add(("CreatePivotTable(Range)", () => EtCom.Invoke(cache, "CreatePivotTable", destRange)));
                attempts.Add(("CreatePivotTable(A1)", () => EtCom.Invoke(cache, "CreatePivotTable", destAddrA1)));
            }
            else
            {
                attempts.Add(("CreatePivotTable(Range,name)", () => EtCom.Invoke(cache, "CreatePivotTable", destRange, name)));
                attempts.Add(("CreatePivotTable(A1,name)", () => EtCom.Invoke(cache, "CreatePivotTable", destAddrA1, name)));
                attempts.Add(("CreatePivotTable(Range)", () => EtCom.Invoke(cache, "CreatePivotTable", destRange)));
            }

            var errors = new List<string>();
            foreach (var attempt in attempts)
            {
                try
                {
                    object result = attempt.run();
                    if (result != null)
                    {
                        table = result;
                        return true;
                    }

                    errors.Add(attempt.label + "=null");
                }
                catch (Exception ex)
                {
                    errors.Add(attempt.label + ": " + FormatPivotError(ex));
                }
            }

            error = "创建透视表失败（et）。" + string.Join(" | ", errors);
            return false;
        }

        private static bool TryResolveEtPivotField(
            object table,
            string requestedName,
            out object field,
            out string error)
        {
            field = null;
            error = null;
            string want = (requestedName ?? "").Trim();
            if (string.IsNullOrEmpty(want))
            {
                error = "字段名不能为空";
                return false;
            }

            try
            {
                field = EtCom.Invoke(table, "PivotFields", want);
                if (field != null)
                {
                    return true;
                }
            }
            catch (Exception)
            {
            }

            var available = new List<string>();
            try
            {
                object all = null;
                try
                {
                    all = EtCom.Invoke(table, "PivotFields", Type.Missing);
                }
                catch (Exception)
                {
                    all = EtCom.GetProperty(table, "PivotFields");
                }

                if (all != null)
                {
                    int count = Convert.ToInt32(EtCom.GetProperty(all, "Count"));
                    for (int i = 1; i <= count; i++)
                    {
                        object candidate = TryGetPivotItem(all, i);
                        if (candidate == null)
                        {
                            continue;
                        }

                        string name = "";
                        string caption = "";
                        string sourceName = "";
                        try
                        {
                            name = Convert.ToString(EtCom.GetProperty(candidate, "Name"))?.Trim() ?? "";
                        }
                        catch (Exception)
                        {
                        }

                        try
                        {
                            caption = Convert.ToString(EtCom.GetProperty(candidate, "Caption"))?.Trim() ?? "";
                        }
                        catch (Exception)
                        {
                        }

                        try
                        {
                            sourceName = Convert.ToString(EtCom.GetProperty(candidate, "SourceName"))?.Trim() ?? "";
                        }
                        catch (Exception)
                        {
                        }

                        if (!string.IsNullOrEmpty(name))
                        {
                            available.Add(name);
                        }

                        if (NamesMatch(want, name)
                            || NamesMatch(want, caption)
                            || NamesMatch(want, sourceName))
                        {
                            field = candidate;
                            return true;
                        }
                    }
                }
            }
            catch (Exception)
            {
            }

            error = "无法匹配透视字段: " + want;
            if (available.Count > 0)
            {
                error += "（可用: " + string.Join(", ", available) + "）";
            }

            return false;
        }

        private static void TryClearEtPivotOrphan(object table)
        {
            if (table == null)
            {
                return;
            }

            try
            {
                object tr = EtCom.GetProperty(table, "TableRange2")
                    ?? EtCom.GetProperty(table, "TableRange1");
                if (tr != null)
                {
                    EtCom.Invoke(tr, "Clear");
                }
            }
            catch (Exception)
            {
            }
        }

        private static bool TryValidateSourceEt(
            object sheet,
            object sourceRange,
            int firstRow,
            int firstCol,
            int lastRow,
            int lastCol,
            SpreadsheetPivotRequest request,
            out string error)
        {
            error = null;
            try
            {
                object merge = EtCom.GetProperty(sourceRange, "MergeCells");
                if (merge is bool merged)
                {
                    if (merged)
                    {
                        error = "源区域含合并格，首期不支持";
                        return false;
                    }
                }
                else if (merge != null)
                {
                    error = "源区域含合并格，首期不支持";
                    return false;
                }
            }
            catch (Exception)
            {
            }

            var headers = new HashSet<string>(StringComparer.Ordinal);
            object cells = EtCom.GetProperty(sheet, "Cells");
            for (int c = firstCol; c <= lastCol; c++)
            {
                object cell = EtCom.GetIndexed2(cells, firstRow, c);
                string header = ReadEtCellText(cell);
                if (string.IsNullOrWhiteSpace(header))
                {
                    error = "源第一行须为非空字段名（列 " + A1Address.ColumnLetter(c) + "）";
                    return false;
                }

                header = header.Trim();
                if (!headers.Add(header))
                {
                    error = "源表头字段名重复: " + header;
                    return false;
                }
            }

            if (lastRow <= firstRow)
            {
                error = "源区域除表头外至少需要一行数据";
                return false;
            }

            foreach (string f in request.Rows)
            {
                if (!headers.Contains(f))
                {
                    error = "rows 字段不在源表头: " + f;
                    return false;
                }
            }

            foreach (string f in request.Columns)
            {
                if (!headers.Contains(f))
                {
                    error = "columns 字段不在源表头: " + f;
                    return false;
                }
            }

            foreach (SpreadsheetPivotValueField v in request.Values)
            {
                if (!headers.Contains(v.Field))
                {
                    error = "values 字段不在源表头: " + v.Field;
                    return false;
                }
            }

            return true;
        }

        private static string ReadEtCellText(object cell)
        {
            if (cell == null)
            {
                return "";
            }

            try
            {
                object t = EtCom.GetProperty(cell, "Text");
                if (t != null)
                {
                    string s = Convert.ToString(t);
                    if (!string.IsNullOrWhiteSpace(s))
                    {
                        return s;
                    }
                }
            }
            catch (Exception)
            {
            }

            try
            {
                object v = EtCom.GetProperty(cell, "Value2");
                return v == null ? "" : Convert.ToString(v);
            }
            catch (Exception)
            {
                return "";
            }
        }

        private static bool TryRefreshOrDeleteEt(
            object book,
            SpreadsheetPivotRequest request,
            bool delete,
            SpreadsheetPivotResult partial,
            out string error)
        {
            if (!TryLocateEtPivot(book, request, out object sheet, out object pt, out error))
            {
                return false;
            }

            string sheetName = EtCom.TryReadName(sheet) ?? request.DestSheet ?? "";
            SpreadsheetPivotInfo info = ReadEtPivotInfo(sheetName, pt);
            if (delete)
            {
                try
                {
                    object tr = EtCom.GetProperty(pt, "TableRange2")
                        ?? EtCom.GetProperty(pt, "TableRange1");
                    EtCom.Invoke(tr, "Clear");
                }
                catch (Exception ex)
                {
                    error = "删除透视表失败: " + ex.Message;
                    return false;
                }

                FillResultFromInfo(partial, info, info?.SourceSheet, info?.SourceRange);
                return true;
            }

            try
            {
                object cache = EtCom.GetProperty(pt, "PivotCache");
                EtCom.Invoke(cache, "Refresh");
            }
            catch (Exception)
            {
                try
                {
                    EtCom.Invoke(pt, "RefreshTable");
                }
                catch (Exception ex)
                {
                    error = "刷新透视表失败: " + ex.Message;
                    return false;
                }
            }

            info = ReadEtPivotInfo(sheetName, pt);
            FillResultFromInfo(partial, info, info?.SourceSheet, info?.SourceRange);
            return true;
        }

        private static bool TryLocateEtPivot(
            object book,
            SpreadsheetPivotRequest request,
            out object sheet,
            out object pt,
            out string error)
        {
            sheet = null;
            pt = null;
            error = null;

            if (!string.IsNullOrEmpty(request.Name))
            {
                if (TryFindEtPivotByName(book, request.Name, out sheet, out pt))
                {
                    return true;
                }

                error = "未找到透视表: " + request.Name;
                return false;
            }

            if (!TryGetEtWorksheet(book, request.DestSheet, out sheet, out error))
            {
                return false;
            }

            if (!A1Address.TryParseCell(request.DestCellA1.Trim(), out int row, out int col))
            {
                error = "dest.cell 须为单格 A1（如 H1）";
                return false;
            }

            if (TryFindEtPivotCoveringCell(sheet, row, col, out pt, out _))
            {
                return true;
            }

            error = "未找到透视表（位置 " + request.DestSheet + "!" + A1Address.Cell(row, col) + "）";
            return false;
        }

        private static bool TryFindEtPivotByName(
            object book,
            string name,
            out object sheet,
            out object pt)
        {
            sheet = null;
            pt = null;
            foreach (object candidate in EtCom.EnumerateSheets(book))
            {
                object tables = TryGetPivotTables(candidate);
                if (tables == null)
                {
                    continue;
                }

                int count;
                try
                {
                    count = Convert.ToInt32(EtCom.GetProperty(tables, "Count"));
                }
                catch (Exception)
                {
                    continue;
                }

                for (int i = 1; i <= count; i++)
                {
                    object item = TryGetPivotItem(tables, i);
                    if (item == null)
                    {
                        continue;
                    }

                    string n = "";
                    try
                    {
                        n = Convert.ToString(EtCom.GetProperty(item, "Name")) ?? "";
                    }
                    catch (Exception)
                    {
                    }

                    if (string.Equals(n, name, StringComparison.Ordinal))
                    {
                        sheet = candidate;
                        pt = item;
                        return true;
                    }
                }
            }

            return false;
        }

        private static bool TryFindEtPivotCoveringCell(
            object sheet,
            int row,
            int col,
            out object pt,
            out string pivotName)
        {
            pt = null;
            pivotName = null;
            object tables = TryGetPivotTables(sheet);
            if (tables == null)
            {
                return false;
            }

            int count;
            try
            {
                count = Convert.ToInt32(EtCom.GetProperty(tables, "Count"));
            }
            catch (Exception)
            {
                return false;
            }

            for (int i = 1; i <= count; i++)
            {
                object candidate = TryGetPivotItem(tables, i);
                if (candidate == null)
                {
                    continue;
                }

                try
                {
                    object tr = EtCom.GetProperty(candidate, "TableRange1");
                    int r = Convert.ToInt32(EtCom.GetProperty(tr, "Row"));
                    int c = Convert.ToInt32(EtCom.GetProperty(tr, "Column"));
                    int rows = Convert.ToInt32(EtCom.GetProperty(EtCom.GetProperty(tr, "Rows"), "Count"));
                    int cols = Convert.ToInt32(EtCom.GetProperty(EtCom.GetProperty(tr, "Columns"), "Count"));
                    if (row >= r && row <= r + rows - 1 && col >= c && col <= c + cols - 1)
                    {
                        pt = candidate;
                        pivotName = Convert.ToString(EtCom.GetProperty(candidate, "Name")) ?? "";
                        return true;
                    }
                }
                catch (Exception)
                {
                }
            }

            return false;
        }

        private static bool TryGetEtWorksheet(
            object book,
            string sheetName,
            out object sheet,
            out string error)
        {
            sheet = null;
            error = null;
            try
            {
                foreach (object candidate in EtCom.EnumerateSheets(book))
                {
                    string name = EtCom.TryReadName(candidate) ?? "";
                    if (!string.Equals(name, sheetName, StringComparison.Ordinal))
                    {
                        continue;
                    }

                    try
                    {
                        object type = EtCom.GetProperty(candidate, "Type");
                        if (type != null && Convert.ToInt32(type) == -4109)
                        {
                            error = "sheet_type=chart，无法操作透视表";
                            return false;
                        }
                    }
                    catch (Exception)
                    {
                    }

                    sheet = candidate;
                    return true;
                }
            }
            catch (Exception ex)
            {
                error = "COM 不可用: " + ex.Message;
                return false;
            }

            error = "工作表不存在: " + sheetName;
            return false;
        }

        private static object TryGetPivotTables(object sheet)
        {
            try
            {
                return EtCom.Invoke(sheet, "PivotTables");
            }
            catch (Exception)
            {
            }

            try
            {
                return EtCom.GetProperty(sheet, "PivotTables");
            }
            catch (Exception)
            {
                return null;
            }
        }

        private static object TryGetPivotItem(object tables, int index)
        {
            try
            {
                return EtCom.GetIndexed(tables, index);
            }
            catch (Exception)
            {
            }

            try
            {
                return EtCom.Invoke(tables, "Item", index);
            }
            catch (Exception)
            {
                return null;
            }
        }

        private static void FillResultFromInfo(
            SpreadsheetPivotResult partial,
            SpreadsheetPivotInfo info,
            string sourceSheetFallback,
            string sourceRangeFallback)
        {
            if (info == null)
            {
                return;
            }

            partial.Name = info.Name;
            partial.Sheet = info.Sheet;
            partial.DestCell = info.DestCell;
            partial.SourceSheet = string.IsNullOrEmpty(info.SourceSheet)
                ? sourceSheetFallback
                : info.SourceSheet;
            partial.SourceRange = string.IsNullOrEmpty(info.SourceRange)
                ? sourceRangeFallback
                : info.SourceRange;
        }

        internal static void TryParseSourceData(
            string raw,
            out string sheet,
            out string range)
        {
            sheet = null;
            range = null;
            if (string.IsNullOrWhiteSpace(raw))
            {
                return;
            }

            string s = raw.Trim().Replace("$", "");
            if (s.StartsWith("'", StringComparison.Ordinal))
            {
                int endQuote = s.IndexOf('\'', 1);
                if (endQuote > 1 && endQuote + 1 < s.Length && s[endQuote + 1] == '!')
                {
                    sheet = s.Substring(1, endQuote - 1);
                    range = s.Substring(endQuote + 2);
                    return;
                }
            }

            int bang = s.LastIndexOf('!');
            if (bang >= 0)
            {
                sheet = s.Substring(0, bang).Trim('\'');
                range = s.Substring(bang + 1);
                return;
            }

            range = s;
        }
    }
}
