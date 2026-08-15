using System;
using System.Collections.Generic;
using Excel = Microsoft.Office.Interop.Excel;
using WordAddIn1.OpenFiles;

namespace WordAddIn1.SpreadsheetHost
{
    internal static class SpreadsheetPivotCom
    {
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
                error = "透视表操作失败: " + ex.Message;
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
                error = "透视表操作失败: " + ex.Message;
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
                sourceRange);
            object tableNameArg = string.IsNullOrEmpty(request.Name)
                ? Type.Missing
                : (object)request.Name;
            Excel.PivotTable table = cache.CreatePivotTable(destRange, tableNameArg);

            foreach (string rowField in request.Rows)
            {
                Excel.PivotField pf = table.PivotFields(rowField);
                pf.Orientation = Excel.XlPivotFieldOrientation.xlRowField;
            }

            foreach (string colField in request.Columns)
            {
                Excel.PivotField pf = table.PivotFields(colField);
                pf.Orientation = Excel.XlPivotFieldOrientation.xlColumnField;
            }

            foreach (SpreadsheetPivotValueField value in request.Values)
            {
                if (!SpreadsheetPivotParse.TryMapAgg(value.Agg, out int xlFn, out error))
                {
                    return false;
                }

                table.AddDataField(
                    table.PivotFields(value.Field),
                    Type.Missing,
                    (Excel.XlConsolidationFunction)xlFn);
            }

            SpreadsheetPivotInfo info = ReadExcelPivotInfo(destSheet.Name, table);
            FillResultFromInfo(partial, info, request.SourceSheet, actualSource);
            return true;
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
                else
                {
                    // DBNull 等：部分合并
                    error = "源区域含合并格，首期不支持";
                    return false;
                }
            }
            catch (Exception)
            {
            }

            var headers = new HashSet<string>(StringComparer.Ordinal);
            for (int c = firstCol; c <= lastCol; c++)
            {
                Excel.Range cell = sourceRange.Worksheet.Cells[firstRow, c];
                string header = ReadCellText(cell);
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
            object sourceRange = EtCom.Invoke(sourceSheet, "Range", actualSource);
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
            object destRange = EtCom.Invoke(destSheet, "Range", destCell);
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

            object caches = EtCom.GetProperty(book, "PivotCaches");
            if (caches == null)
            {
                error = "当前 et 不支持 PivotCaches";
                return false;
            }

            object cache;
            try
            {
                cache = EtCom.Invoke(caches, "Create", SpreadsheetPivotParse.XlDatabase, sourceRange);
            }
            catch (Exception ex)
            {
                error = "创建 PivotCache 失败: " + ex.Message;
                return false;
            }

            if (cache == null)
            {
                error = "创建 PivotCache 失败";
                return false;
            }

            object table;
            try
            {
                if (string.IsNullOrEmpty(request.Name))
                {
                    table = EtCom.Invoke(cache, "CreatePivotTable", destRange);
                }
                else
                {
                    table = EtCom.Invoke(cache, "CreatePivotTable", destRange, request.Name);
                }
            }
            catch (Exception ex)
            {
                error = "创建透视表失败: " + ex.Message;
                return false;
            }

            if (table == null)
            {
                error = "创建透视表失败";
                return false;
            }

            foreach (string rowField in request.Rows)
            {
                object pf = EtCom.Invoke(table, "PivotFields", rowField);
                EtCom.TrySetProperty(pf, "Orientation", SpreadsheetPivotParse.XlRowField);
            }

            foreach (string colField in request.Columns)
            {
                object pf = EtCom.Invoke(table, "PivotFields", colField);
                EtCom.TrySetProperty(pf, "Orientation", SpreadsheetPivotParse.XlColumnField);
            }

            foreach (SpreadsheetPivotValueField value in request.Values)
            {
                if (!SpreadsheetPivotParse.TryMapAgg(value.Agg, out int xlFn, out error))
                {
                    return false;
                }

                object pf = EtCom.Invoke(table, "PivotFields", value.Field);
                try
                {
                    EtCom.Invoke(table, "AddDataField", pf, Type.Missing, xlFn);
                }
                catch (Exception)
                {
                    EtCom.TrySetProperty(pf, "Orientation", SpreadsheetPivotParse.XlDataField);
                    EtCom.TrySetProperty(pf, "Function", xlFn);
                }
            }

            string destSheetName = EtCom.TryReadName(destSheet) ?? request.DestSheet;
            SpreadsheetPivotInfo info = ReadEtPivotInfo(destSheetName, table);
            FillResultFromInfo(partial, info, request.SourceSheet, actualSource);
            return true;
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
