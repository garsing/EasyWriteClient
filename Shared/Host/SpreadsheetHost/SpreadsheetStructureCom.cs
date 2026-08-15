using System;
using System.Reflection;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using WordAddIn1.OpenFiles;

namespace WordAddIn1.SpreadsheetHost
{
    internal static class SpreadsheetStructureCom
    {
        private static string FormatError(Exception ex)
        {
            Exception cur = ex;
            while (cur is TargetInvocationException tie && tie.InnerException != null)
            {
                cur = tie.InnerException;
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
            SpreadsheetStructureRequest request,
            out SpreadsheetStructureResult partial,
            out string error)
        {
            partial = new SpreadsheetStructureResult
            {
                Action = request.Action,
                Range = request.RangeA1,
                Count = request.Count
            };
            error = null;

            try
            {
                if (!TryGetExcelWorksheet(book, request.SheetName, out Excel.Worksheet sheet, out error))
                {
                    return false;
                }

                partial.Sheet = sheet.Name;
                bool ok;
                switch (request.Action)
                {
                    case "merge":
                        ok = TryMergeExcel(sheet, request, partial, out error);
                        break;
                    case "unmerge":
                        ok = TryUnmergeExcel(sheet, request, partial, out error);
                        break;
                    case "insert_rows":
                        ok = TryInsertDeleteRowsExcel(sheet, request, insert: true, partial, out error);
                        break;
                    case "delete_rows":
                        ok = TryInsertDeleteRowsExcel(sheet, request, insert: false, partial, out error);
                        break;
                    case "insert_cols":
                        ok = TryInsertDeleteColsExcel(sheet, request, insert: true, partial, out error);
                        break;
                    case "delete_cols":
                        ok = TryInsertDeleteColsExcel(sheet, request, insert: false, partial, out error);
                        break;
                    default:
                        error = "未知 action";
                        return false;
                }

                return ok;
            }
            catch (Exception ex)
            {
                error = "结构操作失败: " + FormatError(ex);
                return false;
            }
        }

        public static bool TryExecuteOnEtWorkbook(
            object book,
            SpreadsheetStructureRequest request,
            out SpreadsheetStructureResult partial,
            out string error)
        {
            partial = new SpreadsheetStructureResult
            {
                Action = request.Action,
                Range = request.RangeA1,
                Count = request.Count
            };
            error = null;

            try
            {
                if (!TryGetEtWorksheet(book, request.SheetName, out object sheet, out error))
                {
                    return false;
                }

                partial.Sheet = EtCom.TryReadName(sheet) ?? request.SheetName;
                bool ok;
                switch (request.Action)
                {
                    case "merge":
                        ok = TryMergeEt(sheet, request, partial, out error);
                        break;
                    case "unmerge":
                        ok = TryUnmergeEt(sheet, request, partial, out error);
                        break;
                    case "insert_rows":
                        ok = TryInsertDeleteRowsEt(book, sheet, request, insert: true, partial, out error);
                        break;
                    case "delete_rows":
                        ok = TryInsertDeleteRowsEt(book, sheet, request, insert: false, partial, out error);
                        break;
                    case "insert_cols":
                        ok = TryInsertDeleteColsEt(book, sheet, request, insert: true, partial, out error);
                        break;
                    case "delete_cols":
                        ok = TryInsertDeleteColsEt(book, sheet, request, insert: false, partial, out error);
                        break;
                    default:
                        error = "未知 action";
                        return false;
                }

                return ok;
            }
            catch (Exception ex)
            {
                error = "结构操作失败: " + FormatError(ex);
                return false;
            }
        }

        private static bool TryMergeExcel(
            Excel.Worksheet sheet,
            SpreadsheetStructureRequest request,
            SpreadsheetStructureResult partial,
            out string error)
        {
            error = null;
            Excel.Range r = sheet.Range[request.RangeA1];
            r.Merge();
            partial.Affected = "merge_area=" + request.RangeA1;
            return true;
        }

        private static bool TryUnmergeExcel(
            Excel.Worksheet sheet,
            SpreadsheetStructureRequest request,
            SpreadsheetStructureResult partial,
            out string error)
        {
            error = null;
            Excel.Range r = sheet.Range[request.RangeA1];
            try
            {
                if (r.MergeCells)
                {
                    r = r.MergeArea;
                }
            }
            catch (Exception)
            {
            }

            string area;
            try
            {
                area = r.Address[false, false];
            }
            catch (Exception)
            {
                area = request.RangeA1;
            }

            r.UnMerge();
            partial.Affected = "merge_area=" + area;
            return true;
        }

        private static bool TryInsertDeleteRowsExcel(
            Excel.Worksheet sheet,
            SpreadsheetStructureRequest request,
            bool insert,
            SpreadsheetStructureResult partial,
            out string error)
        {
            error = null;
            if (!A1Address.TryParseCell(request.RangeA1, out int row, out _))
            {
                error = "插删的 range 须为单格 A1";
                return false;
            }

            int count = request.Count;
            Excel.Application app = sheet.Application;
            bool? oldAlerts = null;
            try
            {
                oldAlerts = app.DisplayAlerts;
                app.DisplayAlerts = false;
            }
            catch (Exception)
            {
            }

            try
            {
                Excel.Range band = sheet.Rows[row];
                if (count > 1)
                {
                    band = band.Resize[count];
                }

                if (insert)
                {
                    band.Insert(Excel.XlInsertShiftDirection.xlShiftDown);
                }
                else
                {
                    band.Delete(Excel.XlDeleteShiftDirection.xlShiftUp);
                }

                partial.Affected = "rows=" + row + ":" + (row + count - 1);
                return true;
            }
            finally
            {
                if (oldAlerts.HasValue)
                {
                    try
                    {
                        app.DisplayAlerts = oldAlerts.Value;
                    }
                    catch (Exception)
                    {
                    }
                }
            }
        }

        private static bool TryInsertDeleteColsExcel(
            Excel.Worksheet sheet,
            SpreadsheetStructureRequest request,
            bool insert,
            SpreadsheetStructureResult partial,
            out string error)
        {
            error = null;
            if (!A1Address.TryParseCell(request.RangeA1, out _, out int col))
            {
                error = "插删的 range 须为单格 A1";
                return false;
            }

            int count = request.Count;
            Excel.Application app = sheet.Application;
            bool? oldAlerts = null;
            try
            {
                oldAlerts = app.DisplayAlerts;
                app.DisplayAlerts = false;
            }
            catch (Exception)
            {
            }

            try
            {
                Excel.Range band = sheet.Columns[col];
                if (count > 1)
                {
                    band = band.Resize[Type.Missing, count];
                }

                if (insert)
                {
                    band.Insert(Excel.XlInsertShiftDirection.xlShiftToRight);
                }
                else
                {
                    band.Delete(Excel.XlDeleteShiftDirection.xlShiftToLeft);
                }

                string c1 = A1Address.ColumnLetter(col);
                string c2 = A1Address.ColumnLetter(col + count - 1);
                partial.Affected = "cols=" + c1 + ":" + c2;
                return true;
            }
            finally
            {
                if (oldAlerts.HasValue)
                {
                    try
                    {
                        app.DisplayAlerts = oldAlerts.Value;
                    }
                    catch (Exception)
                    {
                    }
                }
            }
        }

        private static bool TryMergeEt(
            object sheet,
            SpreadsheetStructureRequest request,
            SpreadsheetStructureResult partial,
            out string error)
        {
            error = null;
            object r = EtGetSheetRange(sheet, request.RangeA1);
            if (r == null)
            {
                error = "无法解析 range";
                return false;
            }

            EtCom.InvokeFlex(r, "Merge");
            partial.Affected = "merge_area=" + request.RangeA1;
            return true;
        }

        private static bool TryUnmergeEt(
            object sheet,
            SpreadsheetStructureRequest request,
            SpreadsheetStructureResult partial,
            out string error)
        {
            error = null;
            object r = EtGetSheetRange(sheet, request.RangeA1);
            if (r == null)
            {
                error = "无法解析 range";
                return false;
            }

            string area = request.RangeA1;
            try
            {
                object mergeCells = EtCom.GetProperty(r, "MergeCells");
                if (mergeCells != null && Convert.ToBoolean(mergeCells))
                {
                    object mergeArea = EtCom.GetProperty(r, "MergeArea");
                    if (mergeArea != null)
                    {
                        r = mergeArea;
                        try
                        {
                            area = Convert.ToString(EtCom.Invoke(mergeArea, "Address", false, false)) ?? area;
                        }
                        catch (Exception)
                        {
                        }
                    }
                }
            }
            catch (Exception)
            {
            }

            EtCom.InvokeFlex(r, "UnMerge");
            partial.Affected = "merge_area=" + area;
            return true;
        }

        private static bool TryInsertDeleteRowsEt(
            object book,
            object sheet,
            SpreadsheetStructureRequest request,
            bool insert,
            SpreadsheetStructureResult partial,
            out string error)
        {
            error = null;
            if (!A1Address.TryParseCell(request.RangeA1, out int row, out _))
            {
                error = "插删的 range 须为单格 A1";
                return false;
            }

            int count = request.Count;
            string rowSpec = count == 1
                ? row.ToString()
                : row + ":" + (row + count - 1);

            object withAlerts = TryGetApplication(book, sheet);
            object oldAlerts = null;
            try
            {
                if (withAlerts != null)
                {
                    oldAlerts = EtCom.GetProperty(withAlerts, "DisplayAlerts");
                    EtCom.TrySetProperty(withAlerts, "DisplayAlerts", false);
                }

                object rows = EtCom.GetProperty(sheet, "Rows") ?? EtCom.InvokeFlex(sheet, "Rows");
                object band = null;
                try
                {
                    band = EtCom.InvokeFlex(rows, "Item", rowSpec);
                }
                catch (Exception)
                {
                }

                if (band == null)
                {
                    try
                    {
                        band = EtCom.GetIndexed(rows, row);
                        if (band != null && count > 1)
                        {
                            band = EtCom.InvokeFlex(band, "Resize", count);
                        }
                    }
                    catch (Exception)
                    {
                    }
                }

                if (band == null)
                {
                    error = "无法定位行: " + rowSpec;
                    return false;
                }

                if (insert)
                {
                    EtCom.InvokeFlex(band, "Insert");
                }
                else
                {
                    EtCom.InvokeFlex(band, "Delete");
                }

                partial.Affected = "rows=" + row + ":" + (row + count - 1);
                return true;
            }
            finally
            {
                if (withAlerts != null && oldAlerts != null)
                {
                    try
                    {
                        EtCom.TrySetProperty(withAlerts, "DisplayAlerts", oldAlerts);
                    }
                    catch (Exception)
                    {
                    }
                }
            }
        }

        private static bool TryInsertDeleteColsEt(
            object book,
            object sheet,
            SpreadsheetStructureRequest request,
            bool insert,
            SpreadsheetStructureResult partial,
            out string error)
        {
            error = null;
            if (!A1Address.TryParseCell(request.RangeA1, out _, out int col))
            {
                error = "插删的 range 须为单格 A1";
                return false;
            }

            int count = request.Count;
            string c1 = A1Address.ColumnLetter(col);
            string c2 = A1Address.ColumnLetter(col + count - 1);
            string colSpec = count == 1 ? c1 : c1 + ":" + c2;

            object withAlerts = TryGetApplication(book, sheet);
            object oldAlerts = null;
            try
            {
                if (withAlerts != null)
                {
                    oldAlerts = EtCom.GetProperty(withAlerts, "DisplayAlerts");
                    EtCom.TrySetProperty(withAlerts, "DisplayAlerts", false);
                }

                object cols = EtCom.GetProperty(sheet, "Columns") ?? EtCom.InvokeFlex(sheet, "Columns");
                object band = null;
                try
                {
                    band = EtCom.InvokeFlex(cols, "Item", colSpec);
                }
                catch (Exception)
                {
                }

                if (band == null)
                {
                    try
                    {
                        band = EtCom.GetIndexed(cols, col);
                        if (band != null && count > 1)
                        {
                            band = EtCom.InvokeFlex(band, "Resize", Type.Missing, count);
                        }
                    }
                    catch (Exception)
                    {
                    }
                }

                if (band == null)
                {
                    error = "无法定位列: " + colSpec;
                    return false;
                }

                if (insert)
                {
                    EtCom.InvokeFlex(band, "Insert");
                }
                else
                {
                    EtCom.InvokeFlex(band, "Delete");
                }

                partial.Affected = "cols=" + c1 + ":" + c2;
                return true;
            }
            finally
            {
                if (withAlerts != null && oldAlerts != null)
                {
                    try
                    {
                        EtCom.TrySetProperty(withAlerts, "DisplayAlerts", oldAlerts);
                    }
                    catch (Exception)
                    {
                    }
                }
            }
        }

        private static object TryGetApplication(object book, object sheet)
        {
            try
            {
                object app = EtCom.GetProperty(sheet, "Application");
                if (app != null)
                {
                    return app;
                }
            }
            catch (Exception)
            {
            }

            try
            {
                return EtCom.GetProperty(book, "Application");
            }
            catch (Exception)
            {
                return null;
            }
        }

        private static bool TryGetExcelWorksheet(
            Excel.Workbook book,
            string sheetName,
            out Excel.Worksheet sheet,
            out string error)
        {
            sheet = null;
            error = null;
            if (string.IsNullOrWhiteSpace(sheetName))
            {
                error = "须提供工作表名";
                return false;
            }

            foreach (object raw in book.Sheets)
            {
                string name;
                try
                {
                    name = ((dynamic)raw).Name;
                }
                catch (Exception)
                {
                    continue;
                }

                if (!string.Equals(name, sheetName.Trim(), StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                if (raw is Excel.Chart)
                {
                    error = "sheet_type=chart，无法操作格子结构";
                    return false;
                }

                sheet = raw as Excel.Worksheet;
                if (sheet == null)
                {
                    error = "sheet_type=chart，无法操作格子结构";
                    return false;
                }

                return true;
            }

            error = "未找到工作表: " + sheetName;
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
            if (string.IsNullOrWhiteSpace(sheetName))
            {
                error = "须提供工作表名";
                return false;
            }

            foreach (object candidate in EtCom.EnumerateSheets(book))
            {
                string name = EtCom.TryReadName(candidate) ?? "";
                if (!string.Equals(name, sheetName.Trim(), StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                try
                {
                    object typeObj = EtCom.GetProperty(candidate, "Type");
                    if (typeObj != null && Convert.ToInt32(typeObj) == -4109)
                    {
                        error = "sheet_type=chart，无法操作格子结构";
                        return false;
                    }
                }
                catch (Exception)
                {
                }

                sheet = candidate;
                return true;
            }

            error = "未找到工作表: " + sheetName;
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
                    BindingFlags.GetProperty
                        | BindingFlags.InvokeMethod
                        | BindingFlags.Instance
                        | BindingFlags.Public,
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

            return null;
        }
    }
}
