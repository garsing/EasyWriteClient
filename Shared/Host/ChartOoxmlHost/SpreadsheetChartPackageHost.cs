using System;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;

namespace WordAddIn1.ChartOoxmlHost
{
    /// <summary>
    /// Excel / ET 活图：另存副本，只改 chart 部件的皮，再把图贴回原表。
    /// 不关用户工作簿，不改 FullName。
    /// Excel 入口在本文件；ET、改包、贴回在同名分部。
    /// </summary>
    internal static partial class SpreadsheetChartPackageHost
    {
        public static bool TryApplyChromeExcel(
            Excel.Workbook book,
            Excel.Worksheet sheet,
            Excel.ChartObject chart,
            ChartInfo format,
            out string error)
        {
            error = null;
            if (book == null || sheet == null || chart == null || format == null)
            {
                error = "改包参数空";
                return false;
            }

            string fullBefore = book.FullName;
            if (!WordAddIn1.SpreadsheetHost.SpreadsheetContentUtil.IsSavedPath(fullBefore) || !File.Exists(fullBefore))
            {
                error = "请先保存工作簿后再画图";
                return false;
            }

            string sheetName = sheet.Name ?? "";
            int index = ChartObjectIndex(sheet, chart);
            if (index < 0)
            {
                error = "找不到要改的图";
                return false;
            }

            string temp = TempPath();
            Excel.Application app = book.Application;
            bool alerts = true;
            try
            {
                alerts = app.DisplayAlerts;
            }
            catch (Exception)
            {
            }

            try
            {
                try
                {
                    app.DisplayAlerts = false;
                }
                catch (Exception)
                {
                }

                book.SaveCopyAs(temp);
                if (!SameFullName(book, fullBefore))
                {
                    error = "请先保存工作簿后再画图";
                    return false;
                }

                if (!TryRewriteChrome(temp, sheetName, index, format, pinFormulasToFile: null, out string snapshot, out error))
                {
                    return false;
                }

                if (!PasteFromCopyExcel(book, sheet, chart, temp, index, out error))
                {
                    return false;
                }

                if (!SameFullName(book, fullBefore))
                {
                    error = "请先保存工作簿后再画图";
                    return false;
                }

                if (TryVerify(book, sheetName, index, snapshot, temp, null, out string verifyError))
                {
                    return true;
                }

                string pinTemp = TempPath();
                try
                {
                    book.SaveCopyAs(pinTemp);
                    if (!TryRewriteChrome(pinTemp, sheetName, index, format, Path.GetFileName(fullBefore), out snapshot, out error))
                    {
                        return false;
                    }

                    Excel.ChartObject current = ChartObjectAt(sheet, index);
                    if (current == null || !PasteFromCopyExcel(book, sheet, current, pinTemp, index, out error))
                    {
                        error = error ?? "二次贴回失败";
                        return false;
                    }

                    if (!TryVerify(book, sheetName, index, snapshot, pinTemp, Path.GetFileName(fullBefore), out verifyError))
                    {
                        error = "贴回后公式或外观没有钉住: " + verifyError;
                        return false;
                    }

                    return true;
                }
                finally
                {
                    TryDelete(pinTemp);
                }
            }
            catch (Exception ex)
            {
                error = "Excel 改包失败: " + ex.Message;
                return false;
            }
            finally
            {
                try
                {
                    app.DisplayAlerts = alerts;
                }
                catch (Exception)
                {
                }

                TryDelete(temp);
            }
        }

        public static bool TryReadFormatExcel(
            Excel.Workbook book,
            Excel.Worksheet sheet,
            Excel.ChartObject chart,
            out ChartInfo format,
            out string error)
        {
            format = null;
            error = null;
            string fullBefore = book.FullName;
            if (!WordAddIn1.SpreadsheetHost.SpreadsheetContentUtil.IsSavedPath(fullBefore) || !File.Exists(fullBefore))
            {
                error = "请先保存工作簿后再画图";
                return false;
            }

            int index = ChartObjectIndex(sheet, chart);
            if (index < 0)
            {
                error = "找不到要读的图";
                return false;
            }

            string temp = TempPath();
            try
            {
                book.SaveCopyAs(temp);
                if (!SameFullName(book, fullBefore))
                {
                    error = "请先保存工作簿后再画图";
                    return false;
                }

                return TryReadFormat(temp, sheet.Name ?? "", index, out format, out error);
            }
            catch (Exception ex)
            {
                error = "Excel 读包失败: " + ex.Message;
                return false;
            }
            finally
            {
                TryDelete(temp);
            }
        }

        private static bool SameFullName(Excel.Workbook book, string fullBefore)
        {
            try
            {
                return string.Equals(book.FullName, fullBefore, StringComparison.OrdinalIgnoreCase);
            }
            catch (Exception)
            {
                return false;
            }
        }

        private static string TempPath()
        {
            return Path.Combine(Path.GetTempPath(), "ew-xlchart-" + Guid.NewGuid().ToString("N") + ".xlsx");
        }

        private static void TryDelete(string path)
        {
            try
            {
                if (!string.IsNullOrEmpty(path) && File.Exists(path))
                {
                    File.Delete(path);
                }
            }
            catch (Exception)
            {
            }
        }
    }
}
