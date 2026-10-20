using System;
using System.Collections.Generic;
using System.IO;
using System.Xml.Linq;
using WordAddIn1.OpenFiles;

namespace WordAddIn1.ChartOoxmlHost
{
    internal static partial class SpreadsheetChartPackageHost
    {
        public static bool TryApplyChromeEt(
            object book,
            object sheet,
            object chartHost,
            ChartInfo format,
            out string error)
        {
            error = null;
            string fullBefore = EtCom.TryReadFullName(book);
            if (!WordAddIn1.SpreadsheetHost.SpreadsheetContentUtil.IsSavedPath(fullBefore))
            {
                error = "请先保存工作簿后再画图";
                return false;
            }

            string sheetName = EtCom.TryReadName(sheet) ?? "";
            int index = ChartObjectIndexEt(sheet, chartHost);
            if (index < 0)
            {
                error = "找不到要改的图";
                return false;
            }

            string temp = TempPath();
            try
            {
                if (!SaveCopyEt(book, fullBefore, temp, out error))
                {
                    return false;
                }

                if (!TryRewriteChrome(temp, sheetName, index, format, null, null, out string snapshot, out error))
                {
                    return false;
                }

                if (!PasteFromCopyEt(book, sheet, chartHost, temp, index, out error))
                {
                    return false;
                }

                if (!string.Equals(EtCom.TryReadFullName(book), fullBefore, StringComparison.OrdinalIgnoreCase))
                {
                    error = "请先保存工作簿后再画图";
                    return false;
                }

                if (TryVerifyEt(book, sheetName, index, snapshot, temp, null, out _))
                {
                    return true;
                }

                string pinTemp = TempPath();
                try
                {
                    if (!SaveCopyEt(book, fullBefore, pinTemp, out error))
                    {
                        return false;
                    }

                    List<string> pinColors = ChartPartXml.SnapshotSeriesColors(XDocument.Parse(snapshot));
                    if (!TryRewriteChrome(pinTemp, sheetName, index, format, Path.GetFileName(fullBefore), pinColors, out snapshot, out error))
                    {
                        return false;
                    }

                    object current = ChartObjectAtEt(sheet, index);
                    if (current == null || !PasteFromCopyEt(book, sheet, current, pinTemp, index, out error))
                    {
                        error = error ?? "二次贴回失败";
                        return false;
                    }

                    if (!TryVerifyEt(book, sheetName, index, snapshot, pinTemp, Path.GetFileName(fullBefore), out string verifyError))
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
                error = "ET 改包失败: " + ex.Message;
                return false;
            }
            finally
            {
                TryDelete(temp);
            }
        }

        public static bool TryReadFormatEt(
            object book,
            object sheet,
            object chartHost,
            out ChartInfo format,
            out string error)
        {
            format = null;
            error = null;
            string fullBefore = EtCom.TryReadFullName(book);
            if (!WordAddIn1.SpreadsheetHost.SpreadsheetContentUtil.IsSavedPath(fullBefore))
            {
                error = "请先保存工作簿后再画图";
                return false;
            }

            int index = ChartObjectIndexEt(sheet, chartHost);
            if (index < 0)
            {
                error = "找不到要读的图";
                return false;
            }

            string temp = TempPath();
            try
            {
                if (!SaveCopyEt(book, fullBefore, temp, out error))
                {
                    return false;
                }

                return TryReadFormat(temp, EtCom.TryReadName(sheet) ?? "", index, out format, out error);
            }
            finally
            {
                TryDelete(temp);
            }
        }

        private static bool SaveCopyEt(object book, string fullBefore, string temp, out string error)
        {
            error = null;
            try
            {
                EtCom.Invoke(book, "SaveCopyAs", temp);
            }
            catch (Exception)
            {
                try
                {
                    EtCom.Invoke(book, "Save");
                    if (!File.Exists(fullBefore))
                    {
                        error = "请先保存工作簿后再画图";
                        return false;
                    }

                    File.Copy(fullBefore, temp, true);
                }
                catch (Exception ex)
                {
                    error = "请先保存工作簿后再画图: " + ex.Message;
                    return false;
                }
            }

            if (!string.Equals(EtCom.TryReadFullName(book), fullBefore, StringComparison.OrdinalIgnoreCase))
            {
                error = "请先保存工作簿后再画图";
                return false;
            }

            return File.Exists(temp);
        }
    }
}
