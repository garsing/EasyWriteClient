using System;
using System.Collections.Generic;
using System.IO;
using System.Xml.Linq;
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

                if (!TryRewriteChrome(temp, sheetName, index, format, pinFormulasToFile: null, pinColors: null, out string snapshot, out error))
                {
                    return false;
                }

                if (!PasteFromCopyExcel(book, sheet, chart, temp, index, out error))
                {
                    return false;
                }

                PinLiveChartStyleExcel(ChartObjectAt(sheet, index), format);

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
                    List<string> pinColors = ChartPartXml.SnapshotSeriesColors(XDocument.Parse(snapshot));
                    if (!TryRewriteChrome(pinTemp, sheetName, index, format, Path.GetFileName(fullBefore), pinColors, out snapshot, out error))
                    {
                        return false;
                    }

                    Excel.ChartObject current = ChartObjectAt(sheet, index);
                    if (current == null || !PasteFromCopyExcel(book, sheet, current, pinTemp, index, out error))
                    {
                        error = error ?? "二次贴回失败";
                        return false;
                    }

                    PinLiveChartStyleExcel(ChartObjectAt(sheet, index), format);

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

        private static void PinLiveChartStyleExcel(Excel.ChartObject co, ChartInfo format)
        {
            if (co == null || format == null)
            {
                return;
            }

            Excel.Chart chart = co.Chart;
            if (format.ChartStyle > 0)
            {
                try
                {
                    chart.ChartStyle = format.ChartStyle;
                }
                catch (Exception)
                {
                }
            }

            if (!string.IsNullOrWhiteSpace(format.Gridlines))
            {
                try
                {
                    Excel.Axis axis = (Excel.Axis)chart.Axes(Excel.XlAxisType.xlValue, Excel.XlAxisGroup.xlPrimary);
                    axis.HasMajorGridlines = string.Equals(format.Gridlines, "true", StringComparison.OrdinalIgnoreCase);
                    axis.HasMinorGridlines = false;
                }
                catch (Exception)
                {
                }
            }

            if (format.HasOverlap)
            {
                try
                {
                    Excel.ChartGroup group = (Excel.ChartGroup)chart.ChartGroups(1);
                    group.Overlap = format.Overlap;
                }
                catch (Exception)
                {
                }
            }

            try
            {
                    bool pinLabels = format.HasShowDataLabels
                    || format.ShowValue.HasValue
                    || format.ShowPercentage.HasValue;
                if (!pinLabels && format.SeriesCollection != null)
                {
                    for (int i = 0; i < format.SeriesCollection.Count; i++)
                    {
                        SeriesInfo series = format.SeriesCollection[i];
                        if (series != null
                            && (series.ShowDataLabels.HasValue
                                || series.ShowValue.HasValue
                                || series.ShowPercentage.HasValue))
                        {
                            pinLabels = true;
                            break;
                        }
                    }
                }

                if (!pinLabels)
                {
                    return;
                }

                Excel.SeriesCollection sc = (Excel.SeriesCollection)chart.SeriesCollection();
                for (int i = 1; i <= sc.Count; i++)
                {
                    Excel.Series s = sc.Item(i);
                    SeriesInfo series = null;
                    if (format.SeriesCollection != null && i - 1 < format.SeriesCollection.Count)
                    {
                        series = format.SeriesCollection[i - 1];
                    }

                    bool on = format.ShowDataLabels || format.ShowValue == true || format.ShowPercentage == true;
                    if (series != null
                        && (series.ShowDataLabels.HasValue
                            || series.ShowValue.HasValue
                            || series.ShowPercentage.HasValue))
                    {
                        on = series.ShowDataLabels == true
                            || series.ShowValue == true
                            || series.ShowPercentage == true;
                    }

                    s.HasDataLabels = on;
                    if (!on)
                    {
                        continue;
                    }

                    bool? showVal = series != null && series.ShowValue.HasValue
                        ? series.ShowValue
                        : format.ShowValue;
                    bool? showPct = series != null && series.ShowPercentage.HasValue
                        ? series.ShowPercentage
                        : format.ShowPercentage;
                    if (!showVal.HasValue && !showPct.HasValue)
                    {
                        continue;
                    }

                    Excel.DataLabels labels = (Excel.DataLabels)s.DataLabels();
                    labels.ShowValue = showVal ?? false;
                    labels.ShowPercentage = showPct ?? false;
                }
            }
            catch (Exception)
            {
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
