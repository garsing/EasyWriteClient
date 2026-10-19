using System;
using System.Collections.Generic;
using Excel = Microsoft.Office.Interop.Excel;
using WordAddIn1.OpenFiles;

namespace WordAddIn1.ChartOoxmlHost
{
    internal static partial class SpreadsheetChartPackageHost
    {
        private static bool PasteFromCopyExcel(
            Excel.Workbook book,
            Excel.Worksheet sheet,
            Excel.ChartObject target,
            string copyPath,
            int index,
            out string error)
        {
            error = null;
            Excel.Workbook copy = null;
            double left = target.Left;
            double top = target.Top;
            double width = target.Width;
            double height = target.Height;
            string name = target.Name;
            try
            {
                copy = book.Application.Workbooks.Open(copyPath);
                Excel.Worksheet copySheet = FindSheet(copy, sheet.Name);
                Excel.ChartObject source = ChartObjectAt(copySheet, index);
                if (source == null)
                {
                    error = "副本里没有这张图";
                    return false;
                }

                var before = ChartNames(sheet);
                source.Copy();
                target.Delete();
                before.Remove(name);
                sheet.Paste();
                Excel.ChartObject pasted = FindNewChart(sheet, before);
                if (pasted == null)
                {
                    error = "贴回后找不到新图";
                    return false;
                }

                pasted.Left = left;
                pasted.Top = top;
                pasted.Width = width;
                pasted.Height = height;
                try
                {
                    pasted.Name = name;
                }
                catch (Exception)
                {
                }

                return true;
            }
            catch (Exception ex)
            {
                error = "贴回失败: " + ex.Message;
                return false;
            }
            finally
            {
                if (copy != null)
                {
                    try
                    {
                        copy.Close(false);
                    }
                    catch (Exception)
                    {
                    }
                }
            }
        }

        private static bool PasteFromCopyEt(
            object book,
            object sheet,
            object target,
            string copyPath,
            int index,
            out string error)
        {
            error = null;
            object copy = null;
            double left = Convert.ToDouble(EtCom.GetProperty(target, "Left"));
            double top = Convert.ToDouble(EtCom.GetProperty(target, "Top"));
            double width = Convert.ToDouble(EtCom.GetProperty(target, "Width"));
            double height = Convert.ToDouble(EtCom.GetProperty(target, "Height"));
            string name = Convert.ToString(EtCom.GetProperty(target, "Name")) ?? "";
            try
            {
                object app = EtCom.GetProperty(book, "Application");
                object books = EtCom.GetProperty(app, "Workbooks");
                copy = EtCom.Invoke(books, "Open", copyPath);
                object copySheet = FindSheetEt(copy, EtCom.TryReadName(sheet));
                object source = ChartObjectAtEt(copySheet, index);
                if (source == null)
                {
                    error = "副本里没有这张图";
                    return false;
                }

                var before = ChartNamesEt(sheet);
                EtCom.Invoke(source, "Copy");
                EtCom.Invoke(target, "Delete");
                before.Remove(name);
                EtCom.Invoke(sheet, "Paste");
                object pasted = FindNewChartEt(sheet, before);
                if (pasted == null)
                {
                    error = "贴回后找不到新图";
                    return false;
                }

                EtCom.TrySetProperty(pasted, "Left", left);
                EtCom.TrySetProperty(pasted, "Top", top);
                EtCom.TrySetProperty(pasted, "Width", width);
                EtCom.TrySetProperty(pasted, "Height", height);
                EtCom.TrySetProperty(pasted, "Name", name);
                return true;
            }
            catch (Exception ex)
            {
                error = "贴回失败: " + ex.Message;
                return false;
            }
            finally
            {
                if (copy != null)
                {
                    try
                    {
                        EtCom.Invoke(copy, "Close", false);
                    }
                    catch (Exception)
                    {
                    }
                }
            }
        }

        private static int ChartObjectIndex(Excel.Worksheet sheet, Excel.ChartObject chart)
        {
            Excel.ChartObjects charts = (Excel.ChartObjects)sheet.ChartObjects();
            for (int i = 1; i <= charts.Count; i++)
            {
                if (ReferenceEquals((Excel.ChartObject)charts.Item(i), chart)
                    || string.Equals(((Excel.ChartObject)charts.Item(i)).Name, chart.Name, StringComparison.Ordinal))
                {
                    return i - 1;
                }
            }

            return -1;
        }

        private static Excel.ChartObject ChartObjectAt(Excel.Worksheet sheet, int index)
        {
            if (sheet == null)
            {
                return null;
            }

            Excel.ChartObjects charts = (Excel.ChartObjects)sheet.ChartObjects();
            if (index < 0 || index >= charts.Count)
            {
                return null;
            }

            return (Excel.ChartObject)charts.Item(index + 1);
        }

        private static Excel.Worksheet FindSheet(Excel.Workbook book, string name)
        {
            foreach (Excel.Worksheet sheet in book.Worksheets)
            {
                if (string.Equals(sheet.Name, name, StringComparison.OrdinalIgnoreCase))
                {
                    return sheet;
                }
            }

            return null;
        }

        private static HashSet<string> ChartNames(Excel.Worksheet sheet)
        {
            var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            Excel.ChartObjects charts = (Excel.ChartObjects)sheet.ChartObjects();
            for (int i = 1; i <= charts.Count; i++)
            {
                names.Add(((Excel.ChartObject)charts.Item(i)).Name ?? "");
            }

            return names;
        }

        private static Excel.ChartObject FindNewChart(Excel.Worksheet sheet, HashSet<string> before)
        {
            Excel.ChartObjects charts = (Excel.ChartObjects)sheet.ChartObjects();
            for (int i = 1; i <= charts.Count; i++)
            {
                Excel.ChartObject co = (Excel.ChartObject)charts.Item(i);
                if (!before.Contains(co.Name ?? ""))
                {
                    return co;
                }
            }

            return charts.Count > 0 ? (Excel.ChartObject)charts.Item(charts.Count) : null;
        }

        private static int ChartObjectIndexEt(object sheet, object chartHost)
        {
            object charts = EtCom.Invoke(sheet, "ChartObjects");
            int count = Convert.ToInt32(EtCom.GetProperty(charts, "Count"));
            string name = Convert.ToString(EtCom.GetProperty(chartHost, "Name")) ?? "";
            for (int i = 1; i <= count; i++)
            {
                object item = EtCom.Invoke(charts, "Item", i);
                string itemName = Convert.ToString(EtCom.GetProperty(item, "Name")) ?? "";
                if (string.Equals(itemName, name, StringComparison.Ordinal))
                {
                    return i - 1;
                }
            }

            return -1;
        }

        private static object ChartObjectAtEt(object sheet, int index)
        {
            if (sheet == null)
            {
                return null;
            }

            object charts = EtCom.Invoke(sheet, "ChartObjects");
            int count = Convert.ToInt32(EtCom.GetProperty(charts, "Count"));
            if (index < 0 || index >= count)
            {
                return null;
            }

            return EtCom.Invoke(charts, "Item", index + 1);
        }

        private static object FindSheetEt(object book, string name)
        {
            object sheets = EtCom.GetProperty(book, "Worksheets");
            int count = Convert.ToInt32(EtCom.GetProperty(sheets, "Count"));
            for (int i = 1; i <= count; i++)
            {
                object sheet = EtCom.Invoke(sheets, "Item", i);
                if (string.Equals(EtCom.TryReadName(sheet), name, StringComparison.OrdinalIgnoreCase))
                {
                    return sheet;
                }
            }

            return null;
        }

        private static HashSet<string> ChartNamesEt(object sheet)
        {
            var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            object charts = EtCom.Invoke(sheet, "ChartObjects");
            int count = Convert.ToInt32(EtCom.GetProperty(charts, "Count"));
            for (int i = 1; i <= count; i++)
            {
                object item = EtCom.Invoke(charts, "Item", i);
                names.Add(Convert.ToString(EtCom.GetProperty(item, "Name")) ?? "");
            }

            return names;
        }

        private static object FindNewChartEt(object sheet, HashSet<string> before)
        {
            object charts = EtCom.Invoke(sheet, "ChartObjects");
            int count = Convert.ToInt32(EtCom.GetProperty(charts, "Count"));
            object last = null;
            for (int i = 1; i <= count; i++)
            {
                object item = EtCom.Invoke(charts, "Item", i);
                last = item;
                string name = Convert.ToString(EtCom.GetProperty(item, "Name")) ?? "";
                if (!before.Contains(name))
                {
                    return item;
                }
            }

            return last;
        }
    }
}
