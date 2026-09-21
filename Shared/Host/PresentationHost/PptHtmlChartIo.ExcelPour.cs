using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using System.Xml.Linq;
using WordAddIn1.OpenFiles;


namespace WordAddIn1.PresentationHost
{
    internal static partial class PptHtmlChartIo
    {
        private const int SwHide = 0;
        private const uint WmClose = 0x0010;
        private const uint WmSysCommand = 0x0112;
        private static readonly IntPtr ScClose = new IntPtr(0xF060);

        private delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

        [DllImport("user32.dll")]
        private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        [DllImport("user32.dll")]
        private static extern bool IsWindowVisible(IntPtr hWnd);

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

        [DllImport("user32.dll")]
        private static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);

        [DllImport("user32.dll")]
        private static extern bool PostMessage(IntPtr hWnd, uint msg, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll")]
        private static extern IntPtr SendMessage(IntPtr hWnd, uint msg, IntPtr wParam, IntPtr lParam);

        private static void SuppressExcel(object excelApp)
        {
            if (excelApp == null)
            {
                return;
            }

            try
            {
                WppCom.TrySetProperty(excelApp, "DisplayAlerts", false);
                WppCom.TrySetProperty(excelApp, "ScreenUpdating", false);
                WppCom.TrySetProperty(excelApp, "Visible", false);
            }
            catch (Exception)
            {
            }
        }

        private static void HideEmbeddedExcel(object excelApp)
        {
            SuppressExcel(excelApp);
            if (excelApp == null)
            {
                return;
            }

            try
            {
                object windows = WppCom.GetProperty(excelApp, "Windows");
                int count = Convert.ToInt32(WppCom.GetProperty(windows, "Count"));
                for (int i = count; i >= 1; i--)
                {
                    object win = WppCom.GetIndexed(windows, i);
                    WppCom.TrySetProperty(win, "Visible", false);
                }
            }
            catch (Exception)
            {
            }

            TryHideExcelAppHwnd(excelApp);
        }

        /// <summary>
        /// 建/灌 chart 前先把已知的图表 Excel 窗藏起来，减轻 AddChart2 闪窗。
        /// </summary>
        private static void PrepareChartExcelUiSuppression()
        {
            TryHidePowerPointChartExcelWindows();
        }

        /// <summary>
        /// 单次 chart 操作后：藏 COM + 关 HWND；AddChart2 常在返回后才画出编辑框，须轮询。
        /// </summary>
        private static void DismissChartExcelUiForChart(object chart)
        {
            TryHideChartExcel(chart);
            DismissChartExcelUi();
        }

        private static bool TryGetEmbeddedExcelApp(object chart, out object excelApp)
        {
            excelApp = null;
            if (chart == null)
            {
                return false;
            }

            try
            {
                object chartData = WppCom.GetProperty(chart, "ChartData");
                object workbook = chartData == null ? null : WppCom.GetProperty(chartData, "Workbook");
                excelApp = workbook == null ? null : WppCom.GetProperty(workbook, "Application");
                return excelApp != null;
            }
            catch (Exception)
            {
                return false;
            }
        }

        /// <summary>
        /// 藏图表拉起的内嵌 Excel，不 Quit，避免弄坏包内 embeddings。
        /// COM Visible=false 常只藏内容，PowerPoint 会留下空白编辑框，须再关 HWND。
        /// 不经 Workbook 取 Application：刚 Close 后再取会把表重新 Activate。
        /// </summary>
        private static void TryHideChartExcel(object chart)
        {
            TryHidePowerPointChartExcelWindows();
            TryCloseChartExcelHwnds(forceClose: false);
        }

        /// <summary>
        /// 关掉本次打开的内嵌簿并放开 RCW。不 Quit ET/Excel。
        /// WPP 上一张表还占着时第二次 Activate 会失败；删图能再开，说明 Close 可能够用。
        /// </summary>
        private static void ReleaseEmbeddedChartWorkbook(
            object chart,
            object excelApp,
            List<string> warnings = null)
        {
            object chartData = null;
            object workbook = null;
            try
            {
                if (chart != null)
                {
                    try
                    {
                        chartData = WppCom.GetProperty(chart, "ChartData");
                    }
                    catch (Exception)
                    {
                    }
                }

                if (chartData != null)
                {
                    try
                    {
                        workbook = WppCom.GetProperty(chartData, "Workbook");
                    }
                    catch (Exception)
                    {
                    }
                }

                if (workbook != null)
                {
                    try
                    {
                        WppCom.Invoke(workbook, "Close", false);
                        PourLog(warnings, "已 Close 内嵌簿");
                    }
                    catch (Exception ex1)
                    {
                        try
                        {
                            WppCom.Invoke(workbook, "Close");
                            PourLog(warnings, "已 Close 内嵌簿");
                        }
                        catch (Exception)
                        {
                            PourLog(warnings, "Close 内嵌簿失败: " + ex1.Message);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                PourLog(warnings, "释放内嵌簿失败: " + ex.Message);
            }
            finally
            {
                HideEmbeddedExcel(excelApp);
                TryReleaseCom(workbook);
                TryReleaseCom(chartData);
                try
                {
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                }
                catch (Exception)
                {
                }

                Thread.Sleep(200);
            }
        }

        private static void TryReleaseCom(object com)
        {
            if (com == null)
            {
                return;
            }

            try
            {
                Marshal.FinalReleaseComObject(com);
            }
            catch (Exception)
            {
            }
        }

        /// <summary>
        /// apply 整页结束后再清一次：AddChart2 常在 COM 返回后才把编辑框画出来。
        /// </summary>
        public static void DismissChartExcelUi()
        {
            for (int i = 0; i < 6; i++)
            {
                TryHidePowerPointChartExcelWindows();
                TryCloseChartExcelHwnds(forceClose: i >= 2);
                if (!HasVisibleChartExcelWindow())
                {
                    return;
                }

                try
                {
                    Application.DoEvents();
                }
                catch (Exception)
                {
                }

                Thread.Sleep(80);
            }
        }

        private static void TryHidePowerPointChartExcelWindows()
        {
            foreach (string progId in new[] { "Excel.Application", "Ket.Application", "et.Application" })
            {
                object excelApp = null;
                try
                {
                    excelApp = Marshal.GetActiveObject(progId);
                }
                catch (Exception)
                {
                    continue;
                }

                if (excelApp == null)
                {
                    continue;
                }

                try
                {
                    object windows = WppCom.GetProperty(excelApp, "Windows");
                    int count = Convert.ToInt32(WppCom.GetProperty(windows, "Count"));
                    bool hidChartWindow = false;
                    int stillVisible = 0;
                    for (int i = 1; i <= count; i++)
                    {
                        object win = WppCom.GetIndexed(windows, i);
                        string caption = Convert.ToString(WppCom.GetProperty(win, "Caption") ?? "");
                        if (IsPowerPointChartExcelCaption(caption))
                        {
                            WppCom.TrySetProperty(win, "Visible", false);
                            hidChartWindow = true;
                        }
                        else if (IsTruthy(WppCom.GetProperty(win, "Visible")))
                        {
                            stillVisible++;
                        }
                    }

                    if (hidChartWindow && stillVisible == 0)
                    {
                        SuppressExcel(excelApp);
                    }

                    TryHideExcelAppHwnd(excelApp);
                }
                catch (Exception)
                {
                }
            }
        }

        private static void TryHideExcelAppHwnd(object excelApp)
        {
            if (excelApp == null)
            {
                return;
            }

            try
            {
                object raw = WppCom.GetProperty(excelApp, "Hwnd");
                if (raw == null)
                {
                    return;
                }

                IntPtr hwnd = new IntPtr(Convert.ToInt64(raw, CultureInfo.InvariantCulture));
                if (hwnd == IntPtr.Zero)
                {
                    return;
                }

                string title = GetWindowTitle(hwnd);
                if (!IsPowerPointChartExcelCaption(title))
                {
                    return;
                }

                ShowWindow(hwnd, SwHide);
                PostMessage(hwnd, WmClose, IntPtr.Zero, IntPtr.Zero);
            }
            catch (Exception)
            {
            }
        }

        private static bool TryCloseChartExcelHwnds(bool forceClose)
        {
            List<IntPtr> targets = FindChartExcelHwnds(visibleOnly: !forceClose);
            if (targets.Count == 0 && forceClose)
            {
                targets = FindChartExcelHwnds(visibleOnly: false);
            }

            foreach (IntPtr hwnd in targets)
            {
                try
                {
                    ShowWindow(hwnd, SwHide);
                    PostMessage(hwnd, WmClose, IntPtr.Zero, IntPtr.Zero);
                    if (forceClose && IsWindowVisible(hwnd))
                    {
                        SendMessage(hwnd, WmSysCommand, ScClose, IntPtr.Zero);
                    }
                }
                catch (Exception)
                {
                }
            }

            return targets.Count > 0;
        }

        private static bool HasVisibleChartExcelWindow()
        {
            return FindChartExcelHwnds(visibleOnly: true).Count > 0;
        }

        private static List<IntPtr> FindChartExcelHwnds(bool visibleOnly)
        {
            var found = new List<IntPtr>();
            try
            {
                EnumWindows((hWnd, _) =>
                {
                    if (visibleOnly && !IsWindowVisible(hWnd))
                    {
                        return true;
                    }

                    if (IsPowerPointChartExcelCaption(GetWindowTitle(hWnd)))
                    {
                        found.Add(hWnd);
                    }

                    return true;
                }, IntPtr.Zero);
            }
            catch (Exception)
            {
            }

            return found;
        }

        private static string GetWindowTitle(IntPtr hWnd)
        {
            try
            {
                var sb = new StringBuilder(512);
                if (GetWindowText(hWnd, sb, sb.Capacity) <= 0)
                {
                    return "";
                }

                return sb.ToString();
            }
            catch (Exception)
            {
                return "";
            }
        }

        private static bool IsPowerPointChartExcelCaption(string caption)
        {
            if (string.IsNullOrEmpty(caption))
            {
                return false;
            }

            if (caption.IndexOf("PowerPoint 中的图表", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return true;
            }

            if (caption.IndexOf("Chart in Microsoft PowerPoint", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return true;
            }

            if (caption.IndexOf("中的图表 - Excel", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return true;
            }

            if (caption.IndexOf("中的图表 - WPS", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return true;
            }

            if (caption.IndexOf("Chart in WPS", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return true;
            }

            if (caption.IndexOf("演示中的图表", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return true;
            }

            if (caption.IndexOf("演示文稿", StringComparison.OrdinalIgnoreCase) >= 0
                && caption.IndexOf("图表", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return true;
            }

            return false;
        }

        private static bool TryPourViaChartData(
            object chart,
            PptHtmlChartGrid grid,
            out object excelApp,
            out string error,
            List<string> warnings = null)
        {
            excelApp = null;
            error = null;
            try
            {
                object chartData = WppCom.GetProperty(chart, "ChartData");
                if (IsChartDataLinked(chart))
                {
                    if (!TryBreakChartDataLink(chart, warnings))
                    {
                        error = "图表为外链数据源且无法断链，请先换成内嵌表";
                        return false;
                    }
                }

                if (!TryOpenChartWorksheet(chart, out object ws, out excelApp, out error))
                {
                    return false;
                }

                object workbook = chartData == null ? null : WppCom.GetProperty(chartData, "Workbook");
                PourLog(warnings, "ChartData IsLinked=" + (TryPropString(chartData, "IsLinked") ?? "?"));
                PourLog(warnings, "打开内嵌簿 " + (TryPropString(workbook, "Name") ?? "?")
                    + " | " + DescribeSheetCells(ws, 5, 2));
                int cols = grid.Columns.Count;
                int rows = grid.Rows.Count;
                TryResizeEmbeddedListObject(ws, rows + 1, cols, warnings);
                for (int c = 0; c < cols; c++)
                {
                    SetCell(ws, 1, c + 1, grid.Columns[c].Name ?? "");
                }

                for (int r = 0; r < rows; r++)
                {
                    for (int c = 0; c < cols; c++)
                    {
                        string raw = r < grid.Rows.Count && c < grid.Rows[r].Count ? grid.Rows[r][c] : "";
                        if (grid.Columns[c].Role == "value"
                            && double.TryParse(raw, NumberStyles.Float, CultureInfo.InvariantCulture, out double n))
                        {
                            SetCell(ws, r + 2, c + 1, n);
                        }
                        else
                        {
                            SetCell(ws, r + 2, c + 1, raw ?? "");
                        }
                    }
                }

                // PPT 上 Chart.SetSourceData 常失败，不调用；图绑定由后续 Series 同步完成。
                // 内嵌表仍是读/验真源（I6/I8）。
                PourLog(warnings, "写入内嵌表 " + DescribeSheetCells(ws, rows + 1, cols)
                    + " range=" + TryRangeAddress(TryGetDataRange(ws, rows + 1, cols))
                    + "（跳过 SetSourceData，改由 Series 绑定）");
                PourLog(warnings, "ChartData 写入完成 " + DescribeLiveSeries(chart));
                HideEmbeddedExcel(excelApp);
                return true;
            }
            catch (Exception ex)
            {
                error = "灌入 ChartData 内嵌表失败: " + ex.Message;
                PourLog(warnings, error);
                return false;
            }
        }

        /// <summary>
        /// 只读 IsLinked，不访问 Workbook（访问会触发「链接的文件不可用」模态框）。
        /// </summary>
        private static bool IsChartDataLinked(object chart)
        {
            if (chart == null)
            {
                return false;
            }

            try
            {
                object chartData = WppCom.GetProperty(chart, "ChartData");
                if (chartData == null)
                {
                    return false;
                }

                object linked = WppCom.GetProperty(chartData, "IsLinked");
                if (linked == null)
                {
                    return false;
                }

                if (linked is bool b)
                {
                    return b;
                }

                return Convert.ToBoolean(linked);
            }
            catch (Exception)
            {
                return false;
            }
        }

        private static bool TryBreakChartDataLink(object chart, List<string> warnings = null)
        {
            if (chart == null || !IsChartDataLinked(chart))
            {
                return true;
            }

            try
            {
                object chartData = WppCom.GetProperty(chart, "ChartData");
                if (chartData == null)
                {
                    return false;
                }

                WppCom.Invoke(chartData, "BreakLink");
                PourLog(warnings, "已 BreakLink，改为内嵌表");
                return !IsChartDataLinked(chart);
            }
            catch (Exception ex)
            {
                PourLog(warnings, "BreakLink 失败: " + ex.Message);
                return false;
            }
        }

        private static bool TryOpenChartWorksheet(
            object chart,
            out object ws,
            out object excelApp,
            out string error)
        {
            ws = null;
            excelApp = null;
            error = null;
            if (chart == null)
            {
                error = "chart 为 null";
                return false;
            }

            try
            {
                object chartData = WppCom.GetProperty(chart, "ChartData");
                if (chartData == null)
                {
                    error = "无法访问 ChartData";
                    return false;
                }

                // 外链且未断链时禁止碰 Workbook，否则 Office 弹「链接的文件不可用」并卡住自动化。
                if (IsChartDataLinked(chart))
                {
                    error = "图表数据源为外链，跳过打开 Workbook";
                    return false;
                }

                object workbook = TryOpenChartWorkbook(chartData, out error);
                if (workbook == null)
                {
                    if (string.IsNullOrEmpty(error))
                    {
                        error = "无法打开图表内嵌工作簿";
                    }

                    return false;
                }

                excelApp = WppCom.GetProperty(workbook, "Application");
                HideEmbeddedExcel(excelApp);
                object sheets = WppCom.GetProperty(workbook, "Worksheets");
                ws = WppCom.GetIndexed(sheets, 1);
                if (ws == null)
                {
                    error = "图表内嵌表不存在";
                    return false;
                }

                return true;
            }
            catch (Exception ex)
            {
                error = "打开 ChartData 失败: " + ex.Message;
                return false;
            }
        }

        /// <summary>
        /// 先 Activate 再取 Workbook。连续建图时嵌入 Excel 常未就绪，直取会炸。
        /// </summary>
        private static object TryOpenChartWorkbook(object chartData, out string error)
        {
            error = null;
            if (chartData == null)
            {
                error = "无法访问 ChartData";
                return null;
            }

            Exception last = null;
            for (int i = 0; i < 6; i++)
            {
                TryActivateChartData(chartData);
                PumpChartUi(80 + i * 70);
                try
                {
                    object workbook = WppCom.GetProperty(chartData, "Workbook");
                    if (workbook != null)
                    {
                        return workbook;
                    }
                }
                catch (Exception ex)
                {
                    last = ex;
                }
            }

            error = last == null
                ? "无法打开图表内嵌工作簿"
                : "打开 ChartData 失败: " + last.Message;
            return null;
        }

        private static void TryActivateChartData(object chartData)
        {
            if (chartData == null)
            {
                return;
            }

            try
            {
                WppCom.Invoke(chartData, "Activate");
            }
            catch (Exception)
            {
            }
        }

        private static void PumpChartUi(int milliseconds)
        {
            try
            {
                Application.DoEvents();
            }
            catch (Exception)
            {
            }

            if (milliseconds > 0)
            {
                Thread.Sleep(milliseconds);
            }
        }

        private static List<string> TryReadCategoryAxisNames(object chart)
        {
            var cats = new List<string>();
            if (chart == null)
            {
                return cats;
            }

            try
            {
                object axis = TryInvoke(chart, "Axes", 1, 1)
                    ?? TryInvoke(chart, "Axes", 1);
                if (axis == null)
                {
                    return cats;
                }

                return ToStringList(WppCom.GetProperty(axis, "CategoryNames"));
            }
            catch (Exception)
            {
                return cats;
            }
        }

        private static bool TryBuildGridFromWorksheet(
            object ws,
            int cols,
            int dataRows,
            out PptHtmlChartGrid grid,
            out string error)
        {
            grid = null;
            error = null;
            if (ws == null || cols < 2 || dataRows < 1)
            {
                error = "内嵌表行列不足";
                return false;
            }

            try
            {
                var columns = new List<PptHtmlChartColumn>();
                for (int c = 0; c < cols; c++)
                {
                    string name = Convert.ToString(GetCell(ws, 1, c + 1) ?? "");
                    columns.Add(new PptHtmlChartColumn
                    {
                        Role = c == 0 ? "category" : "value",
                        Name = name
                    });
                }

                var dataRowList = new List<List<string>>();
                for (int r = 0; r < dataRows; r++)
                {
                    var row = new List<string>();
                    for (int c = 0; c < cols; c++)
                    {
                        object raw = GetCell(ws, r + 2, c + 1);
                        row.Add(FormatCell(raw));
                    }

                    dataRowList.Add(row);
                }

                grid = new PptHtmlChartGrid
                {
                    Columns = columns,
                    Rows = dataRowList,
                    Truncated = dataRows + 1 > MaxRows || cols > MaxCols
                };
                if (!grid.IsPourable)
                {
                    error = "内嵌表内容不可灌";
                    grid = null;
                    return false;
                }

                return true;
            }
            catch (Exception ex)
            {
                error = "解析内嵌表失败: " + ex.Message;
                return false;
            }
        }

        private static bool TryReadGridFromEmbeddedSheetAuto(
            object chart,
            out PptHtmlChartGrid grid,
            out string error)
        {
            grid = null;
            error = null;
            object excelApp = null;
            try
            {
                if (!TryOpenChartWorksheet(chart, out object ws, out excelApp, out error))
                {
                    return false;
                }

                object used = WppCom.GetProperty(ws, "UsedRange");
                if (used == null)
                {
                    error = "内嵌表 UsedRange 为空";
                    return false;
                }

                int totalRows = Convert.ToInt32(WppCom.GetProperty(WppCom.GetProperty(used, "Rows"), "Count"));
                int totalCols = Convert.ToInt32(WppCom.GetProperty(WppCom.GetProperty(used, "Columns"), "Count"));
                if (totalRows < 2 || totalCols < 2)
                {
                    error = "内嵌表行/列不足（需表头+至少一行数据、两列）";
                    return false;
                }

                int dataRows = Math.Min(totalRows - 1, MaxRows - 1);
                int cols = Math.Min(totalCols, MaxCols);
                return TryBuildGridFromWorksheet(ws, cols, dataRows, out grid, out error);
            }
            catch (Exception ex)
            {
                error = "自动读内嵌表失败: " + ex.Message;
                return false;
            }
            finally
            {
                HideEmbeddedExcel(excelApp);
                DismissChartExcelUiForChart(chart);
            }
        }

        private static object TryGetDataRange(object ws, int lastRow, int lastCol)
        {
            string a2 = ColLetter(lastCol) + lastRow.ToString(CultureInfo.InvariantCulture);
            object range = TryGetExcelRange(ws, "A1", a2);
            if (range != null)
            {
                return range;
            }

            range = TryGetExcelRange(ws, "A1:" + a2, null);
            if (range != null)
            {
                return range;
            }

            object c1 = GetCellObject(ws, 1, 1);
            object c2 = GetCellObject(ws, lastRow, lastCol);
            if (c1 != null && c2 != null)
            {
                range = TryGetExcelRange(ws, c1, c2);
                if (range != null)
                {
                    return range;
                }
            }

            try
            {
                return WppCom.GetProperty(ws, "UsedRange");
            }
            catch (Exception)
            {
                return null;
            }
        }

        private static object TryGetExcelRange(object ws, object arg1, object arg2)
        {
            if (ws == null || arg1 == null)
            {
                return null;
            }

            object[] args = arg2 == null ? new object[] { arg1 } : new object[] { arg1, arg2 };
            const BindingFlags flags = BindingFlags.GetProperty
                | BindingFlags.InvokeMethod
                | BindingFlags.Instance
                | BindingFlags.Public;
            try
            {
                return ws.GetType().InvokeMember("Range", flags, null, ws, args);
            }
            catch (Exception)
            {
            }

            try
            {
                return ws.GetType().InvokeMember("get_Range", flags, null, ws, args);
            }
            catch (Exception)
            {
                return null;
            }
        }

        /// <summary>
        /// PPT 内嵌表常锁在 ListObject；先 Resize 再写入，勿整块 Clear 拆表。
        /// </summary>
        private static void TryResizeEmbeddedListObject(
            object ws,
            int lastRow,
            int lastCol,
            List<string> warnings)
        {
            if (ws == null || lastRow < 2 || lastCol < 2)
            {
                return;
            }

            try
            {
                object lists = WppCom.GetProperty(ws, "ListObjects");
                if (lists == null)
                {
                    return;
                }

                int count = Convert.ToInt32(WppCom.GetProperty(lists, "Count"));
                if (count < 1)
                {
                    return;
                }

                object lo = WppCom.GetIndexed(lists, 1);
                string corner = ColLetter(lastCol) + lastRow.ToString(CultureInfo.InvariantCulture);
                object resizeRange = TryGetExcelRange(ws, "A1", corner)
                    ?? TryGetExcelRange(ws, "A1:" + corner, null);
                if (resizeRange == null)
                {
                    PourLog(warnings, "ListObject Resize 跳过：无法构造范围 A1:" + corner);
                    return;
                }

                WppCom.Invoke(lo, "Resize", resizeRange);
                PourLog(warnings, "ListObject Resize A1:" + corner + " ok");
            }
            catch (Exception ex)
            {
                PourLog(warnings, "ListObject Resize 失败: " + ex.Message);
            }
        }

        private static object GetSeries(object chart, int index)
        {
            object series = TryInvoke(chart, "SeriesCollection", index);
            if (series != null)
            {
                return series;
            }

            object sc = TryInvoke(chart, "SeriesCollection");
            if (sc == null)
            {
                sc = WppCom.GetProperty(chart, "SeriesCollection");
            }

            return WppCom.GetIndexed(sc, index);
        }

        private static int GetSeriesCount(object chart)
        {
            object sc = TryInvoke(chart, "SeriesCollection");
            if (sc == null)
            {
                sc = WppCom.GetProperty(chart, "SeriesCollection");
            }

            return Convert.ToInt32(WppCom.GetProperty(sc, "Count"));
        }

        private static bool FillRgbAlreadyMatches(object fill, int rgb)
        {
            try
            {
                object vis = WppCom.GetProperty(fill, "Visible");
                if (vis != null && Convert.ToInt32(vis) == 0)
                {
                    return false;
                }

                object fc = WppCom.GetProperty(fill, "ForeColor");
                object cur = fc == null ? null : WppCom.GetProperty(fc, "RGB");
                if (cur == null)
                {
                    return false;
                }

                return (Convert.ToInt32(cur) & 0x00FFFFFF) == (rgb & 0x00FFFFFF);
            }
            catch (Exception)
            {
                return false;
            }
        }

        private static bool IsSameChartType(object chart, string canon)
        {
            try
            {
                object t = WppCom.GetProperty(chart, "ChartType");
                if (t == null)
                {
                    return false;
                }

                return string.Equals(
                    CanonicalTypeFromXl(Convert.ToInt32(t)),
                    canon,
                    StringComparison.OrdinalIgnoreCase);
            }
            catch (Exception)
            {
                return false;
            }
        }

        private static int CountValueColumns(PptHtmlChartGrid grid)
        {
            int n = 0;
            if (grid == null || grid.Columns == null)
            {
                return 0;
            }

            for (int c = 0; c < grid.Columns.Count; c++)
            {
                if (grid.Columns[c] != null && grid.Columns[c].Role != "category")
                {
                    n++;
                }
            }

            return n;
        }

        private static bool TryVerifyPouredGrid(
            object chart,
            PptHtmlChartGrid want,
            out string error,
            List<string> warnings = null)
        {
            error = null;
            if (chart == null || want == null || !want.IsPourable)
            {
                error = "图表换数后无法校验内嵌表";
                return false;
            }

            PourLog(warnings, "校验 " + DescribeWantGrid(want) + " | " + DescribeLiveSeries(chart));
            if (!TryReadGridFromEmbeddedSheet(chart, want.Columns.Count, want.Rows.Count, out PptHtmlChartGrid live, out string readErr)
                && !TryReadGridFromEmbeddedSheetAuto(chart, out live, out readErr))
            {
                error = "无法回读内嵌表校验: " + (readErr ?? "?");
                PourLog(warnings, "校验失败 " + error);
                return false;
            }

            PourLog(warnings, "校验读回 " + DescribeGridCompact(live, "内嵌表"));
            string mismatch = FindFirstGridMismatch(want, live, compareNames: false);
            if (mismatch != null)
            {
                error = "图表内嵌表与稿不一致: " + mismatch;
                PourLog(warnings, "校验失败 " + mismatch);
                PourLog(warnings, "校验对照 稿=" + DescribeGridCompact(want, "稿")
                    + " | 内嵌表=" + DescribeGridCompact(live, "内嵌表"));
                return false;
            }

            PourLog(warnings, "校验通过（内嵌表与稿一致）");
            return true;
        }

        private static bool TryReadGridFromEmbeddedSheet(
            object chart,
            int cols,
            int rows,
            out PptHtmlChartGrid grid,
            out string error)
        {
            grid = null;
            error = null;
            object excelApp = null;
            try
            {
                if (!TryOpenChartWorksheet(chart, out object ws, out excelApp, out error))
                {
                    return false;
                }

                return TryBuildGridFromWorksheet(ws, cols, rows, out grid, out error);
            }
            catch (Exception ex)
            {
                error = "回读图表内嵌表失败: " + ex.Message;
                return false;
            }
            finally
            {
                HideEmbeddedExcel(excelApp);
                DismissChartExcelUiForChart(chart);
            }
        }

        private static bool GridAlreadyMatches(object chart, PptHtmlChartGrid grid)
        {
            if (chart == null || grid == null || !grid.IsPourable)
            {
                return false;
            }

            if (TryReadGridFromEmbeddedSheet(chart, grid.Columns.Count, grid.Rows.Count, out PptHtmlChartGrid cur, out _)
                && GridsMatch(grid, cur, compareNames: true))
            {
                return true;
            }

            return TryReadGridFromSeries(chart, out cur, out _)
                && GridsMatch(grid, cur, compareNames: true);
        }

        private static bool SeriesRowCountMatches(object chart, PptHtmlChartGrid grid)
        {
            return grid != null && grid.Rows != null && ReadSeriesRowCount(chart) == grid.Rows.Count;
        }

        private static int ReadSeriesRowCount(object chart)
        {
            try
            {
                object s1 = GetSeries(chart, 1);
                if (s1 == null)
                {
                    return 0;
                }

                List<string> cats = ToStringList(WppCom.GetProperty(s1, "XValues"));
                if (cats.Count > 0)
                {
                    return cats.Count;
                }

                return ToStringList(WppCom.GetProperty(s1, "Values")).Count;
            }
            catch (Exception)
            {
                return 0;
            }
        }

        private static bool GridsMatch(PptHtmlChartGrid want, PptHtmlChartGrid got, bool compareNames)
        {
            return FindFirstGridMismatch(want, got, compareNames) == null;
        }

        private static bool TryParseLooseDouble(string raw, out double n)
        {
            return double.TryParse(raw, NumberStyles.Float, CultureInfo.InvariantCulture, out n)
                || double.TryParse(raw, NumberStyles.Float, CultureInfo.CurrentCulture, out n);
        }

        private static int TryGetPointCount(object series)
        {
            object pts = TryGetPoints(series);
            if (pts == null)
            {
                return 0;
            }

            try
            {
                return Convert.ToInt32(WppCom.GetProperty(pts, "Count"));
            }
            catch (Exception)
            {
                return 0;
            }
        }

        /// <summary>
        /// 主题盘已在套回里先套。这里只盖更细的色：饼扇区或自动分色；柱/线系列填+点填。
        /// 饼图不套系列填。
        /// </summary>
    }
}
