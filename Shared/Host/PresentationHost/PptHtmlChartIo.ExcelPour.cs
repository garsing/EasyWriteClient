using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
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
        /// PPT 脏会话里 ChartData.Activate 会 RPC 弄死稿。WPP 仍走 ChartData。
        /// 不能复用 LooksLikeWppApp：PowerPoint 也有 ProductCode，会被误判成 WPP。
        /// </summary>
        private static bool HostAvoidsChartDataCom(object chart)
        {
            return HostAvoidsFromApp(TryGetChartHostApp(chart));
        }

        private static bool HostAvoidsFromShapes(object shapes)
        {
            return HostAvoidsFromApp(TryGetShapesHostApp(shapes));
        }

        private static bool HostAvoidsFromApp(object app)
        {
            if (app == null)
            {
                return true;
            }

            try
            {
                if (WppCom.GetProperty(app, "WpsPresentation") != null)
                {
                    return false;
                }
            }
            catch (Exception)
            {
            }

            string name = TryPropString(app, "Name") ?? "";
            if (name.IndexOf("wps", StringComparison.OrdinalIgnoreCase) >= 0
                || name.IndexOf("wpp", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return false;
            }

            return true;
        }

        private static object TryGetShapesHostApp(object shapes)
        {
            if (shapes == null)
            {
                return null;
            }

            try
            {
                object slide = WppCom.GetProperty(shapes, "Parent");
                object pres = slide == null ? null : WppCom.GetProperty(slide, "Parent");
                return pres == null ? null : WppCom.GetProperty(pres, "Application");
            }
            catch (Exception)
            {
                return null;
            }
        }

        private static object TryGetChartHostApp(object chart)
        {
            if (chart == null)
            {
                return null;
            }

            // 必须用 Presentation.Application：WPP 的 chart.Application.Name 会报 PowerPoint。
            try
            {
                object pres = TryGetShapePresentation(TryGetChartShape(chart));
                object app = pres == null ? null : WppCom.GetProperty(pres, "Application");
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
                return WppCom.GetProperty(chart, "Application");
            }
            catch (Exception)
            {
                return null;
            }
        }

        /// <summary>
        /// 单次 chart 操作后：关内嵌簿、藏窗、退出没有用户簿的孤儿 Excel。
        /// AddChart2 常在返回后才画出编辑框，须轮询。
        /// </summary>
        private static void DismissChartExcelUiForChart(object chart)
        {
            TryCloseChartDataWorkbook(chart);
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
        /// 藏图表拉起的内嵌 Excel。COM Visible=false 常只藏内容，须再关 HWND。
        /// 孤儿进程由 DismissChartExcelUi → TryQuitOrphanChartExcel 退。
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
        /// 藏完后退出只剩图表簿 / 空壳的 Excel，不碰用户自己的表。
        /// </summary>
        public static void DismissChartExcelUi()
        {
            for (int i = 0; i < 6; i++)
            {
                TryHidePowerPointChartExcelWindows();
                TryCloseChartExcelHwnds(forceClose: i >= 2);
                if (!HasVisibleChartExcelWindow())
                {
                    break;
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

            TryQuitOrphanChartExcel();
        }

        /// <summary>
        /// 不 Activate：有 Workbook 就 Close。网格还开着时再 AddChart 会报「图表数据网格已在…打开」。
        /// </summary>
        private static void TryCloseChartDataWorkbook(object chart)
        {
            if (chart == null)
            {
                return;
            }

            try
            {
                object chartData = WppCom.GetProperty(chart, "ChartData");
                object workbook = chartData == null ? null : WppCom.GetProperty(chartData, "Workbook");
                if (workbook == null)
                {
                    return;
                }

                try
                {
                    WppCom.Invoke(workbook, "Close", false);
                }
                catch (Exception)
                {
                    try
                    {
                        WppCom.Invoke(workbook, "Close");
                    }
                    catch (Exception)
                    {
                    }
                }

                TryReleaseCom(workbook);
                TryReleaseCom(chartData);
            }
            catch (Exception)
            {
            }
        }

        /// <summary>
        /// AddChart2 会留下 EXCEL.EXE。GetActiveObject 一次一个，Quit 空壳后再取下一个。
        /// 见到用户工作簿就停，禁止 Quit 用户表。
        /// </summary>
        private static void TryQuitOrphanChartExcel()
        {
            foreach (string progId in new[] { "Excel.Application", "Ket.Application", "et.Application" })
            {
                for (int n = 0; n < 12; n++)
                {
                    object excelApp = null;
                    try
                    {
                        excelApp = Marshal.GetActiveObject(progId);
                    }
                    catch (Exception)
                    {
                        break;
                    }

                    if (excelApp == null)
                    {
                        break;
                    }

                    if (HasUserWorkbook(excelApp))
                    {
                        TryReleaseCom(excelApp);
                        break;
                    }

                    CloseChartWorkbooks(excelApp);
                    if (WorkbookCount(excelApp) > 0)
                    {
                        TryReleaseCom(excelApp);
                        break;
                    }

                    try
                    {
                        WppCom.TrySetProperty(excelApp, "DisplayAlerts", false);
                        WppCom.Invoke(excelApp, "Quit");
                    }
                    catch (Exception)
                    {
                    }

                    TryReleaseCom(excelApp);
                    Thread.Sleep(60);
                }
            }
        }

        private static int WorkbookCount(object excelApp)
        {
            try
            {
                object books = WppCom.GetProperty(excelApp, "Workbooks");
                return books == null ? -1 : Convert.ToInt32(WppCom.GetProperty(books, "Count") ?? -1);
            }
            catch (Exception)
            {
                return -1;
            }
        }

        private static bool HasUserWorkbook(object excelApp)
        {
            try
            {
                object books = WppCom.GetProperty(excelApp, "Workbooks");
                int count = books == null ? 0 : Convert.ToInt32(WppCom.GetProperty(books, "Count") ?? 0);
                for (int i = 1; i <= count; i++)
                {
                    object book = WppCom.GetIndexed(books, i);
                    if (book != null && !IsChartWorkbook(book))
                    {
                        return true;
                    }
                }
            }
            catch (Exception)
            {
            }

            return false;
        }

        private static bool IsChartWorkbook(object workbook)
        {
            string name = Convert.ToString(WppCom.GetProperty(workbook, "Name") ?? "");
            string full = Convert.ToString(WppCom.GetProperty(workbook, "FullName") ?? "");
            if (IsPowerPointChartExcelCaption(name) || IsPowerPointChartExcelCaption(full))
            {
                return true;
            }

            if (!string.IsNullOrEmpty(full)
                && (full.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase)
                    || full.EndsWith(".xls", StringComparison.OrdinalIgnoreCase)
                    || full.EndsWith(".xlsm", StringComparison.OrdinalIgnoreCase))
                && File.Exists(full))
            {
                return false;
            }

            return false;
        }

        private static void CloseChartWorkbooks(object excelApp)
        {
            try
            {
                object books = WppCom.GetProperty(excelApp, "Workbooks");
                int count = books == null ? 0 : Convert.ToInt32(WppCom.GetProperty(books, "Count") ?? 0);
                for (int i = count; i >= 1; i--)
                {
                    object book = WppCom.GetIndexed(books, i);
                    if (book == null || !IsChartWorkbook(book))
                    {
                        continue;
                    }

                    try
                    {
                        WppCom.Invoke(book, "Close", false);
                    }
                    catch (Exception)
                    {
                    }
                }
            }
            catch (Exception)
            {
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
                PourLog(warnings, "ChartData IsLinked=" + (TryPropString(chartData, "IsLinked") ?? "?")
                    + " 内嵌应用=" + (TryPropString(excelApp, "Name") ?? "?")
                    + " ver=" + (TryPropString(excelApp, "Version") ?? "?"));
                PourLog(warnings, "打开内嵌簿 " + (TryPropString(workbook, "Name") ?? "?")
                    + " | " + DescribeSheetCells(ws, 5, 4)
                    + " | " + DescribeListBind(ws));
                DumpPourState("打开后", chart, ws, grid, warnings);
                int cols = grid.Columns.Count;
                int rows = grid.Rows.Count;
                // WPP 默认无 ListObject；强行 Add 会把 SERIES 从 $A$2:$A$5 错位成 $A$3:$A$6。
                // PPT 自带表对象，只 Resize。
                TryResizeEmbeddedListObject(ws, rows + 1, cols, warnings);
                DumpPourState("扩表后", chart, ws, grid, warnings);
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

                TryClearSheetBeyond(ws, rows + 1, cols, warnings);
                object source = TryGetDataRange(ws, rows + 1, cols);
                PourLog(warnings, "写入内嵌表 " + DescribeSheetCells(ws, rows + 1, cols)
                    + " range=" + DescribeComArg(source)
                    + " | " + DescribeListBind(ws));
                DumpPourState("写格后", chart, ws, grid, warnings);
                if (ReadSeriesRowCount(chart) != rows)
                {
                    TryBindSeriesToSheetFormulas(chart, ws, grid, warnings);
                    DumpPourState("改公式后", chart, ws, grid, warnings);
                }

                if (ReadSeriesRowCount(chart) != rows)
                {
                    TryInvoke(chart, "Refresh");
                    PourLog(warnings, "Refresh 后 " + DescribeLiveSeries(chart)
                        + " | " + DescribeSeriesExtra(chart));
                }

                if (ReadSeriesRowCount(chart) != rows)
                {
                    PourLog(warnings, "系列仍是 " + ReadSeriesRowCount(chart)
                        + " 点，稿要 " + rows + "（WPP 赋数组常截在 AddChart2 默认 4 点）");
                    DumpPourState("绑源失败", chart, ws, grid, warnings);
                }

                PourLog(warnings, "ChartData 写入完成 " + DescribeLiveSeries(chart)
                    + " | " + DescribeListBind(ws)
                    + " | " + DescribeSeriesExtra(chart));
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
                int curCols = 0;
                int curRows = 0;
                try
                {
                    object loRange = WppCom.GetProperty(lo, "Range");
                    curRows = Convert.ToInt32(WppCom.GetProperty(WppCom.GetProperty(loRange, "Rows"), "Count"));
                    curCols = Convert.ToInt32(WppCom.GetProperty(WppCom.GetProperty(loRange, "Columns"), "Count"));
                }
                catch (Exception)
                {
                }

                // 先扩行、后缩列：WPP/PPT 从默认 A1:D5 一次缩成 A1:B7 常失败。
                if (curCols > lastCol && lastRow > curRows)
                {
                    string expandCorner = ColLetter(curCols) + lastRow.ToString(CultureInfo.InvariantCulture);
                    object expandRange = TryGetExcelRange(ws, "A1", expandCorner)
                        ?? TryGetExcelRange(ws, "A1:" + expandCorner, null);
                    if (expandRange != null)
                    {
                        WppCom.Invoke(lo, "Resize", expandRange);
                        PourLog(warnings, "ListObject 先扩行 A1:" + expandCorner + " ok");
                    }
                }

                string corner = ColLetter(lastCol) + lastRow.ToString(CultureInfo.InvariantCulture);
                object resizeRange = TryGetExcelRange(ws, "A1", corner)
                    ?? TryGetExcelRange(ws, "A1:" + corner, null);
                if (resizeRange == null)
                {
                    PourLog(warnings, "ListObject Resize 跳过：无法构造范围 A1:" + corner);
                    return;
                }

                WppCom.Invoke(lo, "Resize", resizeRange);
                PourLog(warnings, "ListObject Resize A1:" + corner + " ok 前="
                    + curRows + "x" + curCols + " 后=" + DescribeListBind(ws));
            }
            catch (Exception ex)
            {
                PourLog(warnings, "ListObject Resize 失败: " + FormatComError(ex));
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
            PptHtmlChartGrid live;
            string readErr;
            bool liveFromOoxml = false;
            if (HostAvoidsChartDataCom(chart))
            {
                if (!TryReadGridFromSeries(chart, out live, out readErr) || live == null || !live.IsPourable)
                {
                    string seriesErr = readErr;
                    if (!TryReadGridFromOoxmlCopy(chart, out live, out readErr)
                        || live == null
                        || !live.IsPourable)
                    {
                        error = "无法回读系列校验: " + (seriesErr ?? readErr ?? "?");
                        PourLog(warnings, "校验失败 " + error);
                        return false;
                    }

                    liveFromOoxml = true;
                    PourLog(warnings, "校验 Series 失败，改读 OOXML");
                }
            }
            else if (!TryReadGridFromEmbeddedSheet(chart, want.Columns.Count, want.Rows.Count, out live, out readErr)
                && !TryReadGridFromEmbeddedSheetAuto(chart, out live, out readErr))
            {
                error = "无法回读内嵌表校验: " + (readErr ?? "?");
                PourLog(warnings, "校验失败 " + error);
                return false;
            }

            PourLog(warnings, "校验读回 " + DescribeGridCompact(live, liveFromOoxml ? "OOXML" : "内嵌表"));
            string mismatch = FindFirstGridMismatch(want, live, compareNames: false);
            if (mismatch != null)
            {
                error = "图表内嵌表与稿不一致: " + mismatch;
                PourLog(warnings, "校验失败 " + mismatch);
                PourLog(warnings, "校验对照 稿=" + DescribeGridCompact(want, "稿")
                    + " | 读回=" + DescribeGridCompact(live, liveFromOoxml ? "OOXML" : "内嵌表"));
                return false;
            }

            // OOXML c:pt 已对上稿；WPP 误判 PPT 时 Series COM 常炸，再数点数只会假失败。
            if (liveFromOoxml)
            {
                PourLog(warnings, "校验通过（OOXML 与稿一致）");
                return true;
            }

            int seriesRows = ReadSeriesRowCount(chart);
            if (seriesRows != want.Rows.Count)
            {
                error = "系列点数 " + seriesRows + " 与稿行数 " + want.Rows.Count
                    + " 不一致（内嵌表已写入，图未绑到整表）";
                PourLog(warnings, "校验失败 " + error + " | " + DescribeLiveSeries(chart));
                return false;
            }

            PourLog(warnings, "校验通过（内嵌表与稿一致，系列 " + seriesRows + " 点）");
            return true;
        }

        /// <summary>
        /// WPP 图认 SERIES 公式，不认 ET Range / 数组扩点。
        /// 默认是 =SERIES(Sheet1!$B$1,Sheet1!$A$2:$A$5,Sheet1!$B$2:$B$5,1)，改成 $A$7/$B$7。
        /// </summary>
        private static void TryBindSeriesToSheetFormulas(
            object chart,
            object ws,
            PptHtmlChartGrid grid,
            List<string> warnings)
        {
            if (chart == null || ws == null || grid == null || grid.Rows == null)
            {
                return;
            }

            string sheet = TryPropString(ws, "Name") ?? "Sheet1";
            string sheetRef = SheetFormulaName(sheet);
            int last = grid.Rows.Count + 1;
            string lastText = last.ToString(CultureInfo.InvariantCulture);
            string cats = sheetRef + "!$A$2:$A$" + lastText;
            int si = 0;
            for (int c = 0; c < grid.Columns.Count; c++)
            {
                if (grid.Columns[c].Role == "category")
                {
                    continue;
                }

                si++;
                object series = GetSeries(chart, si);
                if (series == null)
                {
                    PourLog(warnings, "改公式无系列 " + si);
                    continue;
                }

                string col = ColLetter(c + 1);
                string nameRef = sheetRef + "!$" + col + "$1";
                string vals = sheetRef + "!$" + col + "$2:$" + col + "$" + lastText;
                string formula = "=SERIES(" + nameRef + "," + cats + "," + vals + ","
                    + si.ToString(CultureInfo.InvariantCulture) + ")";

                if (TrySetSeriesProp(series, "Values", "=" + vals, out string valErr)
                    || TrySetSeriesProp(series, "Values", vals, out valErr))
                {
                    PourLog(warnings, "Values 公式 S" + si + " " + vals
                        + " | " + DescribeLiveSeries(chart));
                }
                else
                {
                    PourLog(warnings, "Values 公式 S" + si + " 失败: " + valErr);
                }

                if (TrySetSeriesProp(series, "XValues", "=" + cats, out string xErr)
                    || TrySetSeriesProp(series, "XValues", cats, out xErr))
                {
                    PourLog(warnings, "XValues 公式 S" + si + " " + cats
                        + " | " + DescribeLiveSeries(chart));
                }
                else
                {
                    PourLog(warnings, "XValues 公式 S" + si + " 失败: " + xErr);
                }

                TrySetSeriesProp(series, "Name", "=" + nameRef, out _);
                TrySetSeriesProp(series, "Name", nameRef, out _);

                if (ReadSeriesRowCount(chart) == grid.Rows.Count)
                {
                    PourLog(warnings, "地址公式已把 S" + si + " 扩到 " + grid.Rows.Count + " 点");
                    continue;
                }

                if (TrySetSeriesProp(series, "Formula", formula, out string fErr))
                {
                    PourLog(warnings, "Formula S" + si + " " + formula
                        + " | " + DescribeSeriesExtra(chart) + " | " + DescribeLiveSeries(chart));
                    continue;
                }

                PourLog(warnings, "Formula S" + si + " 失败 " + formula + " | " + fErr);
                if (TrySetSeriesProp(series, "FormulaLocal", formula, out string flErr))
                {
                    PourLog(warnings, "FormulaLocal S" + si + " " + formula
                        + " | " + DescribeSeriesExtra(chart) + " | " + DescribeLiveSeries(chart));
                }
                else
                {
                    PourLog(warnings, "FormulaLocal S" + si + " 失败 " + formula + " | " + flErr);
                }
            }
        }

        private static string SheetFormulaName(string sheet)
        {
            if (string.IsNullOrEmpty(sheet))
            {
                return "Sheet1";
            }

            bool quote = false;
            for (int i = 0; i < sheet.Length; i++)
            {
                char ch = sheet[i];
                if (!(ch == '_' || ch == '.' || char.IsLetterOrDigit(ch)))
                {
                    quote = true;
                    break;
                }
            }

            if (!quote && sheet.Length > 0 && char.IsDigit(sheet[0]))
            {
                quote = true;
            }

            return quote ? "'" + sheet.Replace("'", "''") + "'" : sheet;
        }

        private static bool TrySetSeriesProp(object series, string name, object value, out string error)
        {
            error = null;
            if (series == null || string.IsNullOrEmpty(name) || value == null)
            {
                error = "空";
                return false;
            }

            try
            {
                series.GetType().InvokeMember(
                    name,
                    BindingFlags.SetProperty | BindingFlags.Instance | BindingFlags.Public,
                    null,
                    series,
                    new object[] { value });
                return true;
            }
            catch (Exception ex)
            {
                error = FormatComError(ex);
                return false;
            }
        }

        private static void DumpPourState(
            string tag,
            object chart,
            object ws,
            PptHtmlChartGrid grid,
            List<string> warnings)
        {
            int want = grid != null && grid.Rows != null ? grid.Rows.Count : -1;
            PourLog(warnings, tag
                + " 稿行=" + want
                + " 系列点=" + ReadSeriesRowCount(chart)
                + " | " + DescribeLiveSeries(chart)
                + " | " + DescribeListBind(ws)
                + " | " + DescribeSeriesExtra(chart)
                + " | catAxis=[" + string.Join(",", TryReadCategoryAxisNames(chart)) + "]");
        }

        private static string DescribeListBind(object ws)
        {
            if (ws == null)
            {
                return "表=null";
            }

            try
            {
                string used = TryRangeAddress(TryGetProp(ws, "UsedRange"));
                object lists = WppCom.GetProperty(ws, "ListObjects");
                int count = lists == null ? -1 : Convert.ToInt32(WppCom.GetProperty(lists, "Count") ?? -1);
                if (count < 1)
                {
                    return "lists=" + count + " used=" + used;
                }

                object lo = WppCom.GetIndexed(lists, 1);
                object loRange = lo == null ? null : WppCom.GetProperty(lo, "Range");
                object body = lo == null ? null : TryGetProp(lo, "DataBodyRange");
                return "lists=" + count
                    + " lo=" + TryRangeAddress(loRange)
                    + " body=" + TryRangeAddress(body)
                    + " used=" + used
                    + " name=" + (TryPropString(lo, "Name") ?? "?");
            }
            catch (Exception ex)
            {
                return "表读失败: " + ex.Message;
            }
        }

        private static string DescribeSeriesExtra(object chart)
        {
            if (chart == null)
            {
                return "S1=null";
            }

            try
            {
                object s1 = GetSeries(chart, 1);
                if (s1 == null)
                {
                    return "S1=null PlotBy=" + (TryPropString(chart, "PlotBy") ?? "?");
                }

                return "S1 Formula=" + (TryPropString(s1, "Formula") ?? TryPropString(s1, "FormulaLocal") ?? "?")
                    + " Values=" + DescribeComArg(TryGetProp(s1, "Values"))
                    + " XValues=" + DescribeComArg(TryGetProp(s1, "XValues"))
                    + " PlotBy=" + (TryPropString(chart, "PlotBy") ?? "?");
            }
            catch (Exception ex)
            {
                return "S1读失败: " + ex.Message;
            }
        }

        private static string DescribeComArg(object value)
        {
            if (value == null)
            {
                return "null";
            }

            if (value is string s)
            {
                return "str:" + s;
            }

            if (value is Array arr)
            {
                var dims = new List<string>();
                for (int i = 0; i < arr.Rank; i++)
                {
                    dims.Add(arr.GetLength(i).ToString(CultureInfo.InvariantCulture));
                }

                string sample = "";
                try
                {
                    int take = Math.Min(6, arr.Length);
                    var bits = new List<string>();
                    int n = 0;
                    foreach (object item in arr)
                    {
                        if (n >= take)
                        {
                            break;
                        }

                        bits.Add(FormatCell(item));
                        n++;
                    }

                    sample = " [" + string.Join(",", bits) + (arr.Length > take ? ",…" : "") + "]";
                }
                catch (Exception)
                {
                }

                return arr.GetType().Name + " rank=" + arr.Rank + " dim=" + string.Join("x", dims)
                    + " len=" + arr.Length + sample;
            }

            string addr = null;
            try
            {
                addr = TryRangeAddress(value);
            }
            catch (Exception)
            {
            }

            string typeName = value.GetType().Name;
            if (!string.IsNullOrEmpty(addr) && !string.Equals(addr, "ok", StringComparison.Ordinal)
                && !string.Equals(addr, "null", StringComparison.Ordinal))
            {
                return "Range " + addr + " (" + typeName + ")";
            }

            return typeName + "=" + Convert.ToString(value, CultureInfo.InvariantCulture);
        }

        private static string FormatComError(Exception ex)
        {
            if (ex == null)
            {
                return "?";
            }

            var sb = new StringBuilder();
            Exception cur = ex;
            int depth = 0;
            while (cur != null && depth < 4)
            {
                if (depth > 0)
                {
                    sb.Append(" / ");
                }

                sb.Append(cur.GetType().Name).Append(": ").Append(cur.Message);
                if (cur is COMException com)
                {
                    sb.Append(" hr=0x").Append(com.ErrorCode.ToString("X8", CultureInfo.InvariantCulture));
                }

                cur = cur.InnerException;
                depth++;
            }

            return sb.ToString();
        }

        private static object SheetQualifiedAddress(object ws, object range)
        {
            string addr = TryRangeAddress(range);
            if (string.IsNullOrEmpty(addr))
            {
                return null;
            }

            string sheet = TryPropString(ws, "Name") ?? "Sheet1";
            if (addr.IndexOf('!') >= 0)
            {
                return addr;
            }

            return "'" + sheet.Replace("'", "''") + "'!" + addr;
        }

        private static void TryEnsureSheetListObject(object ws, List<string> warnings)
        {
            if (ws == null)
            {
                return;
            }

            try
            {
                object lists = WppCom.GetProperty(ws, "ListObjects");
                int count = lists == null ? 0 : Convert.ToInt32(WppCom.GetProperty(lists, "Count") ?? 0);
                if (count > 0)
                {
                    PourLog(warnings, "表已有 ListObject " + count);
                    return;
                }

                object used = WppCom.GetProperty(ws, "UsedRange")
                    ?? TryGetExcelRange(ws, "A1", "D5");
                if (used == null)
                {
                    PourLog(warnings, "ListObjects.Add 跳过：无 UsedRange");
                    return;
                }

                object lo = WppCom.Invoke(lists, "Add", 1, used, true);
                PourLog(warnings, "ListObjects.Add 默认表 src=" + DescribeComArg(used)
                    + " " + (lo == null ? "null" : "ok") + " | " + DescribeListBind(ws));
            }
            catch (Exception ex)
            {
                PourLog(warnings, "ListObjects.Add 失败: " + FormatComError(ex));
            }
        }

        private static object TryGetListObjectRange(object ws)
        {
            try
            {
                object lists = WppCom.GetProperty(ws, "ListObjects");
                int count = lists == null ? 0 : Convert.ToInt32(WppCom.GetProperty(lists, "Count") ?? 0);
                if (count < 1)
                {
                    return null;
                }

                object lo = WppCom.GetIndexed(lists, 1);
                return lo == null ? null : WppCom.GetProperty(lo, "Range");
            }
            catch (Exception)
            {
                return null;
            }
        }

        private static void TryChartWizard(object chart, object source, List<string> warnings)
        {
            if (chart == null || source == null)
            {
                return;
            }

            try
            {
                object xl = WppCom.GetProperty(chart, "ChartType");
                WppCom.Invoke(chart, "ChartWizard", source, xl, Type.Missing, 2, 1, 1);
                PourLog(warnings, "ChartWizard(source,type,xlColumns) src=" + DescribeComArg(source)
                    + " ok | " + DescribeLiveSeries(chart));
                return;
            }
            catch (Exception ex)
            {
                PourLog(warnings, "ChartWizard 带参失败 src=" + DescribeComArg(source)
                    + " | " + FormatComError(ex));
            }

            try
            {
                WppCom.Invoke(chart, "ChartWizard", source);
                PourLog(warnings, "ChartWizard(source) src=" + DescribeComArg(source)
                    + " ok | " + DescribeLiveSeries(chart));
            }
            catch (Exception ex)
            {
                PourLog(warnings, "ChartWizard 失败: " + FormatComError(ex));
            }
        }

        private static void TryBindSeriesToSheetValue2(
            object chart,
            object ws,
            PptHtmlChartGrid grid,
            List<string> warnings)
        {
            if (chart == null || ws == null || grid == null || grid.Rows == null || grid.Rows.Count < 1)
            {
                return;
            }

            int rows = grid.Rows.Count;
            object xRange = TryGetExcelRange(ws, "A2", "A" + (rows + 1).ToString(CultureInfo.InvariantCulture));
            object xVal = TryRangeValues(xRange);
            int si = 0;
            for (int c = 0; c < grid.Columns.Count; c++)
            {
                if (grid.Columns[c].Role == "category")
                {
                    continue;
                }

                si++;
                object series = GetSeries(chart, si);
                if (series == null)
                {
                    continue;
                }

                string colLetter = ColLetter(c + 1);
                object yRange = TryGetExcelRange(
                    ws,
                    colLetter + "2",
                    colLetter + (rows + 1).ToString(CultureInfo.InvariantCulture));
                object yVal = TryRangeValues(yRange);
                if (yVal == null)
                {
                    PourLog(warnings, "Value2 S" + si + " 无数组 yRange=" + DescribeComArg(yRange));
                    continue;
                }

                PourLog(warnings, "Value2 S" + si + " y=" + DescribeComArg(yVal)
                    + " x=" + DescribeComArg(xVal));
                if (!TryAssignSeriesValues(series, yVal, xVal, out string bindErr))
                {
                    PourLog(warnings, "Value2 S" + si + " 失败: " + bindErr);
                    continue;
                }

                PourLog(warnings, "Value2 S" + si + " ok | " + DescribeLiveSeries(chart)
                    + " | " + DescribeSeriesExtra(chart));
            }
        }

        private static object TryRangeValues(object range)
        {
            if (range == null)
            {
                return null;
            }

            try
            {
                return WppCom.GetProperty(range, "Value2") ?? WppCom.GetProperty(range, "Value");
            }
            catch (Exception)
            {
                try
                {
                    return WppCom.GetProperty(range, "Value");
                }
                catch (Exception)
                {
                    return null;
                }
            }
        }

        private static void TryRebuildSeriesWithArrays(
            object chart,
            PptHtmlChartGrid grid,
            List<string> warnings)
        {
            if (chart == null || grid == null || grid.Rows == null)
            {
                return;
            }

            int want = CountValueColumns(grid);
            if (want < 1)
            {
                return;
            }

            try
            {
                int total = GetSeriesCount(chart);
                for (int i = total; i >= 1; i--)
                {
                    object series = GetSeries(chart, i);
                    if (series != null)
                    {
                        WppCom.Invoke(series, "Delete");
                    }
                }

                PourLog(warnings, "已删系列，准备 NewSeries ×" + want);
            }
            catch (Exception ex)
            {
                PourLog(warnings, "删系列失败: " + ex.Message);
            }

            if (!TryEnsureSeriesCount(chart, want, out string ensureErr))
            {
                PourLog(warnings, "NewSeries 失败: " + (ensureErr ?? "?"));
                return;
            }

            if (TryAssignGridArrays(chart, grid, warnings, out string assignErr))
            {
                PourLog(warnings, "重建系列后 " + DescribeLiveSeries(chart));
                return;
            }

            PourLog(warnings, "重建系列赋数组失败: " + (assignErr ?? "?")
                + " | " + DescribeLiveSeries(chart));
        }

        private static bool TrySetChartSourceData(object chart, object range, List<string> warnings)
        {
            if (chart == null || range == null)
            {
                return false;
            }

            Exception last = null;
            try
            {
                WppCom.Invoke(chart, "SetSourceData", range);
                PourLog(warnings, "SetSourceData(range) src=" + DescribeComArg(range)
                    + " ok | " + DescribeLiveSeries(chart));
                return true;
            }
            catch (Exception ex)
            {
                last = ex;
                PourLog(warnings, "SetSourceData(range) 失败 src=" + DescribeComArg(range)
                    + " | " + FormatComError(ex));
            }

            try
            {
                WppCom.Invoke(chart, "SetSourceData", range, 2);
                PourLog(warnings, "SetSourceData(range,xlColumns) src=" + DescribeComArg(range)
                    + " ok | " + DescribeLiveSeries(chart));
                return true;
            }
            catch (Exception ex)
            {
                last = ex;
                PourLog(warnings, "SetSourceData(range,xlColumns) 失败 | " + FormatComError(ex));
            }

            string addr = range as string ?? TryRangeAddress(range);
            if (!string.IsNullOrEmpty(addr) && !string.Equals(addr, "ok", StringComparison.Ordinal))
            {
                try
                {
                    WppCom.Invoke(chart, "SetSourceData", addr);
                    PourLog(warnings, "SetSourceData(addr) src=" + addr
                        + " ok | " + DescribeLiveSeries(chart));
                    return true;
                }
                catch (Exception ex)
                {
                    last = ex;
                    PourLog(warnings, "SetSourceData(addr) 失败 src=" + addr
                        + " | " + FormatComError(ex));
                }
            }

            PourLog(warnings, "SetSourceData 失败: " + FormatComError(last));
            return false;
        }

        private static void TryBindSeriesToSheetRanges(
            object chart,
            object ws,
            PptHtmlChartGrid grid,
            List<string> warnings)
        {
            if (chart == null || ws == null || grid == null || grid.Rows == null || grid.Rows.Count < 1)
            {
                return;
            }

            int rows = grid.Rows.Count;
            object xRange = TryGetExcelRange(ws, "A2", "A" + (rows + 1).ToString(CultureInfo.InvariantCulture));
            int si = 0;
            for (int c = 0; c < grid.Columns.Count; c++)
            {
                if (grid.Columns[c].Role == "category")
                {
                    continue;
                }

                si++;
                object series = GetSeries(chart, si);
                if (series == null)
                {
                    PourLog(warnings, "绑 Range 无系列 " + si);
                    continue;
                }

                string colLetter = ColLetter(c + 1);
                object yRange = TryGetExcelRange(
                    ws,
                    colLetter + "2",
                    colLetter + (rows + 1).ToString(CultureInfo.InvariantCulture));
                if (yRange == null)
                {
                    PourLog(warnings, "绑 Range 无 Y " + colLetter);
                    continue;
                }

                PourLog(warnings, "绑 Range S" + si + " y=" + DescribeComArg(yRange)
                    + " x=" + DescribeComArg(xRange));
                if (!TryAssignSeriesValues(series, yRange, xRange, out string bindErr))
                {
                    PourLog(warnings, "绑 Range S" + si + " 失败: " + bindErr);
                    continue;
                }

                PourLog(warnings, "绑 Range S" + si + " ok | " + DescribeLiveSeries(chart)
                    + " | " + DescribeSeriesExtra(chart));
            }
        }

        private static bool TryAssignSeriesValues(
            object series,
            object yRange,
            object xRange,
            out string error)
        {
            error = null;
            if (series == null || yRange == null)
            {
                error = "series/y 为空";
                return false;
            }

            try
            {
                series.GetType().InvokeMember(
                    "Values",
                    BindingFlags.SetProperty | BindingFlags.Instance | BindingFlags.Public,
                    null,
                    series,
                    new object[] { yRange });
            }
            catch (Exception ex)
            {
                error = "Values=" + DescribeComArg(yRange) + " | " + FormatComError(ex);
                return false;
            }

            if (xRange == null)
            {
                return true;
            }

            try
            {
                series.GetType().InvokeMember(
                    "XValues",
                    BindingFlags.SetProperty | BindingFlags.Instance | BindingFlags.Public,
                    null,
                    series,
                    new object[] { xRange });
            }
            catch (Exception ex)
            {
                PourLog(null, "XValues=" + DescribeComArg(xRange) + " 失败: " + FormatComError(ex));
            }

            return true;
        }

        private static void TryClearSheetBeyond(object ws, int lastRow, int lastCol, List<string> warnings)
        {
            if (ws == null)
            {
                return;
            }

            try
            {
                object used = WppCom.GetProperty(ws, "UsedRange");
                if (used == null)
                {
                    return;
                }

                int usedRows = Convert.ToInt32(WppCom.GetProperty(WppCom.GetProperty(used, "Rows"), "Count"));
                int usedCols = Convert.ToInt32(WppCom.GetProperty(WppCom.GetProperty(used, "Columns"), "Count"));
                int maxR = Math.Max(usedRows, lastRow);
                int maxC = Math.Max(usedCols, lastCol);
                bool cleared = false;
                for (int r = 1; r <= maxR; r++)
                {
                    for (int c = 1; c <= maxC; c++)
                    {
                        if (r <= lastRow && c <= lastCol)
                        {
                            continue;
                        }

                        SetCell(ws, r, c, "");
                        cleared = true;
                    }
                }

                if (cleared)
                {
                    PourLog(warnings, "已清表外残留 used=" + usedRows + "x" + usedCols
                        + " keep=" + lastRow + "x" + lastCol);
                }
            }
            catch (Exception ex)
            {
                PourLog(warnings, "清表外残留失败: " + ex.Message);
            }
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
