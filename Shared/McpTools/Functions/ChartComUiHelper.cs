using System;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    /// <summary>
    /// 图表 COM 操作期间抑制 Word/内嵌 Excel 弹窗与抢焦点（避免 apply/create 卡住）。
    /// </summary>
    public sealed class ChartComUiState
    {
        public bool WordScreenUpdating { get; set; } = true;
        public Word.WdAlertLevel WordAlerts { get; set; } = Word.WdAlertLevel.wdAlertsAll;
        public bool ExcelVisible { get; set; } = true;
        public bool ExcelScreenUpdating { get; set; } = true;
        public bool ExcelDisplayAlerts { get; set; } = true;
        /// <summary>BreakLink 后跳过 Workbook 访问，避免再次触发「链接的文件不可用」。</summary>
        public bool ExcelSuppressSkipped { get; set; }
    }

    public static class ChartComUiHelper
    {
        [ThreadStatic]
        private static bool _chartLinkBrokenOnThread;

        public static ChartComUiState EnterBatch(Word.Application app, Word.Chart chart = null)
        {
            var state = new ChartComUiState();
            bool linkBrokenBeforeBatch = _chartLinkBrokenOnThread;
            if (app == null)
            {
                return state;
            }

            try
            {
                state.WordScreenUpdating = app.ScreenUpdating;
                state.WordAlerts = app.DisplayAlerts;
                app.ScreenUpdating = false;
                app.DisplayAlerts = Word.WdAlertLevel.wdAlertsNone;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[ChartComUi] Word suppress: {ex.Message}");
            }

            TryExitChartDataEditMode(app);
            if (TryBreakExternalChartLink(chart))
            {
                state.ExcelSuppressSkipped = true;
            }

            if (!state.ExcelSuppressSkipped && !linkBrokenBeforeBatch && !_chartLinkBrokenOnThread)
            {
                TrySuppressEmbeddedExcel(chart, state);
            }
            else if (state.ExcelSuppressSkipped || linkBrokenBeforeBatch || _chartLinkBrokenOnThread)
            {
                state.ExcelSuppressSkipped = true;
                System.Diagnostics.Debug.WriteLine("[ChartComUi] 已 BreakLink，跳过 Workbook 访问");
            }

            return state;
        }

        public static void ExitBatch(Word.Application app, ChartComUiState state, Word.Chart chart = null)
        {
            if (app == null || state == null)
            {
                return;
            }

            if (!state.ExcelSuppressSkipped)
            {
                try
                {
                    // 仅恢复告警/刷新；可见性始终保持隐藏，避免建图后 Excel 窗口重新弹出
                    TryRestoreEmbeddedExcel(chart, state, restoreVisible: false);
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"[ChartComUi] restore excel: {ex.Message}");
                }
            }

            // AddChart / 写数据常会再次激活内嵌 Excel，退出批处理时强制隐藏
            TryHideEmbeddedExcel(chart);

            try
            {
                app.DisplayAlerts = state.WordAlerts;
                app.ScreenUpdating = state.WordScreenUpdating;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[ChartComUi] restore word: {ex.Message}");
            }

            TryExitChartDataEditMode(app);
            _chartLinkBrokenOnThread = false;
        }

        /// <summary>
        /// 退出图表数据编辑模式，回到主文档视图。
        /// </summary>
        public static void TryExitChartDataEditMode(Word.Application app)
        {
            if (app == null)
            {
                return;
            }

            try
            {
                if (app.ActiveWindow != null && app.ActiveWindow.View != null)
                {
                    app.ActiveWindow.View.SeekView = Word.WdSeekView.wdSeekMainDocument;
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[ChartComUi] SeekView: {ex.Message}");
            }
        }

        /// <summary>
        /// 断开图表对外部 Excel 的链接，避免访问 ChartData 时弹出「链接的文件不可用」并阻塞自动化。
        /// </summary>
        /// <returns>是否执行了 BreakLink。</returns>
        public static bool TryBreakExternalChartLink(Word.Chart chart)
        {
            if (chart == null)
            {
                return false;
            }

            try
            {
                dynamic chartData = chart.ChartData;
                if (chartData == null)
                {
                    return false;
                }

                bool isLinked = false;
                try
                {
                    isLinked = chartData.IsLinked;
                }
                catch
                {
                    // IsLinked 读取失败时仍尝试 BreakLink
                }

                if (!isLinked)
                {
                    return false;
                }

                System.Diagnostics.Debug.WriteLine("[ChartComUi] 检测到图表外链数据，执行 BreakLink");
                chartData.BreakLink();
                System.Diagnostics.Debug.WriteLine("[ChartComUi] BreakLink 完成");
                _chartLinkBrokenOnThread = true;
                return true;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[ChartComUi] BreakLink: {ex.Message}");
                return false;
            }
        }

        public static void TrySuppressEmbeddedExcel(Word.Chart chart, ChartComUiState state = null)
        {
            if (chart == null)
            {
                return;
            }

            try
            {
                if (!TryGetEmbeddedExcelApp(chart, out dynamic excelApp))
                {
                    return;
                }

                if (state != null)
                {
                    try
                    {
                        state.ExcelVisible = excelApp.Visible;
                        state.ExcelScreenUpdating = excelApp.ScreenUpdating;
                        state.ExcelDisplayAlerts = excelApp.DisplayAlerts;
                    }
                    catch
                    {
                        // 读取原状态失败时仍尝试抑制
                    }

                    // 自动化结束后不应再把 Excel 窗口恢复为显示
                    state.ExcelVisible = false;
                }

                HideExcelApp(excelApp);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[ChartComUi] suppress excel: {ex.Message}");
            }
        }

        /// <summary>
        /// 强制隐藏内嵌图表关联的 Excel（不退出进程，避免破坏 Word 内嵌数据源）。
        /// </summary>
        public static void TryHideEmbeddedExcel(Word.Chart chart)
        {
            if (chart == null)
            {
                return;
            }

            try
            {
                if (!TryGetEmbeddedExcelApp(chart, out dynamic excelApp))
                {
                    return;
                }

                HideExcelApp(excelApp);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[ChartComUi] hide excel: {ex.Message}");
            }
        }

        private static void TryRestoreEmbeddedExcel(Word.Chart chart, ChartComUiState state, bool restoreVisible = true)
        {
            if (chart == null || state == null)
            {
                return;
            }

            try
            {
                if (!TryGetEmbeddedExcelApp(chart, out dynamic excelApp))
                {
                    return;
                }

                excelApp.DisplayAlerts = state.ExcelDisplayAlerts;
                excelApp.ScreenUpdating = state.ExcelScreenUpdating;
                if (restoreVisible)
                {
                    excelApp.Visible = state.ExcelVisible;
                }
                else
                {
                    excelApp.Visible = false;
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[ChartComUi] restore excel inner: {ex.Message}");
            }
        }

        private static bool TryGetEmbeddedExcelApp(Word.Chart chart, out dynamic excelApp)
        {
            excelApp = null;
            try
            {
                dynamic chartData = chart.ChartData;
                if (chartData == null)
                {
                    return false;
                }

                dynamic workbook = chartData.Workbook;
                if (workbook == null)
                {
                    return false;
                }

                excelApp = workbook.Application;
                return excelApp != null;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[ChartComUi] get excel app: {ex.Message}");
                return false;
            }
        }

        private static void HideExcelApp(dynamic excelApp)
        {
            if (excelApp == null)
            {
                return;
            }

            try
            {
                excelApp.DisplayAlerts = false;
            }
            catch
            {
                // ignore
            }

            try
            {
                excelApp.ScreenUpdating = false;
            }
            catch
            {
                // ignore
            }

            try
            {
                excelApp.Visible = false;
            }
            catch
            {
                // ignore
            }

            // 尽量关掉已打开的工作簿窗口，避免任务栏仍显示 Excel
            try
            {
                dynamic windows = excelApp.Windows;
                if (windows != null)
                {
                    int count = windows.Count;
                    for (int i = count; i >= 1; i--)
                    {
                        try
                        {
                            windows[i].Visible = false;
                        }
                        catch
                        {
                            // ignore single window
                        }
                    }
                }
            }
            catch
            {
                // ignore
            }
        }
    }
}
