using System;
using System.Diagnostics;
using WordAddIn1.BrowserHost;
using WordAddIn1.HostPlatform;
using Office = Microsoft.Office.Core;
using Excel = Microsoft.Office.Interop.Excel;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1.OpenFiles
{
    /// <summary>
    /// 侧栏「打开文件」行尾新建：文档只 Add 不存盘；Chrome / Edge 拉新窗口。
    /// </summary>
    public static class OpenFilesSidebarCreate
    {
        public static bool TryCreate(string group, out string error)
        {
            error = null;
            string key = (group ?? "").Trim().ToLowerInvariant();
            if (string.IsNullOrEmpty(key))
            {
                error = "缺少分类";
                return false;
            }

            if (string.Equals(key, "yiwrite", StringComparison.OrdinalIgnoreCase))
            {
                error = "易写浏览器请走空白标签";
                return false;
            }

            if (string.Equals(key, "chrome", StringComparison.OrdinalIgnoreCase)
                || string.Equals(key, "edge", StringComparison.OrdinalIgnoreCase))
            {
                return TryLaunchBrowser(key, out error);
            }

            if (OfficeStaScheduler.ShouldHop)
            {
                string err = null;
                bool ok = OfficeStaScheduler.Invoke(() => TryCreateDocumentOnSta(key, out err));
                error = err;
                return ok;
            }

            return TryCreateDocumentOnSta(key, out error);
        }

        private static bool TryCreateDocumentOnSta(string key, out string error)
        {
            error = null;
            try
            {
                if (key == "word")
                {
                    return TryCreateWord(out error);
                }

                if (key == "wps")
                {
                    return TryCreateWps(out error);
                }

                if (key == "excel")
                {
                    return TryCreateExcel(out error);
                }

                if (key == "et")
                {
                    return TryCreateEt(out error);
                }

                if (key == "ppt")
                {
                    return TryCreatePpt(out error);
                }

                if (key == "wpp")
                {
                    return TryCreateWpp(out error);
                }

                error = "不支持的分类";
                return false;
            }
            catch (Exception ex)
            {
                error = "新建失败: " + ex.Message;
                return false;
            }
        }

        private static bool TryCreateWord(out string error)
        {
            if (!WordApplicationResolver.TryResolve(
                    null,
                    out Word.Application app,
                    out error,
                    createIfMissing: true))
            {
                return false;
            }

            Word.Document doc = app.Documents.Add();
            try
            {
                app.Visible = true;
            }
            catch (Exception)
            {
            }

            try
            {
                app.Activate();
            }
            catch (Exception)
            {
            }

            try
            {
                doc.Activate();
            }
            catch (Exception)
            {
            }

            MaximizeWord(app, doc);
            HostCallbacks.RaiseWordApplicationResolved(app);
            HostCallbacks.RaiseOpenFilesRefresh();
            return true;
        }

        private static bool TryCreateWps(out string error)
        {
            if (!WpsApplicationResolver.TryResolve(out object app, out error, createIfMissing: true))
            {
                return false;
            }

            object documents = WpsCom.GetProperty(app, "Documents");
            object doc = WpsCom.Invoke(documents, "Add");
            if (doc == null)
            {
                error = "WPS 文字新建失败";
                return false;
            }

            TryInvoke(doc, "Activate");
            WpsCom.TrySetProperty(app, "Visible", true);
            TryInvoke(app, "Activate");
            MaximizeLateBound(app, WordMaximizeState);
            HostCallbacks.RaiseOpenFilesRefresh();
            return true;
        }

        private static bool TryCreateExcel(out string error)
        {
            if (!ExcelApplicationResolver.TryResolveTyped(
                    out Excel.Application app,
                    out error,
                    createIfMissing: true))
            {
                return false;
            }

            Excel.Workbook book = app.Workbooks.Add();
            ExcelApplicationResolver.EnsureVisible(app);
            try
            {
                book.Activate();
            }
            catch (Exception)
            {
            }

            MaximizeExcel(app);
            HostCallbacks.RaiseExcelApplicationResolved(app);
            HostCallbacks.RaiseOpenFilesRefresh();
            return true;
        }

        private static bool TryCreateEt(out string error)
        {
            if (!EtApplicationResolver.TryResolve(out object app, out error, createIfMissing: true))
            {
                return false;
            }

            object books = EtCom.GetProperty(app, "Workbooks");
            object book = EtCom.Invoke(books, "Add");
            if (book == null)
            {
                error = "WPS 表格新建失败";
                return false;
            }

            EtCom.EnsureVisible(app);
            TryInvoke(book, "Activate");
            MaximizeLateBound(app, ExcelMaximizeState);
            HostCallbacks.RaiseOpenFilesRefresh();
            return true;
        }

        private static bool TryCreatePpt(out string error)
        {
            if (!PowerPointApplicationResolver.TryResolveTyped(
                    out PowerPoint.Application app,
                    out error,
                    createIfMissing: true))
            {
                return false;
            }

            PowerPoint.Presentation presentation = app.Presentations.Add(Office.MsoTriState.msoTrue);
            PowerPointApplicationResolver.EnsureVisible(app);
            try
            {
                if (presentation != null)
                {
                    PowerPointApplicationResolver.EnsurePresentationHasWindow(presentation);
                }
            }
            catch (Exception)
            {
            }

            MaximizePpt(app);
            HostCallbacks.RaisePowerPointApplicationResolved(app);
            HostCallbacks.RaiseOpenFilesRefresh();
            return true;
        }

        private static bool TryCreateWpp(out string error)
        {
            if (!WppApplicationResolver.TryResolve(out object app, out error, createIfMissing: true))
            {
                return false;
            }

            object presentations = WppCom.GetProperty(app, "Presentations");
            object presentation = TryAddWpp(presentations);
            if (presentation == null)
            {
                error = "WPS 演示新建失败";
                return false;
            }

            WppCom.EnsureVisible(app);
            MaximizeLateBound(app, PptMaximizeState);
            HostCallbacks.RaiseOpenFilesRefresh();
            return true;
        }

        private static object TryAddWpp(object presentations)
        {
            if (presentations == null)
            {
                return null;
            }

            try
            {
                return WppCom.Invoke(presentations, "Add", true);
            }
            catch (Exception)
            {
            }

            try
            {
                return WppCom.Invoke(presentations, "Add");
            }
            catch (Exception)
            {
                return null;
            }
        }

        private static bool TryLaunchBrowser(string kind, out string error)
        {
            error = null;
            string exe = ExtensionHost.ResolveBrowserExe(kind);
            if (string.IsNullOrEmpty(exe))
            {
                error = string.Equals(kind, "edge", StringComparison.OrdinalIgnoreCase)
                    ? "本机未找到 Edge"
                    : "本机未找到 Chrome";
                return false;
            }

            try
            {
                Process.Start(new ProcessStartInfo
                {
                    FileName = exe,
                    Arguments = "--new-window",
                    UseShellExecute = false
                });
                return true;
            }
            catch (Exception ex)
            {
                error = "启动 " + (kind == "edge" ? "Edge" : "Chrome") + " 失败: " + ex.Message;
                return false;
            }
        }

        private const int WordMaximizeState = 1;
        private const int ExcelMaximizeState = -4137;
        private const int PptMaximizeState = 3;

        private static void MaximizeWord(Word.Application app, Word.Document doc)
        {
            try
            {
                app.WindowState = Word.WdWindowState.wdWindowStateMaximize;
            }
            catch (Exception)
            {
            }

            try
            {
                if (doc?.ActiveWindow != null)
                {
                    doc.ActiveWindow.WindowState = Word.WdWindowState.wdWindowStateMaximize;
                }
            }
            catch (Exception)
            {
            }

            NativeWindowActivate.MaximizeAndFocus(TryWordHwnd(app, doc));
        }

        private static void MaximizeExcel(Excel.Application app)
        {
            try
            {
                app.WindowState = Excel.XlWindowState.xlMaximized;
            }
            catch (Exception)
            {
            }

            try
            {
                if (app.ActiveWindow != null)
                {
                    app.ActiveWindow.WindowState = Excel.XlWindowState.xlMaximized;
                }
            }
            catch (Exception)
            {
            }

            int hwnd = 0;
            try
            {
                hwnd = app.Hwnd;
            }
            catch (Exception)
            {
            }

            NativeWindowActivate.MaximizeAndFocus(hwnd);
        }

        private static void MaximizePpt(PowerPoint.Application app)
        {
            try
            {
                app.WindowState = PowerPoint.PpWindowState.ppWindowMaximized;
            }
            catch (Exception)
            {
            }

            int hwnd = 0;
            try
            {
                hwnd = app.HWND;
            }
            catch (Exception)
            {
            }

            NativeWindowActivate.MaximizeAndFocus(hwnd);
        }

        private static void MaximizeLateBound(object app, int windowState)
        {
            if (app == null)
            {
                return;
            }

            WpsCom.TrySetProperty(app, "WindowState", windowState);
            object activeWindow = WpsCom.GetProperty(app, "ActiveWindow");
            if (activeWindow != null)
            {
                WpsCom.TrySetProperty(activeWindow, "WindowState", windowState);
            }

            NativeWindowActivate.MaximizeAndFocus(TryLateBoundHwnd(app, activeWindow));
        }

        private static int TryWordHwnd(Word.Application app, Word.Document doc)
        {
            try
            {
                int hwnd = doc != null ? doc.ActiveWindow.Hwnd : 0;
                if (hwnd != 0)
                {
                    return hwnd;
                }
            }
            catch (Exception)
            {
            }

            try
            {
                return app != null ? app.ActiveWindow.Hwnd : 0;
            }
            catch (Exception)
            {
                return 0;
            }
        }

        private static int TryLateBoundHwnd(object app, object activeWindow)
        {
            if (TryReadHwnd(activeWindow, out int hwnd) && hwnd != 0)
            {
                return hwnd;
            }

            TryReadHwnd(app, out hwnd);
            return hwnd;
        }

        private static bool TryReadHwnd(object target, out int hwnd)
        {
            hwnd = 0;
            if (target == null)
            {
                return false;
            }

            if (EtCom.TryGetHwnd(target, out hwnd) && hwnd != 0)
            {
                return true;
            }

            if (WppCom.TryGetHwnd(target, out hwnd) && hwnd != 0)
            {
                return true;
            }

            try
            {
                object raw = WpsCom.GetProperty(target, "Hwnd") ?? WpsCom.GetProperty(target, "HWND");
                if (raw != null)
                {
                    hwnd = Convert.ToInt32(raw);
                    return hwnd != 0;
                }
            }
            catch (Exception)
            {
            }

            hwnd = 0;
            return false;
        }

        private static void TryInvoke(object target, string name)
        {
            if (target == null)
            {
                return;
            }

            try
            {
                WpsCom.Invoke(target, name);
            }
            catch (Exception)
            {
            }
        }
    }
}
