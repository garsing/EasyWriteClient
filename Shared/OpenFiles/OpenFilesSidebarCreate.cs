using System;
using System.Diagnostics;
using WordAddIn1.BrowserHost;
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
