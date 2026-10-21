using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using WordAddIn1;

namespace WordDocumentRoundtripTest
{
    /// <summary>Excel / PPT 检查点短套：经 Desktop 管道调 F_*，COM 只用来关稿和造「不建副本」的改动。</summary>
    internal sealed class OfficeCkptSession
    {
        public OfficeCkptSession(string host)
        {
            Host = NormalizeHost(host);
            App = Host;
            ExpectedKind = Host;
        }

        public string Host { get; }

        public string App { get; }

        public string ExpectedKind { get; }

        public bool IsExcelFamily
        {
            get { return Host == "excel" || Host == "et"; }
        }

        public bool IsEt
        {
            get { return Host == "et"; }
        }

        public bool IsWpp
        {
            get { return Host == "wpp"; }
        }

        public string ProcName
        {
            get
            {
                if (Host == "excel")
                {
                    return "EXCEL";
                }

                if (Host == "et")
                {
                    return "et";
                }

                if (Host == "wpp")
                {
                    return "wpp";
                }

                return "POWERPNT";
            }
        }

        public string NewPath()
        {
            string ext = IsExcelFamily ? ".xlsx" : ".pptx";
            string tag = IsExcelFamily ? "ew-xlk-" : "ew-pk-";
            return Path.Combine(Path.GetTempPath(), tag + Guid.NewGuid().ToString("N") + ext);
        }

        public ToolResult Open(string path)
        {
            return DesktopToolPipe.Call("F_open_document", new Dictionary<string, object>
            {
                ["path"] = path,
                ["create_blank"] = true,
                ["app"] = App
            });
        }

        public ToolResult ListOperations(string channelId)
        {
            var args = new Dictionary<string, object>();
            if (channelId != null)
            {
                args["channel_id"] = channelId;
            }

            return DesktopToolPipe.Call("F_list_document_operations", args);
        }

        public ToolResult RestoreCheckpoint(string channelId, string timestamp)
        {
            var args = new Dictionary<string, object>();
            if (channelId != null)
            {
                args["channel_id"] = channelId;
            }

            if (timestamp != null)
            {
                args["timestamp"] = timestamp;
            }

            return DesktopToolPipe.Call("F_restore_document_checkpoint", args);
        }

        public ToolResult Overview(string channelId)
        {
            var args = new Dictionary<string, object>();
            if (!string.IsNullOrEmpty(channelId))
            {
                args["channel_id"] = channelId;
            }

            return DesktopToolPipe.Call(
                IsExcelFamily ? "F_get_workbook_content" : "F_get_presentation_content",
                args);
        }

        public ToolResult AddSheet(string channelId, string name)
        {
            return DesktopToolPipe.Call("F_manage_excel_sheet", new Dictionary<string, object>
            {
                ["channel_id"] = channelId,
                ["action"] = "add",
                ["name"] = name
            });
        }

        public ToolResult AddSlide(string channelId)
        {
            return DesktopToolPipe.Call("F_manage_ppt_slide", new Dictionary<string, object>
            {
                ["channel_id"] = channelId,
                ["action"] = "add"
            });
        }

        public ToolResult MutatingContractFail(string channelId)
        {
            string tool = IsExcelFamily ? "F_manage_excel_sheet" : "F_manage_ppt_slide";
            return DesktopToolPipe.Call(tool, new Dictionary<string, object>
            {
                ["channel_id"] = channelId
            });
        }

        public bool OverviewHas(string channelId, string needle)
        {
            ToolResult r = Overview(channelId);
            if (r == null || !r.Success)
            {
                return false;
            }

            return CkptCaseUtil.BlobHas(r, needle);
        }

        public int SlideCount(string channelId)
        {
            return WdSession.IntOf(Overview(channelId), "slide_count");
        }

        public bool ComAddSheet(string path, string name)
        {
            for (int i = 0; i < 8; i++)
            {
                if (IsEt ? EtAddSheet(path, name) : ExcelAddSheet(path, name))
                {
                    return true;
                }

                System.Threading.Thread.Sleep(200);
            }

            return false;
        }

        public bool ComAddSlide(string path)
        {
            for (int i = 0; i < 8; i++)
            {
                if (IsWpp ? WppAddSlide(path) : PptAddSlide(path))
                {
                    return true;
                }

                System.Threading.Thread.Sleep(200);
            }

            return false;
        }

        public void CloseByPath(string path)
        {
            if (string.IsNullOrEmpty(path))
            {
                return;
            }

            try
            {
                if (IsExcelFamily)
                {
                    CloseLate(path, IsEt);
                }
                else
                {
                    ClosePresLate(path, IsWpp);
                }
            }
            catch
            {
            }

            try
            {
                if (File.Exists(path))
                {
                    File.Delete(path);
                }
            }
            catch
            {
            }
        }

        public static string NormalizeExcelHost(string host)
        {
            if (string.Equals(host, "et", StringComparison.OrdinalIgnoreCase)
                || string.Equals(host, "wps", StringComparison.OrdinalIgnoreCase))
            {
                return "et";
            }

            return "excel";
        }

        public static string NormalizePptHost(string host)
        {
            if (string.Equals(host, "wpp", StringComparison.OrdinalIgnoreCase)
                || string.Equals(host, "wps", StringComparison.OrdinalIgnoreCase))
            {
                return "wpp";
            }

            return "ppt";
        }

        private static string NormalizeHost(string host)
        {
            if (string.Equals(host, "et", StringComparison.OrdinalIgnoreCase)
                || string.Equals(host, "wpp", StringComparison.OrdinalIgnoreCase)
                || string.Equals(host, "ppt", StringComparison.OrdinalIgnoreCase)
                || string.Equals(host, "excel", StringComparison.OrdinalIgnoreCase))
            {
                return host.ToLowerInvariant();
            }

            return host ?? "";
        }

        private static bool SamePath(string a, string b)
        {
            if (string.IsNullOrEmpty(a) || string.IsNullOrEmpty(b))
            {
                return false;
            }

            try
            {
                return string.Equals(
                    Path.GetFullPath(a).TrimEnd('\\', '/'),
                    Path.GetFullPath(b).TrimEnd('\\', '/'),
                    StringComparison.OrdinalIgnoreCase);
            }
            catch
            {
                return string.Equals(a, b, StringComparison.OrdinalIgnoreCase);
            }
        }

        private static bool ExcelAddSheet(string path, string name)
        {
            return AddSheetLate(path, name, false);
        }

        private static bool EtAddSheet(string path, string name)
        {
            return AddSheetLate(path, name, true);
        }

        private static bool PptAddSlide(string path)
        {
            return AddSlideLate(path, false);
        }

        private static bool WppAddSlide(string path)
        {
            return AddSlideLate(path, true);
        }

        private static bool AddSheetLate(string path, string name, bool et)
        {
            try
            {
                object app;
                string err;
                bool ok = et
                    ? EtApplicationResolver.TryResolve(out app, out err, false)
                    : ExcelApplicationResolver.TryResolve(out app, out err, false);
                if (!ok)
                {
                    return false;
                }

                object book = FindLateBook(app, "Workbooks", path);
                if (book == null)
                {
                    return false;
                }

                object sheets = Late.Get(book, "Worksheets");
                object ws = Late.Call(sheets, "Add");
                Late.Set(ws, "Name", name);
                return true;
            }
            catch
            {
                return false;
            }
        }

        private static void CloseLate(string path, bool et)
        {
            try
            {
                object app;
                string err;
                bool ok = et
                    ? EtApplicationResolver.TryResolve(out app, out err, false)
                    : ExcelApplicationResolver.TryResolve(out app, out err, false);
                if (!ok)
                {
                    return;
                }

                object book = FindLateBook(app, "Workbooks", path);
                if (book != null)
                {
                    Late.Call(book, "Close", false);
                }
            }
            catch
            {
            }
        }

        private static bool AddSlideLate(string path, bool wpp)
        {
            try
            {
                object app;
                string err;
                bool ok = wpp
                    ? WppApplicationResolver.TryResolve(out app, out err, false)
                    : PowerPointApplicationResolver.TryResolve(out app, out err, false);
                if (!ok)
                {
                    return false;
                }

                object pres = FindLateBook(app, "Presentations", path);
                if (pres == null)
                {
                    return false;
                }

                object slides = Late.Get(pres, "Slides");
                int n = Convert.ToInt32(Late.Get(slides, "Count"));
                Late.Call(slides, "Add", n + 1, 2);
                return true;
            }
            catch
            {
                return false;
            }
        }

        private static void ClosePresLate(string path, bool wpp)
        {
            try
            {
                object app;
                string err;
                bool ok = wpp
                    ? WppApplicationResolver.TryResolve(out app, out err, false)
                    : PowerPointApplicationResolver.TryResolve(out app, out err, false);
                if (!ok)
                {
                    return;
                }

                object pres = FindLateBook(app, "Presentations", path);
                if (pres != null)
                {
                    Late.Call(pres, "Close");
                }
            }
            catch
            {
            }
        }

        private static object FindLateBook(object app, string collectionName, string path)
        {
            object col = Late.Get(app, collectionName);
            int n = Convert.ToInt32(Late.Get(col, "Count"));
            for (int i = 1; i <= n; i++)
            {
                object item = Late.Item(col, i);
                string full = Convert.ToString(Late.Get(item, "FullName") ?? "");
                if (SamePath(full, path))
                {
                    return item;
                }
            }

            return null;
        }

        private static class Late
        {
            public static object Get(object target, string name)
            {
                return target.GetType().InvokeMember(
                    name,
                    BindingFlags.GetProperty,
                    null,
                    target,
                    null);
            }

            public static void Set(object target, string name, object value)
            {
                target.GetType().InvokeMember(
                    name,
                    BindingFlags.SetProperty,
                    null,
                    target,
                    new[] { value });
            }

            public static object Call(object target, string name, params object[] args)
            {
                return target.GetType().InvokeMember(
                    name,
                    BindingFlags.InvokeMethod,
                    null,
                    target,
                    args == null || args.Length == 0 ? null : args);
            }

            public static object Item(object target, int index)
            {
                try
                {
                    return target.GetType().InvokeMember(
                        "Item",
                        BindingFlags.GetProperty,
                        null,
                        target,
                        new object[] { index });
                }
                catch
                {
                    return target.GetType().InvokeMember(
                        "Item",
                        BindingFlags.InvokeMethod,
                        null,
                        target,
                        new object[] { index });
                }
            }
        }
    }
}
