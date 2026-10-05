using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using WordAddIn1;
using Excel = Microsoft.Office.Interop.Excel;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Word = Microsoft.Office.Interop.Word;

namespace WordDocumentRoundtripTest
{
    internal sealed class WdSession
    {
        public WdSession(string host)
        {
            Host = string.Equals(host, "wps", StringComparison.OrdinalIgnoreCase) ? "wps" : "word";
            App = Host;
            ExpectedKind = Host;
        }

        public string Host { get; }

        public string App { get; }

        public string ExpectedKind { get; }

        public bool IsWps
        {
            get { return Host == "wps"; }
        }

        public string NewDocxPath()
        {
            return Path.Combine(
                Path.GetTempPath(),
                "ew-wd-" + Guid.NewGuid().ToString("N") + ".docx");
        }

        public static string DummyPath(string ext)
        {
            string path = Path.Combine(
                Path.GetTempPath(),
                "ew-wd-dummy-" + Guid.NewGuid().ToString("N") + ext);
            File.WriteAllBytes(path, new byte[] { 0 });
            return path;
        }

        public ToolResult Open(string path, bool? createBlank = null, string app = null)
        {
            var args = new Dictionary<string, object>();
            if (path != null)
            {
                args["path"] = path;
            }

            if (createBlank.HasValue)
            {
                args["create_blank"] = createBlank.Value;
            }

            if (app != null)
            {
                args["app"] = app;
            }

            return DesktopToolPipe.Call("F_open_document", args);
        }

        public ToolResult PageInfo(string channelId)
        {
            var args = new Dictionary<string, object>();
            if (channelId != null)
            {
                args["channel_id"] = channelId;
            }

            return DesktopToolPipe.Call("F_get_document_page_info", args);
        }

        public static bool IsShortId(string id)
        {
            return !string.IsNullOrEmpty(id)
                && Regex.IsMatch(id, @"^w\d+$", RegexOptions.CultureInvariant);
        }

        public static Dictionary<string, object> Data(ToolResult r)
        {
            return r == null ? null : r.Data as Dictionary<string, object>;
        }

        public static string Str(ToolResult r, string key)
        {
            Dictionary<string, object> d = Data(r);
            if (d == null || !d.ContainsKey(key) || d[key] == null)
            {
                return "";
            }

            return Convert.ToString(d[key]) ?? "";
        }

        public static bool Bool(ToolResult r, string key)
        {
            Dictionary<string, object> d = Data(r);
            if (d == null || !d.ContainsKey(key) || d[key] == null)
            {
                return false;
            }

            object v = d[key];
            if (v is bool b)
            {
                return b;
            }

            return string.Equals(Convert.ToString(v), "true", StringComparison.OrdinalIgnoreCase);
        }

        public static int IntOf(ToolResult r, string key)
        {
            Dictionary<string, object> d = Data(r);
            if (d == null || !d.ContainsKey(key) || d[key] == null)
            {
                return 0;
            }

            try
            {
                return Convert.ToInt32(d[key]);
            }
            catch
            {
                return 0;
            }
        }

        public static bool HasKey(ToolResult r, string key)
        {
            Dictionary<string, object> d = Data(r);
            return d != null && d.ContainsKey(key);
        }

        public bool InsertPageBreak(string fullPath)
        {
            return IsWps ? InsertPageBreakWps(fullPath) : InsertPageBreakWord(fullPath);
        }

        public bool CloseDocByPath(string fullPath)
        {
            return IsWps ? CloseWpsByPath(fullPath) : CloseWordByPath(fullPath);
        }

        public static bool CloseExcelByPath(string fullPath)
        {
            try
            {
                var app = (Excel.Application)Marshal.GetActiveObject("Excel.Application");
                string norm = Normalize(fullPath);
                foreach (Excel.Workbook book in app.Workbooks)
                {
                    try
                    {
                        if (string.Equals(Normalize(book.FullName), norm, StringComparison.OrdinalIgnoreCase))
                        {
                            book.Close(false);
                            return true;
                        }
                    }
                    catch
                    {
                    }
                }
            }
            catch
            {
            }

            return false;
        }

        public static bool ClosePptByPath(string fullPath)
        {
            try
            {
                var app = (PowerPoint.Application)Marshal.GetActiveObject("PowerPoint.Application");
                string norm = Normalize(fullPath);
                foreach (PowerPoint.Presentation pres in app.Presentations)
                {
                    try
                    {
                        if (string.Equals(Normalize(pres.FullName), norm, StringComparison.OrdinalIgnoreCase))
                        {
                            pres.Close();
                            return true;
                        }
                    }
                    catch
                    {
                    }
                }
            }
            catch
            {
            }

            return false;
        }

        private static bool InsertPageBreakWord(string fullPath)
        {
            try
            {
                var app = (Word.Application)Marshal.GetActiveObject("Word.Application");
                Word.Document doc = FindWordDoc(app, fullPath);
                if (doc == null)
                {
                    return false;
                }

                object brk = Word.WdBreakType.wdPageBreak;
                doc.Range(doc.Content.End - 1, doc.Content.End - 1).InsertBreak(ref brk);
                return true;
            }
            catch
            {
                return false;
            }
        }

        private static bool CloseWordByPath(string fullPath)
        {
            try
            {
                var app = (Word.Application)Marshal.GetActiveObject("Word.Application");
                Word.Document doc = FindWordDoc(app, fullPath);
                if (doc == null)
                {
                    return false;
                }

                object save = false;
                doc.Close(ref save);
                return true;
            }
            catch
            {
                return false;
            }
        }

        private static Word.Document FindWordDoc(Word.Application app, string fullPath)
        {
            string norm = Normalize(fullPath);
            foreach (Word.Document d in app.Documents)
            {
                try
                {
                    if (string.Equals(Normalize(d.FullName), norm, StringComparison.OrdinalIgnoreCase))
                    {
                        return d;
                    }
                }
                catch
                {
                }
            }

            return null;
        }

        private static bool InsertPageBreakWps(string fullPath)
        {
            return WithWpsDoc(fullPath, doc =>
            {
                try
                {
                    object content = Invoke(doc, "Content");
                    int end = Convert.ToInt32(GetProp(content, "End"));
                    int pos = end > 1 ? end - 1 : end;
                    object range = Invoke(doc, "Range", pos, pos);
                    Invoke(range, "InsertBreak", 7);
                    return true;
                }
                catch
                {
                }

                try
                {
                    object content = GetProp(doc, "Content");
                    Invoke(content, "InsertAfter", "\f");
                    return true;
                }
                catch
                {
                }

                return false;
            });
        }

        private static bool CloseWpsByPath(string fullPath)
        {
            return WithWpsDoc(fullPath, doc =>
            {
                Invoke(doc, "Close", false);
                return true;
            });
        }

        private static bool WithWpsDoc(string fullPath, Func<object, bool> action)
        {
            foreach (object app in EnumerateWpsApps())
            {
                object doc = FindLateDoc(app, fullPath);
                if (doc == null)
                {
                    continue;
                }

                try
                {
                    if (action(doc))
                    {
                        return true;
                    }
                }
                catch
                {
                }
            }

            return false;
        }

        private static IEnumerable<object> EnumerateWpsApps()
        {
            var apps = new List<object>();
            string[] ids = { "Kwps.Application", "kwps.Application", "wps.Application" };
            foreach (string id in ids)
            {
                try
                {
                    object app = Marshal.GetActiveObject(id);
                    if (app != null)
                    {
                        apps.Add(app);
                    }
                }
                catch
                {
                }
            }

            CollectWpsAppsFromWindows(apps);
            return apps;
        }

        private static void CollectWpsAppsFromWindows(List<object> apps)
        {
            var pids = new HashSet<uint>();
            try
            {
                foreach (Process p in Process.GetProcessesByName("wps"))
                {
                    try
                    {
                        pids.Add((uint)p.Id);
                    }
                    finally
                    {
                        p.Dispose();
                    }
                }
            }
            catch
            {
                return;
            }

            if (pids.Count == 0)
            {
                return;
            }

            EnumWindows(
                (hWnd, _) =>
                {
                    if (!IsWindowVisible(hWnd))
                    {
                        return true;
                    }

                    GetWindowThreadProcessId(hWnd, out uint pid);
                    if (!pids.Contains(pid))
                    {
                        return true;
                    }

                    TryAddWpsAppFromWindow(hWnd, apps);
                    return true;
                },
                IntPtr.Zero);
        }

        private static void TryAddWpsAppFromWindow(IntPtr hwnd, List<object> apps)
        {
            Guid iid = IidIDispatch;
            object obj;
            if (AccessibleObjectFromWindow(hwnd, ObjIdNativeOm, ref iid, out obj) != 0 || obj == null)
            {
                return;
            }

            try
            {
                if (GetProp(obj, "Documents") != null)
                {
                    apps.Add(obj);
                    return;
                }
            }
            catch
            {
            }

            try
            {
                object app = GetProp(obj, "Application");
                if (app != null)
                {
                    apps.Add(app);
                }
            }
            catch
            {
            }
        }

        private static object FindLateDoc(object app, string fullPath)
        {
            try
            {
                object docs = GetProp(app, "Documents");
                int n = Convert.ToInt32(GetProp(docs, "Count"));
                string wantFile = Path.GetFileName(Normalize(fullPath));
                for (int i = 1; i <= n; i++)
                {
                    object d = Invoke(docs, "Item", i);
                    string name = Convert.ToString(GetProp(d, "FullName") ?? "");
                    if (PathMatches(name, fullPath))
                    {
                        return d;
                    }

                    if (!string.IsNullOrEmpty(wantFile)
                        && string.Equals(
                            Path.GetFileName(Normalize(name)),
                            wantFile,
                            StringComparison.OrdinalIgnoreCase))
                    {
                        return d;
                    }
                }
            }
            catch
            {
            }

            return null;
        }

        private static bool PathMatches(string a, string b)
        {
            string na = Normalize(a);
            string nb = Normalize(b);
            return !string.IsNullOrEmpty(na)
                && !string.IsNullOrEmpty(nb)
                && string.Equals(na, nb, StringComparison.OrdinalIgnoreCase);
        }

        private const uint ObjIdNativeOm = 0xFFFFFFF0;
        private static readonly Guid IidIDispatch = new Guid("00020400-0000-0000-C000-000000000046");

        private delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

        [DllImport("user32.dll")]
        private static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);

        [DllImport("user32.dll")]
        private static extern bool IsWindowVisible(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

        [DllImport("oleacc.dll")]
        private static extern int AccessibleObjectFromWindow(
            IntPtr hwnd,
            uint dwObjectID,
            ref Guid riid,
            [MarshalAs(UnmanagedType.IDispatch)] out object ppvObject);

        private static object GetProp(object target, string name)
        {
            return target.GetType().InvokeMember(
                name,
                System.Reflection.BindingFlags.GetProperty,
                null,
                target,
                null);
        }

        private static object Invoke(object target, string name, params object[] args)
        {
            return target.GetType().InvokeMember(
                name,
                System.Reflection.BindingFlags.InvokeMethod,
                null,
                target,
                args);
        }

        private static string Normalize(string path)
        {
            if (string.IsNullOrEmpty(path))
            {
                return "";
            }

            try
            {
                return Path.GetFullPath(path).TrimEnd('\\').ToLowerInvariant();
            }
            catch
            {
                return path.Trim().ToLowerInvariant();
            }
        }
    }
}
