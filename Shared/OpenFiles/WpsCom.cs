using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace WordAddIn1.OpenFiles
{
    /// <summary>
    /// WPS 文字 COM 晚绑定辅助：GetActiveObject + 多进程 HWND，不 new / 不 Quit。
    /// </summary>
    internal static class WpsCom
    {
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

        public static readonly string[] ProgIds =
        {
            "Kwps.Application",
            "kwps.Application",
            "wps.Application"
        };

        public static object TryGetActiveApplication(out string progId)
        {
            progId = null;
            object bestWithDocs = null;
            int bestCount = 0;
            object anyAlive = null;
            string bestId = null;
            string anyId = null;

            foreach (WpsAppCandidate candidate in EnumerateRunningApplications())
            {
                if (candidate.Application == null || !CanReadDocuments(candidate.Application))
                {
                    continue;
                }

                int count = SafeDocumentCount(candidate.Application);
                if (count > bestCount)
                {
                    bestWithDocs = candidate.Application;
                    bestCount = count;
                    bestId = candidate.ProgId;
                }
                else if (anyAlive == null && count >= 0)
                {
                    anyAlive = candidate.Application;
                    anyId = candidate.ProgId;
                }
            }

            if (bestWithDocs != null)
            {
                progId = bestId ?? ProgIds[0];
                return bestWithDocs;
            }

            if (anyAlive != null)
            {
                progId = anyId ?? ProgIds[0];
                return anyAlive;
            }

            return null;
        }

        private struct WpsAppCandidate
        {
            public object Application;
            public string ProgId;
        }

        private static List<WpsAppCandidate> EnumerateRunningApplications()
        {
            var list = new List<WpsAppCandidate>();
            var seenHwnd = new HashSet<int>();
            foreach (string id in ProgIds)
            {
                object app = null;
                try
                {
                    app = Marshal.GetActiveObject(id);
                }
                catch (COMException)
                {
                    app = null;
                }
                catch (Exception)
                {
                    app = null;
                }

                if (app == null)
                {
                    continue;
                }

                int key = SafeAppHwnd(app);
                if (key == 0)
                {
                    key = RuntimeHelpers.GetHashCode(app);
                }

                if (seenHwnd.Add(key))
                {
                    list.Add(new WpsAppCandidate { Application = app, ProgId = id });
                }
            }

            foreach (object app in CollectApplicationsFromWindows())
            {
                int key = SafeAppHwnd(app);
                if (key == 0)
                {
                    key = RuntimeHelpers.GetHashCode(app);
                }

                if (!seenHwnd.Add(key))
                {
                    continue;
                }

                list.Add(new WpsAppCandidate { Application = app, ProgId = ProgIds[0] });
            }

            return list;
        }

        private static List<object> CollectApplicationsFromWindows()
        {
            var apps = new List<object>();
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
            catch (Exception)
            {
                return apps;
            }

            if (pids.Count == 0)
            {
                return apps;
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

                    object app = TryApplicationFromWindow(hWnd);
                    if (app != null)
                    {
                        apps.Add(app);
                    }

                    return true;
                },
                IntPtr.Zero);

            return apps;
        }

        private static object TryApplicationFromWindow(IntPtr hwnd)
        {
            Guid iid = IidIDispatch;
            object obj;
            if (AccessibleObjectFromWindow(hwnd, ObjIdNativeOm, ref iid, out obj) != 0 || obj == null)
            {
                return null;
            }

            try
            {
                if (GetProperty(obj, "Documents") != null)
                {
                    return obj;
                }
            }
            catch (Exception)
            {
            }

            try
            {
                return GetProperty(obj, "Application");
            }
            catch (Exception)
            {
                return null;
            }
        }

        private static int SafeDocumentCount(object app)
        {
            try
            {
                object docs = GetProperty(app, "Documents");
                object countObj = GetProperty(docs, "Count");
                return countObj == null ? -1 : Convert.ToInt32(countObj);
            }
            catch (Exception)
            {
                return -1;
            }
        }

        private static int SafeAppHwnd(object app)
        {
            if (app == null)
            {
                return 0;
            }

            try
            {
                object raw = GetProperty(app, "Hwnd") ?? GetProperty(app, "HWND");
                return raw == null ? 0 : Convert.ToInt32(raw);
            }
            catch (Exception)
            {
                return 0;
            }
        }

        public static bool CanReadDocuments(object app)
        {
            try
            {
                object docs = GetProperty(app, "Documents");
                if (docs == null)
                {
                    return false;
                }

                object countObj = GetProperty(docs, "Count");
                if (countObj == null)
                {
                    return false;
                }

                Convert.ToInt32(countObj);
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        public static bool IsAlive(object app)
        {
            try
            {
                object name = GetProperty(app, "Name");
                return name != null;
            }
            catch (Exception)
            {
                return false;
            }
        }

        public static IEnumerable<object> EnumerateDocuments(object app)
        {
            if (app == null)
            {
                yield break;
            }

            object docs;
            int count;
            try
            {
                docs = GetProperty(app, "Documents");
                count = Convert.ToInt32(GetProperty(docs, "Count"));
            }
            catch (Exception)
            {
                yield break;
            }

            // Word/WPS 集合多为 1-based
            for (int i = 1; i <= count; i++)
            {
                object doc = null;
                try
                {
                    doc = GetIndexed(docs, i);
                }
                catch (Exception)
                {
                    doc = null;
                }

                if (doc != null)
                {
                    yield return doc;
                }
            }
        }

        public static string TryReadFullName(object doc)
        {
            if (doc == null)
            {
                return null;
            }

            try
            {
                object full = GetProperty(doc, "FullName");
                string fullName = full as string;
                if (string.IsNullOrWhiteSpace(fullName))
                {
                    return null;
                }

                if (fullName.StartsWith("Unsaved", StringComparison.OrdinalIgnoreCase))
                {
                    return null;
                }

                return fullName;
            }
            catch (Exception)
            {
                return null;
            }
        }

        public static string TryReadName(object doc)
        {
            if (doc == null)
            {
                return null;
            }

            try
            {
                return GetProperty(doc, "Name") as string;
            }
            catch (Exception)
            {
                return null;
            }
        }

        public static object GetProperty(object target, string name)
        {
            if (target == null || string.IsNullOrEmpty(name))
            {
                return null;
            }

            return target.GetType().InvokeMember(
                name,
                BindingFlags.GetProperty | BindingFlags.Instance | BindingFlags.Public,
                null,
                target,
                null);
        }

        public static object Invoke(object target, string name, params object[] args)
        {
            return target.GetType().InvokeMember(
                name,
                BindingFlags.InvokeMethod | BindingFlags.Instance | BindingFlags.Public,
                null,
                target,
                args);
        }

        public static void TrySetProperty(object target, string name, object value)
        {
            if (target == null || string.IsNullOrEmpty(name))
            {
                return;
            }

            try
            {
                target.GetType().InvokeMember(
                    name,
                    BindingFlags.SetProperty | BindingFlags.Instance | BindingFlags.Public,
                    null,
                    target,
                    new object[] { value });
            }
            catch (Exception)
            {
            }
        }

        public static object TryGetItem(object collection, int index)
        {
            if (collection == null)
            {
                return null;
            }

            try
            {
                return GetIndexed(collection, index);
            }
            catch (Exception)
            {
                return null;
            }
        }

        private static object GetIndexed(object collection, int index)
        {
            // 优先 Item 属性/方法
            try
            {
                return collection.GetType().InvokeMember(
                    "Item",
                    BindingFlags.GetProperty | BindingFlags.InvokeMethod | BindingFlags.Instance | BindingFlags.Public,
                    null,
                    collection,
                    new object[] { index });
            }
            catch (Exception)
            {
            }

            try
            {
                return collection.GetType().InvokeMember(
                    "get_Item",
                    BindingFlags.InvokeMethod | BindingFlags.Instance | BindingFlags.Public,
                    null,
                    collection,
                    new object[] { index });
            }
            catch (Exception)
            {
            }

            // 部分实现支持默认索引器
            return collection.GetType().InvokeMember(
                "Item",
                BindingFlags.GetProperty | BindingFlags.InvokeMethod | BindingFlags.Instance | BindingFlags.Public | BindingFlags.IgnoreCase,
                null,
                collection,
                new object[] { index });
        }
    }
}
