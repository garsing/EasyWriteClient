using System;
using System.Collections.Generic;
using System.Reflection;
using System.Runtime.InteropServices;

namespace WordAddIn1.OpenFiles
{
    internal static class EtCom
    {
        public static readonly string[] ProgIds =
        {
            "Ket.Application",
            "ket.Application",
            "et.Application"
        };

        public static object TryGetActiveApplication(out string progId)
        {
            progId = null;
            foreach (string id in ProgIds)
            {
                try
                {
                    object app = Marshal.GetActiveObject(id);
                    if (app != null && CanReadWorkbooks(app))
                    {
                        progId = id;
                        return app;
                    }
                }
                catch (COMException)
                {
                }
                catch (Exception)
                {
                }
            }

            return null;
        }

        public static bool TryCreateApplication(out object application, out string error)
        {
            application = null;
            error = null;
            foreach (string id in ProgIds)
            {
                try
                {
                    Type t = Type.GetTypeFromProgID(id);
                    if (t == null)
                    {
                        continue;
                    }

                    application = Activator.CreateInstance(t);
                    if (application == null)
                    {
                        continue;
                    }

                    EnsureVisible(application);
                    return true;
                }
                catch (Exception ex)
                {
                    error = ex.Message;
                    application = null;
                }
            }

            if (string.IsNullOrEmpty(error))
            {
                error = "无法创建 WPS 表格（Ket.Application / et.Application），是否未安装 WPS？";
            }

            return false;
        }

        public static bool IsAlive(object app)
        {
            try
            {
                return GetProperty(app, "Name") != null;
            }
            catch (Exception)
            {
                return false;
            }
        }

        public static bool CanReadWorkbooks(object app)
        {
            try
            {
                object books = GetProperty(app, "Workbooks");
                if (books == null)
                {
                    return false;
                }

                Convert.ToInt32(GetProperty(books, "Count"));
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        public static void EnsureVisible(object app, string titleHint = null)
        {
            if (app == null)
            {
                return;
            }

            TrySetProperty(app, "Visible", true);
            TrySetProperty(app, "UserControl", true);
            const int xlMaximized = -4137;
            TrySetProperty(app, "WindowState", xlMaximized);

            var hwnds = new List<int>();
            if (TryReadHwnd(app, out int appHwnd))
            {
                hwnds.Add(appHwnd);
            }

            try
            {
                object active = GetProperty(app, "ActiveWindow");
                TrySetProperty(active, "Visible", true);
                TrySetProperty(active, "WindowState", xlMaximized);
                if (TryReadHwnd(active, out int activeHwnd))
                {
                    hwnds.Add(activeHwnd);
                }
            }
            catch (Exception)
            {
            }

            HostPlatform.NativeWindowActivate.EnsureUsable(hwnds, titleHint, new[] { "et", "wps" });
        }

        public static IEnumerable<object> EnumerateWorkbooks(object app)
        {
            if (app == null)
            {
                yield break;
            }

            object books = null;
            int count;
            try
            {
                books = GetProperty(app, "Workbooks");
                count = Convert.ToInt32(GetProperty(books, "Count"));
            }
            catch (Exception)
            {
                ComRelease.Safe(books);
                yield break;
            }

            try
            {
                for (int i = 1; i <= count; i++)
                {
                    object book = null;
                    try
                    {
                        book = GetIndexed(books, i);
                    }
                    catch (Exception)
                    {
                    }

                    if (book != null)
                    {
                        yield return book;
                    }
                }
            }
            finally
            {
                ComRelease.Safe(books);
            }
        }

        public static IEnumerable<object> EnumerateSheets(object book)
        {
            if (book == null)
            {
                yield break;
            }

            object sheets;
            int count;
            try
            {
                sheets = GetProperty(book, "Sheets");
                count = Convert.ToInt32(GetProperty(sheets, "Count"));
            }
            catch (Exception)
            {
                yield break;
            }

            for (int i = 1; i <= count; i++)
            {
                object sheet = null;
                try
                {
                    sheet = GetIndexed(sheets, i);
                }
                catch (Exception)
                {
                }

                if (sheet != null)
                {
                    yield return sheet;
                }
            }
        }

        public static string TryReadFullName(object book)
        {
            try
            {
                return GetProperty(book, "FullName") as string;
            }
            catch (Exception)
            {
                return null;
            }
        }

        public static string TryReadName(object book)
        {
            try
            {
                return GetProperty(book, "Name") as string;
            }
            catch (Exception)
            {
                return null;
            }
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

        public static object GetIndexed2(object collection, int row, int col)
        {
            try
            {
                return collection.GetType().InvokeMember(
                    "Item",
                    BindingFlags.GetProperty | BindingFlags.InvokeMethod | BindingFlags.Instance | BindingFlags.Public,
                    null,
                    collection,
                    new object[] { row, col });
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
                    new object[] { row, col });
            }
            catch (Exception)
            {
                return null;
            }
        }

        public static object GetIndexed(object collection, int index)
        {
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
                return null;
            }
        }

        /// <summary>InvokeMethod | GetProperty | OptionalParamBinding，适配 WPS IDispatch。</summary>
        public static object InvokeFlex(object target, string name, params object[] args)
        {
            if (target == null || string.IsNullOrEmpty(name))
            {
                return null;
            }

            return target.GetType().InvokeMember(
                name,
                BindingFlags.InvokeMethod
                    | BindingFlags.GetProperty
                    | BindingFlags.OptionalParamBinding
                    | BindingFlags.Instance
                    | BindingFlags.Public,
                null,
                target,
                args ?? new object[0],
                null,
                null,
                null);
        }

        private static bool TryReadHwnd(object target, out int hwnd)
        {
            hwnd = 0;
            if (target == null)
            {
                return false;
            }

            object raw = null;
            try
            {
                raw = GetProperty(target, "Hwnd") ?? GetProperty(target, "HWND");
            }
            catch (Exception)
            {
            }

            if (raw == null)
            {
                return false;
            }

            try
            {
                long value = Convert.ToInt64(raw);
                if (value == 0 || value > int.MaxValue || value < int.MinValue)
                {
                    return false;
                }

                hwnd = (int)value;
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }
    }
}
