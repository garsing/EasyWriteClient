using System;
using System.Collections.Generic;
using System.Reflection;
using System.Runtime.InteropServices;

namespace WordAddIn1.OpenFiles
{
    /// <summary>WPS 演示晚绑定 COM（对齐 EtCom）。</summary>
    internal static class WppCom
    {
        public static readonly string[] ProgIds =
        {
            "Kwpp.Application",
            "kwpp.Application",
            "wpp.Application"
        };

        public static object TryGetActiveApplication(out string progId)
        {
            progId = null;
            foreach (string id in ProgIds)
            {
                try
                {
                    object app = Marshal.GetActiveObject(id);
                    if (app != null && CanReadPresentations(app))
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
                error = "无法创建 WPS 演示（Kwpp.Application / wpp.Application），是否未安装 WPS？";
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

        public static bool CanReadPresentations(object app)
        {
            try
            {
                object presentations = GetProperty(app, "Presentations");
                if (presentations == null)
                {
                    return false;
                }

                Convert.ToInt32(GetProperty(presentations, "Count"));
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
            var hwnds = new List<int>();
            if (TryReadHwnd(app, out int appHwnd))
            {
                hwnds.Add(appHwnd);
            }

            HostPlatform.NativeWindowActivate.EnsureUsable(hwnds, titleHint, new[] { "wpp", "wps" });
        }

        public static IEnumerable<object> EnumeratePresentations(object app)
        {
            if (app == null)
            {
                yield break;
            }

            object presentations = null;
            int count;
            try
            {
                presentations = GetProperty(app, "Presentations");
                count = Convert.ToInt32(GetProperty(presentations, "Count"));
            }
            catch (Exception)
            {
                ComRelease.Safe(presentations);
                yield break;
            }

            try
            {
                for (int i = 1; i <= count; i++)
                {
                    object presentation = null;
                    try
                    {
                        presentation = GetIndexed(presentations, i);
                    }
                    catch (Exception)
                    {
                    }

                    if (presentation != null)
                    {
                        yield return presentation;
                    }
                }
            }
            finally
            {
                ComRelease.Safe(presentations);
            }
        }

        public static string TryReadName(object presentation)
        {
            try
            {
                object name = GetProperty(presentation, "Name");
                return name == null ? null : Convert.ToString(name);
            }
            catch (Exception)
            {
                return null;
            }
        }

        public static string TryReadFullName(object presentation)
        {
            try
            {
                object full = GetProperty(presentation, "FullName");
                string text = full == null ? null : Convert.ToString(full);
                return string.IsNullOrWhiteSpace(text) ? null : text;
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

        private static bool TryReadHwnd(object target, out int hwnd)
        {
            hwnd = 0;
            if (target == null)
            {
                return false;
            }

            try
            {
                object value = GetProperty(target, "HWND");
                if (value == null)
                {
                    value = GetProperty(target, "Hwnd");
                }

                if (value == null)
                {
                    return false;
                }

                hwnd = Convert.ToInt32(value);
                return hwnd != 0;
            }
            catch (Exception)
            {
                return false;
            }
        }
    }
}
