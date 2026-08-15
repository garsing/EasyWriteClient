using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;

namespace WordAddIn1
{
    /// <summary>
    /// Excel 多实例：GetActiveObject 只返回 ROT 里第一个，常是空壳（Workbooks=0）。
    /// 通过各 EXCEL 主窗的 EXCEL7 子窗取 Native OM，优先选有工作簿的实例。
    /// </summary>
    internal static class ExcelApplicationInstances
    {
        private const uint ObjIdNativeOm = 0xFFFFFFF0;

        private static readonly Guid IidIDispatch = new Guid("00020400-0000-0000-C000-000000000046");

        private delegate bool EnumChildCallback(int hwnd, ref int lParam);

        [DllImport("oleacc.dll")]
        private static extern int AccessibleObjectFromWindow(
            int hwnd,
            uint dwObjectId,
            ref Guid riid,
            [MarshalAs(UnmanagedType.IDispatch)] out object ppvObject);

        [DllImport("User32.dll")]
        private static extern bool EnumChildWindows(int hWndParent, EnumChildCallback lpEnumFunc, ref int lParam);

        [DllImport("User32.dll")]
        private static extern int GetClassName(int hWnd, StringBuilder lpClassName, int nMaxCount);

        /// <summary>
        /// 解析一个可用的 Excel.Application：优先有 Workbooks 的实例；否则退回 GetActiveObject。
        /// </summary>
        public static bool TryFindBest(
            out Excel.Application application,
            out string source,
            out int workbookCount)
        {
            application = null;
            source = null;
            workbookCount = -1;

            Excel.Application bestWithBooks = null;
            int bestCount = 0;
            string bestSource = null;
            Excel.Application anyAlive = null;
            string anySource = null;

            foreach (Excel.Application candidate in EnumerateRunning())
            {
                if (candidate == null)
                {
                    continue;
                }

                int count = SafeWorkbookCount(candidate);
                string name = SafeName(candidate);
                EasyWriteDiagnostics.Log(DebugCategory.OpenFiles,
                    "[ExcelApplicationInstances] candidate name=" + name
                    + " workbooks=" + count);

                if (count > bestCount)
                {
                    bestWithBooks = candidate;
                    bestCount = count;
                    bestSource = "hwnd-om";
                }
                else if (anyAlive == null && count >= 0)
                {
                    anyAlive = candidate;
                    anySource = "hwnd-om";
                }
            }

            try
            {
                var active = (Excel.Application)Marshal.GetActiveObject("Excel.Application");
                if (active != null)
                {
                    int count = SafeWorkbookCount(active);
                    EasyWriteDiagnostics.Log(DebugCategory.OpenFiles,
                        "[ExcelApplicationInstances] GetActiveObject name="
                        + SafeName(active) + " workbooks=" + count);
                    if (count > bestCount)
                    {
                        bestWithBooks = active;
                        bestCount = count;
                        bestSource = "GetActiveObject";
                    }
                    else if (anyAlive == null && count >= 0)
                    {
                        anyAlive = active;
                        anySource = "GetActiveObject";
                    }
                }
            }
            catch (COMException)
            {
            }
            catch (Exception ex)
            {
                EasyWriteDiagnostics.Log(DebugCategory.OpenFiles,
                    "[ExcelApplicationInstances] GetActiveObject: " + ex.Message);
            }

            if (bestWithBooks != null)
            {
                application = bestWithBooks;
                source = bestSource;
                workbookCount = bestCount;
                return true;
            }

            if (anyAlive != null)
            {
                application = anyAlive;
                source = anySource;
                workbookCount = SafeWorkbookCount(anyAlive);
                return true;
            }

            return false;
        }

        private static IEnumerable<Excel.Application> EnumerateRunning()
        {
            Process[] procs;
            try
            {
                procs = Process.GetProcessesByName("EXCEL");
            }
            catch (Exception ex)
            {
                EasyWriteDiagnostics.Log(DebugCategory.OpenFiles,
                    "[ExcelApplicationInstances] GetProcessesByName: " + ex.Message);
                yield break;
            }

            var seen = new HashSet<int>();
            foreach (Process p in procs)
            {
                IntPtr main;
                try
                {
                    main = p.MainWindowHandle;
                }
                catch (Exception)
                {
                    continue;
                }

                if (main == IntPtr.Zero)
                {
                    continue;
                }

                int excel7 = FindExcel7Child((int)main);
                if (excel7 == 0)
                {
                    EasyWriteDiagnostics.Log(DebugCategory.OpenFiles,
                        "[ExcelApplicationInstances] no EXCEL7 pid=" + p.Id);
                    continue;
                }

                Excel.Application app = TryApplicationFromHwnd(excel7);
                if (app == null)
                {
                    continue;
                }

                int key = RuntimeHelpersHash(app);
                if (!seen.Add(key))
                {
                    continue;
                }

                yield return app;
            }
        }

        private static int FindExcel7Child(int mainHwnd)
        {
            int child = 0;
            EnumChildWindows(
                mainHwnd,
                (int hwndChild, ref int lParam) =>
                {
                    var buf = new StringBuilder(128);
                    GetClassName(hwndChild, buf, 128);
                    if (string.Equals(buf.ToString(), "EXCEL7", StringComparison.Ordinal))
                    {
                        lParam = hwndChild;
                        return false;
                    }

                    return true;
                },
                ref child);
            return child;
        }

        private static Excel.Application TryApplicationFromHwnd(int excel7Hwnd)
        {
            try
            {
                Guid iid = IidIDispatch;
                int hr = AccessibleObjectFromWindow(excel7Hwnd, ObjIdNativeOm, ref iid, out object obj);
                if (hr < 0 || obj == null)
                {
                    EasyWriteDiagnostics.Log(DebugCategory.OpenFiles,
                        "[ExcelApplicationInstances] AccessibleObjectFromWindow hr=" + hr);
                    return null;
                }

                // Excel.Window → Application
                var window = obj as Excel.Window;
                if (window != null)
                {
                    return window.Application;
                }

                // 部分环境返回的是 Worksheet / Workbook
                try
                {
                    dynamic dyn = obj;
                    object appObj = dyn.Application;
                    return appObj as Excel.Application ?? (Excel.Application)appObj;
                }
                catch (Exception ex)
                {
                    EasyWriteDiagnostics.Log(DebugCategory.OpenFiles,
                        "[ExcelApplicationInstances] dynamic Application: " + ex.Message
                        + " type=" + obj.GetType().FullName);
                    return null;
                }
            }
            catch (Exception ex)
            {
                EasyWriteDiagnostics.Log(DebugCategory.OpenFiles,
                    "[ExcelApplicationInstances] from hwnd: " + ex.Message);
                return null;
            }
        }

        private static int RuntimeHelpersHash(object com)
        {
            try
            {
                return System.Runtime.CompilerServices.RuntimeHelpers.GetHashCode(com);
            }
            catch (Exception)
            {
                return com?.GetHashCode() ?? 0;
            }
        }

        private static string SafeName(Excel.Application app)
        {
            try
            {
                return app?.Name ?? "";
            }
            catch (Exception)
            {
                return "?";
            }
        }

        private static int SafeWorkbookCount(Excel.Application app)
        {
            try
            {
                return app?.Workbooks?.Count ?? -1;
            }
            catch (Exception)
            {
                return -1;
            }
        }
    }
}
