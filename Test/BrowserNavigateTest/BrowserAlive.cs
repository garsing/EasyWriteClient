using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;

namespace BrowserNavigateTest
{
    /// <summary>
    /// 不得 kill 用户 Chrome/Edge。helper / renderer 进进出出不算关浏览器。
    /// </summary>
    internal static class BrowserAlive
    {
        internal static bool NotQuit(string name, int before, out string detail)
        {
            int now = Count(name);
            bool win = HasVisibleWindow(name);
            detail = name + " " + before + "→" + now + (win ? " 有窗" : " 无窗");
            if (before <= 0)
            {
                return true;
            }

            if (now <= 0)
            {
                return false;
            }

            return win;
        }

        internal static int Count(string name)
        {
            try
            {
                return Process.GetProcessesByName(name).Length;
            }
            catch
            {
                return 0;
            }
        }

        private static bool HasVisibleWindow(string processName)
        {
            var pids = new HashSet<uint>();
            try
            {
                foreach (Process p in Process.GetProcessesByName(processName))
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
                return false;
            }

            if (pids.Count == 0)
            {
                return false;
            }

            bool found = false;
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

                    var sb = new StringBuilder(64);
                    if (GetClassName(hWnd, sb, sb.Capacity) <= 0)
                    {
                        return true;
                    }

                    if (sb.ToString() == "Chrome_WidgetWin_1")
                    {
                        found = true;
                        return false;
                    }

                    return true;
                },
                IntPtr.Zero);
            return found;
        }

        private delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

        [DllImport("user32.dll")]
        private static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);

        [DllImport("user32.dll")]
        private static extern bool IsWindowVisible(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);
    }
}
