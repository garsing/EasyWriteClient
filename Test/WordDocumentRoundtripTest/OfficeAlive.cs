using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;

namespace WordDocumentRoundtripTest
{
    /// <summary>不得 Quit 用户 Word / WPS。helper 进出不算关应用。</summary>
    internal static class OfficeAlive
    {
        internal static bool NotQuit(string name, int before, bool windowBefore, out string detail)
        {
            int now = Count(name);
            bool win = HasVisibleWindow(name);
            detail = name + " " + before + "→" + now
                + (win ? " 有窗" : " 无窗")
                + (windowBefore ? " 开跑前有窗" : " 开跑前无窗");
            if (before <= 0)
            {
                return true;
            }

            if (now <= 0)
            {
                return false;
            }

            if (!windowBefore)
            {
                return true;
            }

            return win;
        }

        internal static bool HasWindow(string name)
        {
            return HasVisibleWindow(name);
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

                    var sb = new StringBuilder(256);
                    if (GetWindowText(hWnd, sb, sb.Capacity) <= 0)
                    {
                        return true;
                    }

                    found = true;
                    return false;
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
        private static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);
    }
}
