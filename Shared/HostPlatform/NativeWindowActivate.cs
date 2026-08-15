using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;

namespace WordAddIn1.HostPlatform
{
    internal static class NativeWindowActivate
    {
        private const int SwRestore = 9;
        private const int SwMaximize = 3;
        private const int SwShow = 5;
        private const int MinWidthPx = 720;
        private const int MinHeightPx = 480;
        private const int SwpShowWindow = 0x0040;
        private const int SwpNoZOrder = 0x0004;

        [DllImport("user32.dll")]
        private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        [DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern bool IsIconic(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern bool IsZoomed(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern bool IsWindowVisible(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

        [DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

        [DllImport("user32.dll")]
        private static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);

        [DllImport("user32.dll")]
        private static extern bool SetWindowPos(
            IntPtr hWnd, IntPtr hWndInsertAfter, int x, int y, int cx, int cy, uint uFlags);

        private delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

        [StructLayout(LayoutKind.Sequential)]
        private struct RECT
        {
            public int Left;
            public int Top;
            public int Right;
            public int Bottom;
        }

        public static void BringToFront(int hwnd)
        {
            EnsureUsable(new[] { hwnd }, null, null);
        }

        public static void EnsureUsable(IList<int> hwndHints, string titleHint, IList<string> processNames)
        {
            IntPtr target = ResolveMainWindow(hwndHints, titleHint, processNames);
            if (target == IntPtr.Zero)
            {
                return;
            }

            try
            {
                if (IsIconic(target))
                {
                    ShowWindow(target, SwRestore);
                }

                ShowWindow(target, SwMaximize);
                if (!IsZoomed(target) && IsTooSmall(target))
                {
                    FitToWorkArea(target);
                }

                ShowWindow(target, SwShow);
                SetForegroundWindow(target);
            }
            catch (Exception)
            {
            }
        }

        private static IntPtr ResolveMainWindow(IList<int> hwndHints, string titleHint, IList<string> processNames)
        {
            var pids = new HashSet<uint>();
            if (hwndHints != null)
            {
                foreach (int raw in hwndHints)
                {
                    if (raw == 0)
                    {
                        continue;
                    }

                    GetWindowThreadProcessId(new IntPtr(raw), out uint pid);
                    if (pid != 0)
                    {
                        pids.Add(pid);
                    }
                }
            }

            if (processNames != null)
            {
                foreach (string name in processNames)
                {
                    if (string.IsNullOrWhiteSpace(name))
                    {
                        continue;
                    }

                    try
                    {
                        foreach (Process process in Process.GetProcessesByName(name))
                        {
                            try
                            {
                                pids.Add((uint)process.Id);
                            }
                            finally
                            {
                                process.Dispose();
                            }
                        }
                    }
                    catch (Exception)
                    {
                    }
                }
            }

            IntPtr titleMatch = IntPtr.Zero;
            int titleArea = 0;
            IntPtr largest = IntPtr.Zero;
            int largestArea = 0;
            string needle = string.IsNullOrWhiteSpace(titleHint) ? null : titleHint.Trim();

            EnumWindows((hWnd, _) =>
            {
                if (!IsWindowVisible(hWnd) || string.IsNullOrEmpty(GetTitle(hWnd)))
                {
                    return true;
                }

                GetWindowThreadProcessId(hWnd, out uint pid);
                bool sameProcess = pid != 0 && pids.Contains(pid);
                bool nameHit = needle != null
                    && GetTitle(hWnd).IndexOf(needle, StringComparison.OrdinalIgnoreCase) >= 0;
                if (!sameProcess && !nameHit)
                {
                    return true;
                }

                int area = GetArea(hWnd);
                if (nameHit && area >= titleArea)
                {
                    titleArea = area;
                    titleMatch = hWnd;
                }

                if (sameProcess && area > largestArea)
                {
                    largestArea = area;
                    largest = hWnd;
                }

                return true;
            }, IntPtr.Zero);

            if (titleMatch != IntPtr.Zero)
            {
                return titleMatch;
            }

            if (largest != IntPtr.Zero)
            {
                return largest;
            }

            if (hwndHints != null)
            {
                foreach (int raw in hwndHints)
                {
                    if (raw != 0)
                    {
                        return new IntPtr(raw);
                    }
                }
            }

            return IntPtr.Zero;
        }

        private static void FitToWorkArea(IntPtr handle)
        {
            try
            {
                var area = Screen.FromHandle(handle).WorkingArea;
                int width = Math.Max(MinWidthPx, (int)Math.Round(area.Width * 0.85));
                int height = Math.Max(MinHeightPx, (int)Math.Round(area.Height * 0.85));
                int x = area.Left + (area.Width - width) / 2;
                int y = area.Top + (area.Height - height) / 2;
                SetWindowPos(handle, IntPtr.Zero, x, y, width, height, SwpShowWindow | SwpNoZOrder);
            }
            catch (Exception)
            {
            }
        }

        private static bool IsTooSmall(IntPtr handle)
        {
            if (!GetWindowRect(handle, out RECT rect))
            {
                return false;
            }

            return (rect.Right - rect.Left) < MinWidthPx || (rect.Bottom - rect.Top) < MinHeightPx;
        }

        private static int GetArea(IntPtr handle)
        {
            if (!GetWindowRect(handle, out RECT rect))
            {
                return 0;
            }

            return Math.Max(0, rect.Right - rect.Left) * Math.Max(0, rect.Bottom - rect.Top);
        }

        private static string GetTitle(IntPtr handle)
        {
            var sb = new StringBuilder(512);
            int n = GetWindowText(handle, sb, sb.Capacity);
            return n <= 0 ? "" : sb.ToString();
        }
    }
}
