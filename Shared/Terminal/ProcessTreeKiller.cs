using System;
using System.Diagnostics;
using System.Runtime.InteropServices;

namespace WordAddIn1.Terminal
{
    internal static class ProcessTreeKiller
    {
        public static void KillDescendants(int parentPid)
        {
            foreach (int child in ListChildren(parentPid))
            {
                KillTree(child);
            }
        }

        public static void KillTree(int rootPid)
        {
            KillDescendants(rootPid);
            TryKill(rootPid);
        }

        private static int[] ListChildren(int parentPid)
        {
            var ids = new System.Collections.Generic.List<int>();
            IntPtr snapshot = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0);
            if (snapshot == InvalidHandle)
            {
                return ids.ToArray();
            }

            try
            {
                var entry = new PROCESSENTRY32 { dwSize = (uint)Marshal.SizeOf(typeof(PROCESSENTRY32)) };
                if (!Process32First(snapshot, ref entry))
                {
                    return ids.ToArray();
                }

                do
                {
                    if (entry.th32ParentProcessID == (uint)parentPid
                        && entry.th32ProcessID != (uint)parentPid)
                    {
                        ids.Add((int)entry.th32ProcessID);
                    }
                }
                while (Process32Next(snapshot, ref entry));
            }
            finally
            {
                CloseHandle(snapshot);
            }

            return ids.ToArray();
        }

        private static void TryKill(int pid)
        {
            try
            {
                Process p = Process.GetProcessById(pid);
                if (!p.HasExited)
                {
                    p.Kill();
                }
            }
            catch
            {
            }
        }

        private const uint TH32CS_SNAPPROCESS = 2;
        private static readonly IntPtr InvalidHandle = new IntPtr(-1);

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern IntPtr CreateToolhelp32Snapshot(uint dwFlags, uint th32ProcessID);

        [DllImport("kernel32.dll")]
        private static extern bool Process32First(IntPtr hSnapshot, ref PROCESSENTRY32 lppe);

        [DllImport("kernel32.dll")]
        private static extern bool Process32Next(IntPtr hSnapshot, ref PROCESSENTRY32 lppe);

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern bool CloseHandle(IntPtr hObject);

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Auto)]
        private struct PROCESSENTRY32
        {
            public uint dwSize;
            public uint cntUsage;
            public uint th32ProcessID;
            public IntPtr th32DefaultHeapID;
            public uint th32ModuleID;
            public uint cntThreads;
            public uint th32ParentProcessID;
            public int pcPriClassBase;
            public uint dwFlags;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 260)]
            public string szExeFile;
        }
    }
}
