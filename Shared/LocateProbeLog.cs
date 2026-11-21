using System;
using System.Diagnostics;
using System.IO;
using System.Text;

namespace WordAddIn1
{
    /// <summary>ul-apply-toc 对照写文档：目录 Interior 表落盘。查完删调用。</summary>
    internal static class LocateProbeLog
    {
        public static readonly string FilePath = Path.Combine(Path.GetTempPath(), "ew-ul-locate.txt");

        public static void Line(string message)
        {
            string line = DateTime.Now.ToString("HH:mm:ss.fff")
                + " [" + Process.GetCurrentProcess().ProcessName + "]"
                + " " + (message ?? "")
                + Environment.NewLine;
            try
            {
                File.AppendAllText(FilePath, line, Encoding.UTF8);
            }
            catch
            {
            }

            Debug.WriteLine("[ul-locate] " + message);
        }
    }
}
