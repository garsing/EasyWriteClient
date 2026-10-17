using System;
using System.Globalization;
using System.IO;
using System.Threading;

namespace WordAddIn1.PresentationHost
{
    /// <summary>
    /// 套用演示文稿时逐步落盘。进程卡住时，文件最后一行就是还没走完的那一步。
    /// 路径：%TEMP%\ew-ppt-apply-hang.log
    /// </summary>
    internal static class PptHtmlApplyHangLog
    {
        private static readonly object Sync = new object();
        private static int _seq;

        public static string FilePath => Path.Combine(Path.GetTempPath(), "ew-ppt-apply-hang.log");

        public static int NextApply()
        {
            return Interlocked.Increment(ref _seq);
        }

        public static void Mark(string message)
        {
            try
            {
                string line = DateTime.Now.ToString("HH:mm:ss.fff", CultureInfo.InvariantCulture)
                    + " t=" + Thread.CurrentThread.ManagedThreadId.ToString(CultureInfo.InvariantCulture)
                    + " " + (message ?? "")
                    + Environment.NewLine;
                lock (Sync)
                {
                    File.AppendAllText(FilePath, line);
                }
            }
            catch (Exception)
            {
            }
        }
    }
}
