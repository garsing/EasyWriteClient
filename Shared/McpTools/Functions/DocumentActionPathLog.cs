using System.Collections.Generic;

namespace WordAddIn1
{
    /// <summary>
    /// process_actions 执行路径金标：集合 vs 单步、定位次数、映射、写后 PD。
    /// 进会话 log，并回包 <c>geo_plan_log</c>，供测试断言「走了设计路径」而不是只看落字。
    /// </summary>
    public static class DocumentActionPathLog
    {
        public const string LogTag = "[几何码路径]";

        private static readonly object Sync = new object();
        private static readonly List<string> Lines = new List<string>();

        public static void Clear()
        {
            lock (Sync)
            {
                Lines.Clear();
            }
        }

        public static List<string> Snapshot()
        {
            lock (Sync)
            {
                return new List<string>(Lines);
            }
        }

        public static void Record(string line)
        {
            if (string.IsNullOrEmpty(line))
            {
                return;
            }

            if (line.IndexOf(LogTag, System.StringComparison.Ordinal) < 0)
            {
                line = LogTag + " " + line;
            }

            lock (Sync)
            {
                Lines.Add(line);
            }

            EasyWriteLog.WriteLine(line);
            EasyWriteLog.Flush();
            System.Diagnostics.Debug.WriteLine(line);
        }
    }
}
