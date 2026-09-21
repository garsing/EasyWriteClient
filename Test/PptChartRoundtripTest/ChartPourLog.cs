using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace PptChartRoundtripTest
{
    /// <summary>把 [PptChartPour] 落到测试输出目录，不用去 Desktop logs 翻。</summary>
    internal static class ChartPourLog
    {
        public static string FilePath { get; private set; }

        public static void Init(string outDir)
        {
            if (string.IsNullOrEmpty(outDir))
            {
                return;
            }

            Directory.CreateDirectory(outDir);
            FilePath = Path.Combine(outDir, "chart-pour.log");
            var sb = new StringBuilder();
            sb.AppendLine("===== PptChartRoundtrip pour =====");
            sb.AppendLine("开始: " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff"));
            sb.AppendLine("文件: " + FilePath);
            sb.AppendLine("==================================");
            File.WriteAllText(FilePath, sb.ToString(), Encoding.UTF8);
            Console.WriteLine("pour 日志: " + FilePath);
        }

        public static void WriteCase(
            string tag,
            string result,
            string error,
            List<string> warnings,
            bool toConsole)
        {
            var sb = new StringBuilder();
            sb.AppendLine();
            sb.Append("==== ").Append(tag ?? "?").Append(' ').Append(result ?? "?").AppendLine(" ====");
            if (!string.IsNullOrEmpty(error))
            {
                sb.AppendLine(error);
            }

            int n = warnings == null ? 0 : warnings.Count;
            sb.Append("pour ").Append(n).AppendLine(" 条");
            if (warnings != null)
            {
                for (int i = 0; i < warnings.Count; i++)
                {
                    sb.AppendLine(warnings[i]);
                }
            }

            if (!string.IsNullOrEmpty(FilePath))
            {
                File.AppendAllText(FilePath, sb.ToString(), Encoding.UTF8);
            }

            if (!toConsole)
            {
                return;
            }

            Console.WriteLine("         ---- pour " + n + " 条"
                + (string.IsNullOrEmpty(FilePath) ? "" : " → " + FilePath)
                + " ----");
            if (warnings == null)
            {
                return;
            }

            for (int i = 0; i < warnings.Count; i++)
            {
                Console.WriteLine("         " + warnings[i]);
            }
        }

        public static string SeeFile()
        {
            return string.IsNullOrEmpty(FilePath) ? "" : " | 见 " + FilePath;
        }
    }
}
