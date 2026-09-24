using System;
using System.Collections.Generic;
using System.IO;

namespace PptHtmlContentRoundtripTest
{
    internal static class Program
    {
        [STAThread]
        private static int Main(string[] args)
        {
            Console.OutputEncoding = System.Text.Encoding.UTF8;
            var run = new TestRun();
            bool parseOnly = HasFlag(args, "--parse-only");
            bool formalOnly = HasFlag(args, "--formal-only");
            bool failFast = HasFlag(args, "--fail-fast");
            bool useWpp = IsWppHost(args);
            int? batch = TryParseBatch(args);
            IList<string> nameFilters = ParseCaseFilters(args);

            Console.WriteLine("约定 HTML · 文本框 / 表格 / 图片 往返测试（P0）");
            Console.WriteLine(useWpp
                ? "宿主 WPP · 解析契约 + COM 正式用例（WppApplier → WppReader）"
                : "宿主 PPT · 解析契约 + COM 正式用例（Applier → Reader）");
            Console.WriteLine();

            if (!formalOnly)
            {
                Console.WriteLine("--- 解析契约（不启 COM）---");
                ContentParseTests.Run(run);
            }

            if (!parseOnly)
            {
                string outDir = Path.Combine(
                    AppDomain.CurrentDomain.BaseDirectory,
                    useWpp ? "content-batches-wpp" : "content-batches");
                if (useWpp)
                {
                    ContentBatchRunnerWpp.Run(run, batch, outDir, nameFilters);
                }
                else
                {
                    ContentBatchRunner.Run(run, batch, outDir, nameFilters, failFast);
                }

                Console.WriteLine();
                Console.WriteLine("正式用例 通过 " + run.CasesPassed
                    + "  失败 " + run.CasesFailed
                    + "  跳过 " + run.CasesSkipped
                    + "  （目标 " + ContentCatalog.Total + "）");
            }

            Console.WriteLine();
            Console.WriteLine("断言 通过 " + run.Passed + "  失败 " + run.Failed + "  跳过 " + run.Skipped);
            return run.Failed == 0 && run.CasesFailed == 0 ? 0 : 1;
        }

        private static bool HasFlag(string[] args, string flag)
        {
            if (args == null)
            {
                return false;
            }

            for (int i = 0; i < args.Length; i++)
            {
                if (string.Equals(args[i], flag, StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
            }

            return false;
        }

        /// <summary>--host=wpp 走 WPS 演示晚绑定；默认 / --host=ppt 走 PowerPoint Interop。</summary>
        private static bool IsWppHost(string[] args)
        {
            if (args == null)
            {
                return false;
            }

            for (int i = 0; i < args.Length; i++)
            {
                string a = args[i] ?? "";
                if (a.StartsWith("--host=", StringComparison.OrdinalIgnoreCase))
                {
                    string v = a.Substring(7).Trim();
                    return string.Equals(v, "wpp", StringComparison.OrdinalIgnoreCase)
                        || string.Equals(v, "wps", StringComparison.OrdinalIgnoreCase);
                }
            }

            return false;
        }

        private static int? TryParseBatch(string[] args)
        {
            if (args == null)
            {
                return null;
            }

            for (int i = 0; i < args.Length; i++)
            {
                string a = args[i] ?? "";
                if (a.StartsWith("--batch=", StringComparison.OrdinalIgnoreCase)
                    && int.TryParse(a.Substring(8), out int n)
                    && n >= 1
                    && n <= ContentCatalog.BatchCount)
                {
                    return n;
                }
            }

            return null;
        }

        private static IList<string> ParseCaseFilters(string[] args)
        {
            var list = new List<string>();
            if (args == null)
            {
                return list;
            }

            for (int i = 0; i < args.Length; i++)
            {
                string a = args[i] ?? "";
                if (!a.StartsWith("--case=", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                string raw = a.Substring(7);
                string[] parts = raw.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
                for (int p = 0; p < parts.Length; p++)
                {
                    string one = parts[p].Trim();
                    if (one.Length > 0)
                    {
                        list.Add(one);
                    }
                }
            }

            return list;
        }
    }
}
