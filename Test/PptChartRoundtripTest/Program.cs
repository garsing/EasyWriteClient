using System;
using System.IO;

namespace PptChartRoundtripTest
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
            int? batch = TryParseBatch(args);

            Console.WriteLine("PPT 图表测试");
            Console.WriteLine("正式用例：一页一张完整图（新建或改已有），约 50 页 × 6 份 PPT");
            Console.WriteLine();

            if (!formalOnly)
            {
                Console.WriteLine("--- 解析契约（不启 PPT，按属性断言）---");
                ChartParseTests.Run(run);
            }

            if (!parseOnly)
            {
                string outDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "chart-batches");
                ChartBatchRunner.Run(run, batch, outDir);
                Console.WriteLine();
                Console.WriteLine("正式用例 通过 " + run.CasesPassed
                    + "  失败 " + run.CasesFailed
                    + "  跳过 " + run.CasesSkipped
                    + "  （目标 " + ChartCaseCatalog.Total + " 张图 / 6 份 PPT）");
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
                    && n <= ChartCaseCatalog.BatchCount)
                {
                    return n;
                }
            }

            return null;
        }
    }
}
