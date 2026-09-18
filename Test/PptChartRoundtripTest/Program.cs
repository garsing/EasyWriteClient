using System;
using System.Collections.Generic;
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
            bool attrsOnly = HasFlag(args, "--attrs-only") || HasFlag(args, "--suite=attr");
            bool suiteAll = HasFlag(args, "--suite=all");
            int? batch = TryParseBatch(args);
            IList<string> nameFilters = ParseCaseFilters(args);
            bool runParse = !formalOnly && !attrsOnly && !suiteAll;
            bool runStruct = !parseOnly && (suiteAll || (!attrsOnly));
            bool runAttrs = !parseOnly && (attrsOnly || suiteAll);

            Console.WriteLine("PPT 图表测试");
            Console.WriteLine("结构用例：约 50 页 × 6 份；属性用例：单点+交叉 × 2 份（--attrs-only）");
            Console.WriteLine();

            if (runParse)
            {
                Console.WriteLine("--- 解析契约（不启 PPT，按属性断言）---");
                ChartParseTests.Run(run);
            }

            string outDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "chart-batches");
            if (runStruct)
            {
                ChartBatchRunner.Run(run, batch, outDir, nameFilters);
                Console.WriteLine();
                Console.WriteLine("结构用例 通过 " + run.CasesPassed
                    + "  失败 " + run.CasesFailed
                    + "  跳过 " + run.CasesSkipped
                    + "  （目标 " + ChartCaseCatalog.Total + " 张图 / 6 份 PPT）");
            }

            if (runAttrs)
            {
                int beforePass = run.CasesPassed;
                int beforeFail = run.CasesFailed;
                int beforeSkip = run.CasesSkipped;
                ChartBatchRunner.RunAttrs(run, batch, outDir, nameFilters);
                Console.WriteLine();
                Console.WriteLine("属性用例 通过 " + (run.CasesPassed - beforePass)
                    + "  失败 " + (run.CasesFailed - beforeFail)
                    + "  跳过 " + (run.CasesSkipped - beforeSkip)
                    + "  （目标 " + ChartAttrCatalog.Total + " 张图 / 2 份 PPT）");
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

        /// <summary>--case=title 或 --case=title,sink-column，用例名包含即可。</summary>
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
