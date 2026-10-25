using System;
using System.Collections.Generic;
using System.Diagnostics;
using WordAddIn1;

namespace ExcelWorkbookRoundtripTest
{
    internal static class Program
    {
        [STAThread]
        private static int Main(string[] args)
        {
            Console.OutputEncoding = System.Text.Encoding.UTF8;
            ChannelHost.Kind = ChannelHostKind.Desktop;
            bool useEt = IsEtHost(args);
            string caseFilter = GetCaseFilter(args);
            bool run550 = caseFilter.StartsWith("xlg550", StringComparison.OrdinalIgnoreCase);
            bool runChart = !run550 && ShouldRun(caseFilter, "xlg-");
            bool runImage = ShouldRun(caseFilter, "xi-");
            bool runPivot = ShouldRun(caseFilter, "xlp-");
            bool runCond = ShouldRun(caseFilter, "xlc-");
            bool runFmt = ShouldRun(caseFilter, "xlf-");
            bool runManage = ShouldRun(caseFilter, "xlm-");
            bool runStruct = ShouldRun(caseFilter, "xls-");
            bool runWrite = ShouldRun(caseFilter, "xlw-");
            bool runRead = ShouldRun(caseFilter, "xl-")
                && !runWrite && !runStruct && !runManage && !runFmt && !runCond && !runPivot && !runChart
                && !runImage;
            if (string.IsNullOrEmpty(caseFilter))
            {
                runChart = true;
                runPivot = false;
                runCond = false;
                runFmt = false;
                runManage = false;
                runStruct = false;
                runWrite = false;
                runRead = false;
            }

            if (run550)
            {
                Console.WriteLine("Excel / ET 图 550（xlg550）");
            }
            else if (runChart)
            {
                Console.WriteLine("Excel / ET 图（xlg-）");
            }
            else if (runImage)
            {
                Console.WriteLine("Excel / ET 贴图（xi-）");
            }
            else if (runPivot)
            {
                Console.WriteLine("Excel / ET 透视（xlp-）");
            }
            else if (runCond)
            {
                Console.WriteLine("Excel / ET 条件格式（xlc-）");
            }
            else if (runFmt)
            {
                Console.WriteLine("Excel / ET 套格式（xlf-）");
            }
            else if (runManage)
            {
                Console.WriteLine("Excel / ET 管表（xlm-）");
            }
            else if (runStruct)
            {
                Console.WriteLine("Excel / ET 格子结构（xls-）");
            }
            else if (runWrite)
            {
                Console.WriteLine("Excel / ET 写区域（xlw-）");
            }
            else
            {
                Console.WriteLine("Excel / ET 打开 + 概览 + 读一块（xl-）");
            }
            Console.WriteLine(useEt ? "宿主 ET · WPS 表格" : "宿主 Excel · Microsoft Excel");
            Console.WriteLine();

            var run = new TestRun();
            var sw = Stopwatch.StartNew();
            if (runRead)
            {
                XlCases.Run(run, useEt);
            }

            if (runWrite)
            {
                XlwCases.Run(run, useEt);
            }

            if (runStruct)
            {
                XlsCases.Run(run, useEt);
            }

            if (runManage)
            {
                XlmCases.Run(run, useEt);
            }

            if (runFmt)
            {
                XlfCases.Run(run, useEt);
            }

            if (runCond)
            {
                XlcCases.Run(run, useEt);
            }

            if (runPivot)
            {
                XlpCases.Run(run, useEt);
            }

            if (run550)
            {
                Xl550Suite.Run(run, useEt, args);
            }

            if (runChart)
            {
                XlgCases.Run(run, useEt);
            }

            if (runImage)
            {
                XiCases.Run(run, useEt);
            }

            sw.Stop();

            Console.WriteLine();
            int target = run550
                ? run.TotalTarget
                : (runImage ? 26 : ((runCond || runChart) ? 38 : 30));
            Console.WriteLine("正式用例 通过 " + run.CasesPassed
                + "  失败 " + run.CasesFailed
                + "  跳过 " + run.CasesSkipped
                + "  （目标 " + target + "）");
            Console.WriteLine("ELAPSED_SEC=" + sw.Elapsed.TotalSeconds.ToString("F1"));
            return run.CasesFailed == 0 ? 0 : 1;
        }

        private static string GetCaseFilter(string[] args)
        {
            if (args == null)
            {
                return "";
            }

            for (int i = 0; i < args.Length; i++)
            {
                string a = args[i] ?? "";
                if (a.StartsWith("--case=", StringComparison.OrdinalIgnoreCase))
                {
                    return a.Substring(7).Trim();
                }
            }

            return "";
        }

        private static bool ShouldRun(string caseFilter, string prefix)
        {
            if (string.IsNullOrEmpty(caseFilter))
            {
                return false;
            }

            return caseFilter.StartsWith(prefix, StringComparison.OrdinalIgnoreCase);
        }

        private static bool IsEtHost(string[] args)
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
                    return string.Equals(v, "et", StringComparison.OrdinalIgnoreCase)
                        || string.Equals(v, "wps", StringComparison.OrdinalIgnoreCase);
                }
            }

            return false;
        }
    }
}
