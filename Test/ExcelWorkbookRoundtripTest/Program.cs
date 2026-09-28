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
            Console.WriteLine("Excel / ET 打开 + 概览 + 读一块（xl-）");
            Console.WriteLine(useEt ? "宿主 ET · WPS 表格" : "宿主 Excel · Microsoft Excel");
            Console.WriteLine();

            var run = new TestRun();
            var sw = Stopwatch.StartNew();
            XlCases.Run(run, useEt);
            sw.Stop();

            Console.WriteLine();
            Console.WriteLine("正式用例 通过 " + run.CasesPassed
                + "  失败 " + run.CasesFailed
                + "  跳过 " + run.CasesSkipped
                + "  （目标 30）");
            Console.WriteLine("ELAPSED_SEC=" + sw.Elapsed.TotalSeconds.ToString("F1"));
            return run.CasesFailed == 0 ? 0 : 1;
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
