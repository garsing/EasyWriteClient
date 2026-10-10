using System;
using System.Diagnostics;
using WordAddIn1;

namespace WordDocumentRoundtripTest
{
    internal static class Program
    {
        [STAThread]
        private static int Main(string[] args)
        {
            Console.OutputEncoding = System.Text.Encoding.UTF8;
            string host = GetHost(args);
            string caseFilter = GetArg(args, "--case=");
            string docPath = GetArg(args, "--doc=");
            bool runFindStale = caseFilter.StartsWith("fs-", StringComparison.OrdinalIgnoreCase);
            bool runLayout = !runFindStale
                && (caseFilter.StartsWith("lb-", StringComparison.OrdinalIgnoreCase)
                    || caseFilter.StartsWith("wc-lb-", StringComparison.OrdinalIgnoreCase));
            bool runActions = !runFindStale && !runLayout
                && caseFilter.StartsWith("wa-", StringComparison.OrdinalIgnoreCase);
            bool runContent = !runFindStale && !runLayout && !runActions
                && caseFilter.StartsWith("wc-", StringComparison.OrdinalIgnoreCase);

            if (runFindStale)
            {
                Console.WriteLine("Find stale 探针（fs-）：游标推进，看是不是只在进表后返回游标前");
            }
            else if (runLayout)
            {
                Console.WriteLine("分页/分节 读入+写入（lb-）");
            }
            else if (runActions)
            {
                Console.WriteLine("文字 改字（wa-）");
            }
            else
            {
                Console.WriteLine(runContent ? "文字 读编码（wc-）" : "文字 打开 + 页数（wd-）");
            }

            Console.WriteLine("经 Desktop 管道 " + DesktopToolPipe.PipeName + " 调本机已开的易写");
            Console.WriteLine(host == "wps" ? "宿主 WPS 文字" : "宿主 Word");
            Console.WriteLine();

            if (!runFindStale && !DesktopToolPipe.TryWaitUntilAvailable(8000))
            {
                Console.WriteLine("连不上易写 Desktop。请先打开易写应用（需含 DesktopTools 管道的新版本），再跑本套。");
                return 1;
            }

            int wordBefore = OfficeAlive.Count("WINWORD");
            int wpsBefore = OfficeAlive.Count("wps");
            bool wordWinBefore = OfficeAlive.HasWindow("WINWORD");
            bool wpsWinBefore = OfficeAlive.HasWindow("wps");
            var run = new TestRun();
            var sw = Stopwatch.StartNew();
            if (runFindStale)
            {
                FindStaleCases.Run(run, host, caseFilter, docPath);
            }
            else if (runLayout)
            {
                LbCases.Run(run, host);
            }
            else if (runActions)
            {
                WaCases.Run(run, host, wordBefore, wpsBefore, wordWinBefore, wpsWinBefore);
            }
            else if (runContent)
            {
                WcCases.Run(run, host, wordBefore, wpsBefore, wordWinBefore, wpsWinBefore);
            }
            else
            {
                WdCases.Run(run, host, wordBefore, wpsBefore, wordWinBefore, wpsWinBefore);
            }

            sw.Stop();
            int target = runFindStale
                ? FindStaleCases.Target
                : (runLayout
                    ? LbCases.Target
                    : (runActions ? WaCases.Target : (runContent ? WcCases.Target : 30)));
            Console.WriteLine();
            Console.WriteLine("正式用例 通过 " + run.CasesPassed
                + "  失败 " + run.CasesFailed
                + "  跳过 " + run.CasesSkipped
                + "  （目标 " + target + "）");
            Console.WriteLine("ELAPSED_SEC=" + sw.Elapsed.TotalSeconds.ToString("F1"));
            return run.CasesFailed == 0 ? 0 : 1;
        }

        private static string GetHost(string[] args)
        {
            string v = GetArg(args, "--host=");
            if (string.Equals(v, "wps", StringComparison.OrdinalIgnoreCase))
            {
                return "wps";
            }

            return "word";
        }

        private static string GetArg(string[] args, string prefix)
        {
            if (args == null || string.IsNullOrEmpty(prefix))
            {
                return "";
            }

            for (int i = 0; i < args.Length; i++)
            {
                string a = args[i] ?? "";
                if (a.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
                {
                    return a.Substring(prefix.Length).Trim();
                }
            }

            return "";
        }
    }
}
