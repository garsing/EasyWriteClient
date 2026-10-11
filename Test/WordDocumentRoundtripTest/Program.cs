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
            if (string.Equals(caseFilter, "wa-rollback-live", StringComparison.OrdinalIgnoreCase))
            {
                Console.WriteLine("文字 改字回滚探针（wa-rollback-live）");
                Console.WriteLine("经 Desktop 管道 " + DesktopToolPipe.PipeName + " 调本机已开的易写");
                return RollbackLiveProbe.Run(host);
            }

            if (string.Equals(caseFilter, "ckpt-skip-parse", StringComparison.OrdinalIgnoreCase))
            {
                Console.WriteLine("经 Desktop 管道 " + DesktopToolPipe.PipeName + " 调本机已开的易写");
                return ContractSkipCkptProbe.Run();
            }

            if (string.Equals(caseFilter, "wt-create-after-two", StringComparison.OrdinalIgnoreCase))
            {
                Console.WriteLine("文字 两句后建表短套（wt-create-after-two）");
                Console.WriteLine("经 Desktop 管道 " + DesktopToolPipe.PipeName + " 调本机已开的易写");
                Console.WriteLine(host == "wps" ? "宿主 WPS 文字" : "宿主 Word");
                if (!DesktopToolPipe.TryWaitUntilAvailable(8000))
                {
                    Console.WriteLine("连不上易写 Desktop。");
                    return 1;
                }

                int wordBeforeTwo = OfficeAlive.Count("WINWORD");
                int wpsBeforeTwo = OfficeAlive.Count("wps");
                bool wordWinBeforeTwo = OfficeAlive.HasWindow("WINWORD");
                bool wpsWinBeforeTwo = OfficeAlive.HasWindow("wps");
                var twoRun = new TestRun();
                WtCases.RunCreateAfterTwoOnly(
                    twoRun, host, wordBeforeTwo, wpsBeforeTwo, wordWinBeforeTwo, wpsWinBeforeTwo);
                Console.WriteLine("正式用例 通过 " + twoRun.CasesPassed
                    + "  失败 " + twoRun.CasesFailed
                    + "  跳过 " + twoRun.CasesSkipped);
                return twoRun.CasesFailed == 0 ? 0 : 1;
            }

            if (string.Equals(caseFilter, "wt-create-after", StringComparison.OrdinalIgnoreCase))
            {
                Console.WriteLine("文字 句后建表短套（wt-create-after）");
                Console.WriteLine("经 Desktop 管道 " + DesktopToolPipe.PipeName + " 调本机已开的易写");
                Console.WriteLine(host == "wps" ? "宿主 WPS 文字" : "宿主 Word");
                if (!DesktopToolPipe.TryWaitUntilAvailable(8000))
                {
                    Console.WriteLine("连不上易写 Desktop。");
                    return 1;
                }

                int wordBeforeAfter = OfficeAlive.Count("WINWORD");
                int wpsBeforeAfter = OfficeAlive.Count("wps");
                bool wordWinBeforeAfter = OfficeAlive.HasWindow("WINWORD");
                bool wpsWinBeforeAfter = OfficeAlive.HasWindow("wps");
                var afterRun = new TestRun();
                WtCases.RunCreateAfterOnly(
                    afterRun, host, wordBeforeAfter, wpsBeforeAfter, wordWinBeforeAfter, wpsWinBeforeAfter);
                Console.WriteLine("正式用例 通过 " + afterRun.CasesPassed
                    + "  失败 " + afterRun.CasesFailed
                    + "  跳过 " + afterRun.CasesSkipped);
                return afterRun.CasesFailed == 0 ? 0 : 1;
            }

            if (string.Equals(caseFilter, "wf-tbl-after", StringComparison.OrdinalIgnoreCase))
            {
                Console.WriteLine("文字 从文件插换短套（wf-tbl-after）");
                Console.WriteLine("经 Desktop 管道 " + DesktopToolPipe.PipeName + " 调本机已开的易写");
                Console.WriteLine(host == "wps" ? "宿主 WPS 文字" : "宿主 Word");
                if (!DesktopToolPipe.TryWaitUntilAvailable(8000))
                {
                    Console.WriteLine("连不上易写 Desktop。");
                    return 1;
                }

                var tableRun = new TestRun();
                WfCases.RunTableOnly(tableRun, host);
                return tableRun.CasesFailed == 0 ? 0 : 1;
            }

            bool runActions = !runFindStale && !runLayout
                && caseFilter.StartsWith("wa-", StringComparison.OrdinalIgnoreCase);
            bool runFromFile = !runFindStale && !runLayout && !runActions
                && caseFilter.StartsWith("wf-", StringComparison.OrdinalIgnoreCase);
            bool runFormat = !runFindStale && !runLayout && !runActions && !runFromFile
                && caseFilter.StartsWith("wg-", StringComparison.OrdinalIgnoreCase);
            bool runTable = !runFindStale && !runLayout && !runActions && !runFromFile && !runFormat
                && caseFilter.StartsWith("wt-", StringComparison.OrdinalIgnoreCase);
            bool runContent = !runFindStale && !runLayout && !runActions && !runFromFile && !runFormat && !runTable
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
            else if (runFromFile)
            {
                Console.WriteLine("文字 从文件插换（wf-）");
            }
            else if (runFormat)
            {
                Console.WriteLine("文字 套格式（wg-）");
            }
            else if (runTable)
            {
                Console.WriteLine("文字 表格（wt-）");
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
            else if (runFromFile)
            {
                WfCases.Run(run, host, wordBefore, wpsBefore, wordWinBefore, wpsWinBefore);
            }
            else if (runFormat)
            {
                WgCases.Run(run, host, wordBefore, wpsBefore, wordWinBefore, wpsWinBefore);
            }
            else if (runTable)
            {
                WtCases.Run(run, host, wordBefore, wpsBefore, wordWinBefore, wpsWinBefore);
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
                    : (runActions
                        ? WaCases.Target
                        : (runFromFile
                            ? WfCases.Target
                            : (runFormat
                                ? WgCases.Target
                                : (runTable
                                    ? WtCases.Target
                                    : (runContent ? WcCases.Target : 30))))));
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
