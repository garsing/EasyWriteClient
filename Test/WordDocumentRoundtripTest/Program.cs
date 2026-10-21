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
            var stdout = new System.IO.StreamWriter(Console.OpenStandardOutput(), Console.OutputEncoding) { AutoFlush = true };
            Console.SetOut(stdout);
            string host = GetHost(args);
            string caseFilter = GetArg(args, "--case=");
            string docPath = GetArg(args, "--doc=");
            if (caseFilter.StartsWith("xlk-", StringComparison.OrdinalIgnoreCase)
                || caseFilter.StartsWith("pk-", StringComparison.OrdinalIgnoreCase))
            {
                return RunOfficeCkpt(args, caseFilter);
            }
            bool runFindStale = caseFilter.StartsWith("fs-", StringComparison.OrdinalIgnoreCase);
            bool runLayout = !runFindStale
                && (caseFilter.StartsWith("lb-", StringComparison.OrdinalIgnoreCase)
                    || caseFilter.StartsWith("wc-lb-", StringComparison.OrdinalIgnoreCase));
            if (string.Equals(caseFilter, "wps-chart-rewrite-copyback", StringComparison.OrdinalIgnoreCase))
            {
                Console.WriteLine("WPS：存副本改 OOXML 再复制回用户窗口");
                if (!DesktopLauncher.EnsureAvailable())
                {
                    Console.WriteLine("连不上易写 Desktop。");
                    return 1;
                }

                return WhWpsChartOoxmlProbe.RunRewriteCopyBack();
            }

            if (string.Equals(caseFilter, "wps-chart-ooxml-read-dead", StringComparison.OrdinalIgnoreCase))
            {
                Console.WriteLine("WPS 贴图 Series 废了时，另存副本用 OOXML 读数");
                if (!DesktopLauncher.EnsureAvailable())
                {
                    Console.WriteLine("连不上易写 Desktop。");
                    return 1;
                }

                return WhWpsChartOoxmlProbe.RunOoxmlReadWhenSeriesDead();
            }

            if (string.Equals(caseFilter, "wps-chart-click-wake", StringComparison.OrdinalIgnoreCase))
            {
                Console.WriteLine("WPS 贴图后取屏幕坐标，模拟点击看能不能唤醒 Series");
                if (!DesktopLauncher.EnsureAvailable())
                {
                    Console.WriteLine("连不上易写 Desktop。");
                    return 1;
                }

                return WhWpsChartOoxmlProbe.RunClickWakeAfterPaste();
            }

            if (string.Equals(caseFilter, "wps-chart-wake-paste", StringComparison.OrdinalIgnoreCase))
            {
                Console.WriteLine("WPS 贴图后唤醒 Series，不关用户窗口");
                if (!DesktopLauncher.EnsureAvailable())
                {
                    Console.WriteLine("连不上易写 Desktop。");
                    return 1;
                }

                return WhWpsChartOoxmlProbe.RunWakeAfterPaste();
            }

            if (string.Equals(caseFilter, "wps-chart-auto-copy", StringComparison.OrdinalIgnoreCase))
            {
                Console.WriteLine("把带图稿的图用 COM 复制到空稿");
                if (!DesktopLauncher.EnsureAvailable())
                {
                    Console.WriteLine("连不上易写 Desktop。");
                    return 1;
                }

                return WhWpsChartOoxmlProbe.RunAutoCopyToBlank();
            }

            if (string.Equals(caseFilter, "wps-chart-inspect-blank", StringComparison.OrdinalIgnoreCase))
            {
                Console.WriteLine("检查手动粘贴后的空稿是否仍是活图");
                if (!DesktopLauncher.EnsureAvailable())
                {
                    Console.WriteLine("连不上易写 Desktop。");
                    return 1;
                }

                return WhWpsChartOoxmlProbe.RunInspectManualDest();
            }

            if (string.Equals(caseFilter, "wps-chart-manual-samples", StringComparison.OrdinalIgnoreCase))
            {
                Console.WriteLine("WPS 文字图：打开两份样例供手动复制");
                if (!DesktopLauncher.EnsureAvailable())
                {
                    Console.WriteLine("连不上易写 Desktop。");
                    return 1;
                }

                return WhWpsChartOoxmlProbe.RunManualSamples();
            }

            if (string.Equals(caseFilter, "wps-chart-why-dead", StringComparison.OrdinalIgnoreCase))
            {
                Console.WriteLine("WPS 文字图：插入后 Series 废了是包丢了还是 COM 废了");
                Console.WriteLine("不 Quit 用户 WPS / Word；只动临时稿");
                if (!DesktopLauncher.EnsureAvailable())
                {
                    Console.WriteLine("连不上易写 Desktop。");
                    return 1;
                }

                return WhWpsChartOoxmlProbe.RunWhyDead();
            }

            if (string.Equals(caseFilter, "wps-chart-ooxml-probe", StringComparison.OrdinalIgnoreCase))
            {
                Console.WriteLine("WPS 文字图 OOXML / InsertFile 探针（wps-chart-ooxml-probe）");
                Console.WriteLine("经 Desktop 管道 " + DesktopToolPipe.PipeName + " 调本机已开的易写");
                Console.WriteLine("不 Quit 用户 WPS / Word；只动临时稿");
                if (!DesktopLauncher.EnsureAvailable())
                {
                    Console.WriteLine("连不上易写 Desktop。");
                    return 1;
                }

                return WhWpsChartOoxmlProbe.Run();
            }

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

            if (string.Equals(caseFilter, "wt-wps-empty-slots", StringComparison.OrdinalIgnoreCase))
            {
                Console.WriteLine("WPS 巩固表空槽回包（wt-wps-empty-slots）");
                if (!DesktopToolPipe.TryWaitUntilAvailable(8000))
                {
                    Console.WriteLine("连不上易写 Desktop。");
                    return 1;
                }

                WtCases.RunEmptySlotDump(host);
                return 0;
            }

            if (string.Equals(caseFilter, "wt-before-tbl-join", StringComparison.OrdinalIgnoreCase))
            {
                Console.WriteLine("文字 表前（句后）建表是否并表（wt-before-tbl-join）");
                if (!DesktopToolPipe.TryWaitUntilAvailable(8000))
                {
                    Console.WriteLine("连不上易写 Desktop。");
                    return 1;
                }

                WtCases.RunBeforeTblJoinProbe(host);
                return 0;
            }

            if (string.Equals(caseFilter, "wt-after-tbl-join", StringComparison.OrdinalIgnoreCase))
            {
                Console.WriteLine("文字 表后建表是否并表（wt-after-tbl-join）");
                Console.WriteLine("经 Desktop 管道 " + DesktopToolPipe.PipeName + " 调本机已开的易写");
                if (!DesktopToolPipe.TryWaitUntilAvailable(8000))
                {
                    Console.WriteLine("连不上易写 Desktop。");
                    return 1;
                }

                WtCases.RunAfterTblJoinProbe(host);
                return 0;
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
            bool runPptChart = !runFindStale && !runLayout && !runActions && !runFromFile && !runFormat && !runTable
                && caseFilter.StartsWith("wh-ppt", StringComparison.OrdinalIgnoreCase);
            bool runChart = !runFindStale && !runLayout && !runActions && !runFromFile && !runFormat && !runTable
                && !runPptChart
                && caseFilter.StartsWith("wh-", StringComparison.OrdinalIgnoreCase);
            bool runImage = !runFindStale && !runLayout && !runActions && !runFromFile && !runFormat && !runTable
                && !runChart
                && (caseFilter.StartsWith("wi-", StringComparison.OrdinalIgnoreCase)
                    || caseFilter.StartsWith("wp-", StringComparison.OrdinalIgnoreCase));
            bool runCkpt = !runFindStale && !runLayout && !runActions && !runFromFile && !runFormat && !runTable
                && !runChart && !runImage
                && (caseFilter.StartsWith("wk-", StringComparison.OrdinalIgnoreCase)
                    || caseFilter.StartsWith("wv-", StringComparison.OrdinalIgnoreCase));
            bool runContent = !runFindStale && !runLayout && !runActions && !runFromFile && !runFormat && !runTable
                && !runChart && !runImage && !runCkpt
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
            else if (runChart)
            {
                Console.WriteLine("文字 图表（wh-）");
            }
            else if (runImage)
            {
                Console.WriteLine("文字 插图与页图（wi-/wp-）");
            }
            else if (runCkpt)
            {
                Console.WriteLine(WkCases.IsCaretOnly(caseFilter)
                    ? "文字 检查点光标/视口（wk-rst-caret）"
                    : "文字 检查点与溯源（wk-/wv-）");
            }
            else
            {
                Console.WriteLine(runContent ? "文字 读编码（wc-）" : "文字 打开 + 页数（wd-）");
            }

            Console.WriteLine("经 Desktop 管道 " + DesktopToolPipe.PipeName + " 调本机已开的易写");
            Console.WriteLine(host == "wps" ? "宿主 WPS 文字" : "宿主 Word");
            Console.WriteLine();

            if (!runFindStale && !DesktopLauncher.EnsureAvailable())
            {
                Console.WriteLine("连不上易写 Desktop（已尝试自动拉起）。请手动打开易写应用，再跑本套。");
                return 1;
            }

            int wordBefore = OfficeAlive.Count("WINWORD");
            int wpsBefore = OfficeAlive.Count("wps");
            bool wordWinBefore = OfficeAlive.HasWindow("WINWORD");
            bool wpsWinBefore = OfficeAlive.HasWindow("wps");
            WdSession.SweepBeforeRun(host);
            var run = new TestRun
            {
                TotalTarget = runFindStale
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
                                        : (runChart
                                            ? WhCases.Target
                                            : (runImage ? WiCases.Target : (runCkpt
                                                ? (WkCases.IsCaretOnly(caseFilter) ? WkCases.CaretTarget : WkCases.Target)
                                                : (runContent ? WcCases.Target : 30)))))))))
            };
            var sw = Stopwatch.StartNew();
            try
            {
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
                else if (runPptChart)
                {
                    if (string.Equals(caseFilter, "wh-ppt-style-sweep", StringComparison.OrdinalIgnoreCase))
                    {
                        WhPptSuite.StyleSweep(run, host);
                    }
                    else
                    {
                        WhPptSuite.Run(run, host, caseFilter, GetArg(args, "--wh-from="), GetArg(args, "--wh-ids="), GetArg(args, "--wh-from-id="));
                    }
                }
                else if (runChart)
                {
                    WhCases.Run(run, host, wordBefore, wpsBefore, wordWinBefore, wpsWinBefore, GetArg(args, "--wh-part="));
                }
                else if (runImage)
                {
                    WiCases.Run(run, host, wordBefore, wpsBefore, wordWinBefore, wpsWinBefore);
                }
                else if (runCkpt)
                {
                    WkCases.Run(run, host, wordBefore, wpsBefore, wordWinBefore, wpsWinBefore, caseFilter);
                }
                else if (runContent)
                {
                    WcCases.Run(run, host, wordBefore, wpsBefore, wordWinBefore, wpsWinBefore);
                }
                else
                {
                    WdCases.Run(run, host, wordBefore, wpsBefore, wordWinBefore, wpsWinBefore);
                }
            }
            finally
            {
                sw.Stop();
                WdSession.SweepAfterRun(host);
            }
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
                                        : (runPptChart
                                            ? WhPptSuite.Target
                                            : (runChart
                                            ? WhCases.Target
                                            : (runImage ? WiCases.Target : (runCkpt
                                                ? (WkCases.IsCaretOnly(caseFilter) ? WkCases.CaretTarget : WkCases.Target)
                                                : (runContent ? WcCases.Target : 30))))))))));
            Console.WriteLine();
            Console.WriteLine("正式用例 通过 " + run.CasesPassed
                + "  失败 " + run.CasesFailed
                + "  跳过 " + run.CasesSkipped
                + "  （目标 " + target + "）");
            Console.WriteLine("ELAPSED_SEC=" + sw.Elapsed.TotalSeconds.ToString("F1"));
            return run.CasesFailed == 0 ? 0 : 1;
        }

        private static int RunOfficeCkpt(string[] args, string caseFilter)
        {
            bool runXlk = caseFilter.StartsWith("xlk-", StringComparison.OrdinalIgnoreCase);
            string rawHost = GetArg(args, "--host=");
            string host = runXlk
                ? OfficeCkptSession.NormalizeExcelHost(rawHost)
                : OfficeCkptSession.NormalizePptHost(rawHost);
            Console.WriteLine(runXlk ? "表格 检查点（xlk-）" : "演示 检查点（pk-）");
            Console.WriteLine("经 Desktop 管道 " + DesktopToolPipe.PipeName + " 调本机已开的易写");
            if (runXlk)
            {
                Console.WriteLine(host == "et" ? "宿主 WPS 表格" : "宿主 Excel");
            }
            else
            {
                Console.WriteLine(host == "wpp" ? "宿主 WPS 演示" : "宿主 PowerPoint");
            }

            Console.WriteLine();
            if (!DesktopLauncher.EnsureAvailable())
            {
                Console.WriteLine("连不上易写 Desktop（已尝试自动拉起）。请手动打开易写应用，再跑本套。");
                return 1;
            }

            var s = new OfficeCkptSession(host);
            s.SweepLeftovers();
            bool winBefore = OfficeAlive.HasWindow(s.ProcName);
            var run = new TestRun { TotalTarget = runXlk ? XlkCases.Target : PkCases.Target };
            var sw = Stopwatch.StartNew();
            if (runXlk)
            {
                XlkCases.Run(run, host, winBefore);
            }
            else
            {
                PkCases.Run(run, host, winBefore);
            }

            sw.Stop();
            Console.WriteLine();
            Console.WriteLine("正式用例 通过 " + run.CasesPassed
                + "  失败 " + run.CasesFailed
                + "  跳过 " + run.CasesSkipped
                + "  （目标 " + run.TotalTarget + "）");
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
