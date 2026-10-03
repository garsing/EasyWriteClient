using System;
using System.Diagnostics;
using WordAddIn1;

namespace BrowserNavigateTest
{
    internal static class Program
    {
        [STAThread]
        private static int Main(string[] args)
        {
            Console.OutputEncoding = System.Text.Encoding.UTF8;
            if (string.Equals(GetArg(args, "--probe="), "jd-list", StringComparison.OrdinalIgnoreCase))
            {
                return JdListClickProbe.Run();
            }

            string caseFilter = GetArg(args, "--case=");
            bool runShot = caseFilter.StartsWith("shot-", StringComparison.OrdinalIgnoreCase);
            bool runDl = caseFilter.StartsWith("dl-", StringComparison.OrdinalIgnoreCase);
            bool runTab = caseFilter.StartsWith("tab-", StringComparison.OrdinalIgnoreCase);
            bool runInt = caseFilter.StartsWith("int-", StringComparison.OrdinalIgnoreCase);
            bool runNav = caseFilter.StartsWith("nav-", StringComparison.OrdinalIgnoreCase);
            bool runSnp = !runInt && !runNav && !runTab && !runDl && !runShot;

            string host = GetHost(args);
            if ((runInt || runTab || runDl || runShot) && string.Equals(host, "attach", StringComparison.OrdinalIgnoreCase))
            {
                Console.WriteLine(
                    (runShot ? "可视区截图" : (runDl ? "download" : (runTab ? "中间栏" : "interact")))
                    + " 本波只测 agent，忽略 --host=attach");
                host = "agent";
            }

            Console.WriteLine(
                runShot
                    ? "浏览器 可视区截图（shot-）"
                    : (runDl
                        ? "浏览器 download（dl-）"
                        : (runTab
                            ? "浏览器 中间栏 / 可点补全（tab-）"
                            : (runInt
                                ? "浏览器 interact（int-）"
                                : (runSnp ? "浏览器 snapshot + iframe（snp-）" : "浏览器 渠道 + navigate（nav-）")))));
            Console.WriteLine("经 Desktop 管道 " + DesktopToolPipe.PipeName + " 调本机已开的易写");
            Console.WriteLine(
                string.Equals(host, "attach", StringComparison.OrdinalIgnoreCase)
                    ? "宿主 attach · Chrome / Edge 扩展"
                    : "宿主 agent · 易写 WebView2");
            Console.WriteLine();

            if (!DesktopToolPipe.TryWaitUntilAvailable(8000))
            {
                Console.WriteLine("连不上易写 Desktop。请先打开易写应用（需含 DesktopTools 管道的新版本），再跑本套。");
                return 1;
            }

            int chromeBefore = CountProc("chrome");
            int edgeBefore = CountProc("msedge");
            var run = new TestRun();
            var sw = Stopwatch.StartNew();
            if (runShot)
            {
                using (var http = new FixtureHttp())
                {
                    ShotCases.Run(run, http, chromeBefore, edgeBefore);
                }
            }
            else if (runDl)
            {
                using (var http = new FixtureHttp())
                {
                    DlCases.Run(run, http, chromeBefore, edgeBefore);
                }
            }
            else if (runTab)
            {
                using (var http = new FixtureHttp())
                {
                    TabCases.Run(run, http, chromeBefore, edgeBefore);
                }
            }
            else if (runInt)
            {
                using (var http = new FixtureHttp())
                {
                    IntCases.Run(run, http, chromeBefore, edgeBefore);
                }
            }
            else if (runNav)
            {
                using (var http = new NavHttp())
                {
                    NavCases.Run(run, http, chromeBefore, edgeBefore);
                }
            }
            else
            {
                using (var http = new FixtureHttp())
                {
                    SnpCases.Run(
                        run,
                        http,
                        string.Equals(host, "attach", StringComparison.OrdinalIgnoreCase),
                        chromeBefore,
                        edgeBefore);
                }
            }

            sw.Stop();
            int target = runShot ? 30 : (runDl ? 80 : (runTab ? 75 : (runInt ? 160 : (runSnp ? 46 : 30))));
            Console.WriteLine();
            Console.WriteLine("正式用例 通过 " + run.CasesPassed
                + "  失败 " + run.CasesFailed
                + "  跳过 " + run.CasesSkipped
                + "  （目标 " + target + "）");
            Console.WriteLine("ELAPSED_SEC=" + sw.Elapsed.TotalSeconds.ToString("F1"));
            return run.CasesFailed == 0 ? 0 : 1;
        }

        private static string GetArg(string[] args, string prefix)
        {
            if (args == null)
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

        private static string GetHost(string[] args)
        {
            string v = GetArg(args, "--host=");
            if (v == "attach" || v == "chrome" || v == "edge")
            {
                return "attach";
            }

            return "agent";
        }

        private static int CountProc(string name)
        {
            try
            {
                return Process.GetProcessesByName(name).Length;
            }
            catch
            {
                return 0;
            }
        }
    }
}
