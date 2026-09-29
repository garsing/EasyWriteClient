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

            string host = GetHost(args);
            Console.WriteLine("浏览器 渠道 + navigate（nav-）");
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
            using (var http = new NavHttp())
            {
                NavCases.Run(run, http, chromeBefore, edgeBefore);
            }

            sw.Stop();
            Console.WriteLine();
            Console.WriteLine("正式用例 通过 " + run.CasesPassed
                + "  失败 " + run.CasesFailed
                + "  跳过 " + run.CasesSkipped
                + "  （目标 30）");
            Console.WriteLine("ELAPSED_SEC=" + sw.Elapsed.TotalSeconds.ToString("F1"));
            return run.CasesFailed == 0 ? 0 : 1;
        }

        private static string GetHost(string[] args)
        {
            if (args == null)
            {
                return "agent";
            }

            for (int i = 0; i < args.Length; i++)
            {
                string a = args[i] ?? "";
                if (a.StartsWith("--host=", StringComparison.OrdinalIgnoreCase))
                {
                    string v = a.Substring(7).Trim().ToLowerInvariant();
                    if (v == "attach" || v == "chrome" || v == "edge")
                    {
                        return "attach";
                    }

                    return "agent";
                }
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
