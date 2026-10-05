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
            Console.WriteLine("文字 打开 + 页数（wd-）");
            Console.WriteLine("经 Desktop 管道 " + DesktopToolPipe.PipeName + " 调本机已开的易写");
            Console.WriteLine(host == "wps" ? "宿主 WPS 文字" : "宿主 Word");
            Console.WriteLine();

            if (!DesktopToolPipe.TryWaitUntilAvailable(8000))
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
            WdCases.Run(run, host, wordBefore, wpsBefore, wordWinBefore, wpsWinBefore);
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
                return "word";
            }

            for (int i = 0; i < args.Length; i++)
            {
                string a = args[i] ?? "";
                if (a.StartsWith("--host=", StringComparison.OrdinalIgnoreCase))
                {
                    string v = a.Substring("--host=".Length).Trim();
                    if (string.Equals(v, "wps", StringComparison.OrdinalIgnoreCase))
                    {
                        return "wps";
                    }
                }
            }

            return "word";
        }
    }
}
