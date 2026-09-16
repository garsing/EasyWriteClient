using System;

namespace PptChartRoundtripTest
{
    internal static class Program
    {
        [STAThread]
        private static int Main(string[] args)
        {
            Console.OutputEncoding = System.Text.Encoding.UTF8;
            var run = new TestRun();
            bool parseOnly = args != null
                && args.Length > 0
                && string.Equals(args[0], "--parse-only", StringComparison.OrdinalIgnoreCase);

            Console.WriteLine("PPT 图表写读往返测试");
            Console.WriteLine("方法：写约定 HTML → 解析 / 建图或换数 → 读回对预期");
            Console.WriteLine();

            Console.WriteLine("--- 解析契约（不启 PPT）---");
            ChartParseTests.Run(run);

            if (parseOnly)
            {
                Console.WriteLine();
                Console.WriteLine("--- COM 往返已跳过（--parse-only）---");
            }
            else
            {
                Console.WriteLine();
                Console.WriteLine("--- COM 往返（需本机 PowerPoint）---");
                ChartRoundtripTests.Run(run);
            }

            Console.WriteLine();
            Console.WriteLine("通过 " + run.Passed + "  失败 " + run.Failed + "  跳过 " + run.Skipped);
            return run.Failed == 0 ? 0 : 1;
        }
    }
}
