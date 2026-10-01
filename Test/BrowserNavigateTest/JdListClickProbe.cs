using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using WordAddIn1;

namespace BrowserNavigateTest
{
    /// <summary>
    /// 对着已开 Desktop，用 attach 打开京东列表并点封面图，打印树行与诊断日志。
    /// </summary>
    internal static class JdListClickProbe
    {
        public const string ListUrl =
            "https://list.jd.com/list.html?cat=9987,830,13660&sourcePage=itemBreadCrumb";

        public static int Run()
        {
            Console.OutputEncoding = System.Text.Encoding.UTF8;
            Console.WriteLine("京东列表点图诊断 · attach · " + DesktopToolPipe.PipeName);
            if (!DesktopToolPipe.TryWaitUntilAvailable(8000))
            {
                Console.WriteLine("连不上 Desktop。请先打开易写（含本管道的版本）。");
                return 1;
            }

            var s = new NavSession();
            ToolResult nav = s.Nav(ListUrl, host: "chrome");
            Console.WriteLine("--- navigate ---");
            PrintResult(nav);
            if (!nav.Success)
            {
                return 1;
            }

            string id = Str(nav, "channel_id");
            if (string.IsNullOrEmpty(id))
            {
                Console.WriteLine("没有 channel_id");
                return 1;
            }

            ToolResult snap = WaitForProductTree(s, id);
            Console.WriteLine("--- snapshot ---");
            Console.WriteLine("url=" + Str(snap, "url"));
            Console.WriteLine("title=" + Str(snap, "title"));
            string tree = Str(snap, "snapshot");
            PrintTreeHits(tree);

            string imageRef = FindCoverImageRef(tree);
            string priceRef = FindRefOnLine(tree, "价格区间");
            Console.WriteLine("价格区间 ref=" + (priceRef ?? "(无)"));
            Console.WriteLine("将要点的封面图 ref=" + (imageRef ?? "(无)"));
            if (string.IsNullOrEmpty(imageRef))
            {
                Console.WriteLine("树上没找到带商品名的 image，停止。");
                DumpTree(tree);
                return 1;
            }

            ToolResult click = s.Interact(id, "click", imageRef);
            Console.WriteLine("--- click " + imageRef + " ---");
            PrintResult(click);

            ToolResult after = s.Snap(id);
            Console.WriteLine("--- snapshot after ---");
            Console.WriteLine("url=" + Str(after, "url"));
            Console.WriteLine("title=" + Str(after, "title"));
            PrintTreeHits(Str(after, "snapshot"));
            PrintAfterHints(Str(after, "snapshot"));

            PrintDiagLog();
            return click.Success ? 0 : 1;
        }

        private static ToolResult WaitForProductTree(NavSession s, string id)
        {
            ToolResult snap = s.Snap(id);
            string tree = Str(snap, "snapshot");
            if (!string.IsNullOrEmpty(FindCoverImageRef(tree)))
            {
                return snap;
            }

            Console.WriteLine("商品树未齐，等 2s 再拍");
            System.Threading.Thread.Sleep(2000);
            snap = s.Snap(id);
            tree = Str(snap, "snapshot");
            if (!string.IsNullOrEmpty(FindCoverImageRef(tree)))
            {
                return snap;
            }

            Console.WriteLine("仍无封面，向下滚动再拍");
            s.Interact(id, "scroll", null);
            System.Threading.Thread.Sleep(800);
            return s.Snap(id);
        }

        private static void DumpTree(string tree)
        {
            string[] lines = (tree ?? "").Split('\n');
            Console.WriteLine("tree lines=" + lines.Length + " chars=" + (tree ?? "").Length);
            int n = Math.Min(60, lines.Length);
            for (int i = 0; i < n; i++)
            {
                Console.WriteLine(lines[i].TrimEnd());
            }

            foreach (string line in lines)
            {
                if (line.IndexOf("image", StringComparison.OrdinalIgnoreCase) >= 0
                    || line.IndexOf("登录", StringComparison.Ordinal) >= 0
                    || line.IndexOf("验证", StringComparison.Ordinal) >= 0
                    || line.IndexOf("iframe", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    Console.WriteLine(line.TrimEnd());
                }
            }
        }

        private static void PrintTreeHits(string tree)
        {
            foreach (string line in (tree ?? "").Split('\n'))
            {
                if (line.IndexOf("价格区间", StringComparison.Ordinal) >= 0
                    || (line.IndexOf("image", StringComparison.OrdinalIgnoreCase) >= 0
                        && line.IndexOf("ref=e", StringComparison.Ordinal) >= 0
                        && (line.IndexOf("车载", StringComparison.Ordinal) >= 0
                            || line.IndexOf("小米", StringComparison.Ordinal) >= 0)))
                {
                    Console.WriteLine(line.TrimEnd());
                }
            }
        }

        private static string FindCoverImageRef(string tree)
        {
            foreach (string line in (tree ?? "").Split('\n'))
            {
                if (line.IndexOf("image", StringComparison.OrdinalIgnoreCase) < 0)
                {
                    continue;
                }

                if (line.IndexOf("车载充电器", StringComparison.Ordinal) < 0
                    && line.IndexOf("小米（MI）", StringComparison.Ordinal) < 0)
                {
                    continue;
                }

                Match m = Regex.Match(line, @"ref=(e\d+)");
                if (m.Success)
                {
                    return m.Groups[1].Value;
                }
            }

            return null;
        }

        private static string FindRefOnLine(string tree, string text)
        {
            foreach (string line in (tree ?? "").Split('\n'))
            {
                if (line.IndexOf(text, StringComparison.Ordinal) < 0)
                {
                    continue;
                }

                Match m = Regex.Match(line, @"ref=(e\d+)");
                if (m.Success)
                {
                    return m.Groups[1].Value;
                }
            }

            return null;
        }

        private static void PrintAfterHints(string tree)
        {
            int price = CountLine(tree, "价格区间");
            int min = CountLine(tree, "最低价");
            int max = CountLine(tree, "最高价");
            int item = CountLine(tree, "item.jd.com");
            Console.WriteLine(
                "after hints 价格区间=" + price
                + " 最低价=" + min
                + " 最高价=" + max
                + " item.jd.com=" + item
                + " 行数=" + ((tree ?? "").Split('\n').Length));
        }

        private static int CountLine(string tree, string text)
        {
            int n = 0;
            foreach (string line in (tree ?? "").Split('\n'))
            {
                if (line.IndexOf(text, StringComparison.Ordinal) >= 0)
                {
                    n++;
                }
            }

            return n;
        }

        private static void PrintResult(ToolResult r)
        {
            if (r == null)
            {
                Console.WriteLine("(无回包)");
                return;
            }

            Console.WriteLine("ok=" + r.Success + " err=" + (r.Error ?? ""));
            Console.WriteLine("channel=" + Str(r, "channel_id")
                + " url=" + Str(r, "url")
                + " msg=" + Str(r, "message"));
        }

        private static void PrintDiagLog()
        {
            string path = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "YiWrite",
                "logs",
                "attach-click-diag.log");
            Console.WriteLine("--- " + path + " ---");
            if (!File.Exists(path))
            {
                Console.WriteLine("(还没有诊断文件：须用刚编过的 Desktop，扩展建议重载 0.1.26)");
                return;
            }

            string[] lines = File.ReadAllLines(path);
            int from = 0;
            for (int i = lines.Length - 1; i >= 0; i--)
            {
                if (lines[i].IndexOf(" SNAP ", StringComparison.Ordinal) >= 0)
                {
                    from = i;
                    break;
                }
            }

            if (from == 0)
            {
                from = Math.Max(0, lines.Length - 40);
            }

            for (int i = from; i < lines.Length; i++)
            {
                Console.WriteLine(lines[i]);
            }
        }

        private static string Str(ToolResult r, string key)
        {
            var d = r?.Data as Dictionary<string, object>;
            if (d == null || !d.ContainsKey(key) || d[key] == null)
            {
                return "";
            }

            return Convert.ToString(d[key]) ?? "";
        }
    }
}
