using System;
using System.Collections.Generic;
using WordAddIn1;

namespace BrowserNavigateTest
{
    /// <summary>Chrome 比 Edge 多红：dump native-select type 与 center-tabs 点 tab 的树。</summary>
    internal static class ChromeGapProbe
    {
        public static int Run()
        {
            Console.OutputEncoding = System.Text.Encoding.UTF8;
            Console.WriteLine("Chrome gap 诊断 · " + DesktopToolPipe.PipeName);
            if (!DesktopToolPipe.TryWaitUntilAvailable(8000))
            {
                Console.WriteLine("连不上 Desktop");
                return 1;
            }

            using (var http = new FixtureHttp())
            {
                var s = new NavSession();
                DumpType(s, http, "chrome");
                Console.WriteLine();
                DumpType(s, http, "edge");
            }

            return 0;
        }

        private static void DumpType(NavSession s, FixtureHttp http, string host)
        {
            Console.WriteLine("======== " + host + " native-select type 已填 ========");
            ToolResult nav = s.Nav(http.Url("native-select", "index.html"), host: host);
            PrintHead(nav, "nav");
            if (!nav.Success)
            {
                return;
            }

            string id = Str(nav, "channel_id");
            ToolResult snap = s.Snap(id, domSupplement: "input");
            string tree = Str(snap, "snapshot");
            PrintHead(snap, "snap");
            PrintFlags(snap);
            DumpLines(tree, "论文题目", "lwtm", "已填", "状态", "textbox", "value=");
            string box = IntCases.FindRef(tree, "论文题目") ?? IntCases.FindRef(tree, "lwtm");
            Console.WriteLine("FindRef 论文题目=" + (box ?? "(无)"));

            ToolResult typed = string.IsNullOrEmpty(box) ? null : s.Interact(id, "type", box, text: "已填");
            PrintHead(typed, "type");
            PrintFlags(typed);
            string after = typed == null ? "" : IntCases.Tree(typed);
            Console.WriteLine("type-tree has 已填=" + after.Contains("已填")
                + " chars=" + after.Length);
            DumpLines(after, "论文题目", "已填", "状态", "value=");

            ToolResult again = s.Snap(id, domSupplement: "input");
            PrintHead(again, "snap-after");
            string t2 = Str(again, "snapshot");
            Console.WriteLine("fresh-snap has 已填=" + t2.Contains("已填") + " chars=" + t2.Length);
            DumpLines(t2, "论文题目", "已填", "状态", "value=");
        }

        private static void DumpTab(NavSession s, FixtureHttp http)
        {
            Console.WriteLine("======== center-tabs 项目成员 / 项目详情 ========");
            ToolResult nav = s.Nav(http.Url("center-tabs", "index.html"), host: "chrome");
            PrintHead(nav, "nav");
            if (!nav.Success)
            {
                return;
            }

            string id = Str(nav, "channel_id");
            ToolResult snap = s.Snap(id);
            string tree = Str(snap, "snapshot");
            PrintHead(snap, "snap");
            DumpLines(tree, "项目", "状态", "ref=");
            string mem = IntCases.FindRef(tree, "项目成员");
            string det = IntCases.FindRef(tree, "项目详情");
            Console.WriteLine("FindRef 项目成员=" + (mem ?? "(无)") + " 项目详情=" + (det ?? "(无)"));

            ToolResult click = string.IsNullOrEmpty(mem) ? null : s.Interact(id, "click", mem);
            PrintHead(click, "click-member");
            PrintFlags(click);
            string after = click == null ? "" : IntCases.Tree(click);
            Console.WriteLine("after has 项目成员=" + after.Contains("项目成员")
                + " 状态: 项目成员=" + after.Contains("状态: 项目成员")
                + " chars=" + after.Length);
            DumpLines(after, "项目", "状态");
        }

        private static void PrintHead(ToolResult r, string tag)
        {
            if (r == null)
            {
                Console.WriteLine(tag + " (无回包)");
                return;
            }

            Console.WriteLine(tag
                + " ok=" + r.Success
                + " err=" + (r.Error ?? "")
                + " url=" + Str(r, "url")
                + " title=" + Str(r, "title")
                + " msg=" + Str(r, "message"));
        }

        private static void PrintFlags(ToolResult r)
        {
            if (r == null)
            {
                return;
            }

            Console.WriteLine("  resnapshot=" + Flag(r, "resnapshot")
                + " refs_invalidated=" + Flag(r, "refs_invalidated")
                + " truncated=" + Str(r, "truncated")
                + " reason=" + Str(r, "truncated_reason")
                + " mode=" + Str(r, "mode"));
        }

        private static void DumpLines(string tree, params string[] hints)
        {
            string[] lines = (tree ?? "").Split('\n');
            Console.WriteLine("  tree lines=" + lines.Length + " chars=" + (tree ?? "").Length);
            int shown = 0;
            foreach (string raw in lines)
            {
                string line = raw.TrimEnd();
                bool hit = shown < 12;
                if (!hit)
                {
                    for (int i = 0; i < hints.Length; i++)
                    {
                        if (line.IndexOf(hints[i], StringComparison.Ordinal) >= 0)
                        {
                            hit = true;
                            break;
                        }
                    }
                }

                if (!hit)
                {
                    continue;
                }

                Console.WriteLine("  " + line);
                shown++;
                if (shown > 40)
                {
                    Console.WriteLine("  ...");
                    break;
                }
            }
        }

        private static bool Flag(ToolResult r, string key)
        {
            return IntCases.Flag(r, key);
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
