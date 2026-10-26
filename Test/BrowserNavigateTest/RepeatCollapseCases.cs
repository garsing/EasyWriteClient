using System;
using System.Collections.Generic;
using Newtonsoft.Json.Linq;
using WordAddIn1.BrowserHost;

namespace BrowserNavigateTest
{
    internal static class RepeatCollapseCases
    {
        public static int Run(string filter)
        {
            var cases = new List<Action>
            {
                Ac1,
                Ac2,
                Ac3,
                Ac4,
                Ac5,
                Ac6,
                Ac7,
                Ac8,
                Ac9,
                SignatureParagraphs,
                SignatureAlternation
            };

            int failed = 0;
            int ran = 0;
            foreach (Action act in cases)
            {
                string name = act.Method.Name;
                if (!string.Equals(filter, "repeat", StringComparison.OrdinalIgnoreCase)
                    && filter.IndexOf(name, StringComparison.OrdinalIgnoreCase) < 0)
                {
                    continue;
                }

                ran++;
                try
                {
                    act();
                    Console.WriteLine("PASS " + name);
                }
                catch (Exception ex)
                {
                    failed++;
                    Console.WriteLine("FAIL " + name + " " + ex.Message);
                }
            }

            if (ran == 0)
            {
                Console.WriteLine("没有匹配的 repeat 用例: " + filter);
                return 1;
            }

            Console.WriteLine(failed == 0 ? "全部通过 " + ran : "失败 " + failed + "/" + ran);
            return failed == 0 ? 0 : 1;
        }

        private static void Ac1()
        {
            string json = Page(10, header: false, split: false, hiddenSubmit: false);
            BrowserAxBuildResult built = BrowserAxTreeBuilder.BuildOverview(json);
            string tree = MustOk(built);
            MustContain(tree, "共 10 项，当前 1-10");
            Equal(CountOf(tree, "INNER"), 1, "只有窗口第一项展开");
            MustRef(tree, "Q002");
            MustRef(tree, "Q010");
            MustRef(tree, "加入作业");
            True(tree.IndexOf("repeat", StringComparison.Ordinal) < tree.IndexOf("加入作业", StringComparison.Ordinal), "加入作业在段外");
        }

        private static void Ac2()
        {
            string json = Page(100, header: false, split: false, hiddenSubmit: false);
            string tree = MustOk(BrowserAxTreeBuilder.BuildOverview(json));
            MustContain(tree, "共 100 项，当前 1-100");
            MustRef(tree, "Q100");
            MustRef(tree, "加入作业");
        }

        private static void Ac3()
        {
            string json = Page(10, header: true, split: false, hiddenSubmit: false);
            string tree = MustOk(BrowserAxTreeBuilder.BuildOverview(json));
            MustContain(tree, "表头");
            MustContain(tree, "repeat");
            True(tree.IndexOf("表头", StringComparison.Ordinal) < tree.IndexOf("repeat", StringComparison.Ordinal), "表头不进重复段");
        }

        private static void Ac4()
        {
            var nodes = new JArray();
            var kids = new List<string>();
            for (int i = 1; i <= 3; i++)
            {
                string id = "s" + i;
                kids.Add(id);
                nodes.Add(Node(id, "row", "浅" + i, 100 + i, "t" + i));
                nodes.Add(Node("t" + i, "StaticText", "浅文" + i, 0));
            }

            nodes.Add(Node("root", "RootWebArea", "页", 1, kids.ToArray()));
            string tree = MustOk(BrowserAxTreeBuilder.BuildOverview(Wrap(nodes)));
            True(tree.IndexOf("repeat", StringComparison.Ordinal) < 0, "浅节点不收");
        }

        private static void Ac5()
        {
            string json = Page(8, header: false, split: true, hiddenSubmit: false);
            string tree = MustOk(BrowserAxTreeBuilder.BuildOverview(json));
            MustContain(tree, "共 4 项");
            MustContain(tree, "共 3 项");
            MustContain(tree, "extra-btn");
            MustContain(tree, "SPECIAL");
        }

        private static void Ac6()
        {
            string json = Page(10, header: false, split: false, hiddenSubmit: false);
            BrowserAxBuildResult overview = BrowserAxTreeBuilder.BuildOverview(json);
            string tree = MustOk(overview);
            string refId = RefOnLine(tree, "Q002");
            True(overview.Refs.TryGetValue(refId, out BrowserRefEntry entry), "第 2 行有 ref");
            BrowserAxBuildResult detail = BrowserAxTreeBuilder.BuildDetail(json, refId, entry);
            string inner = MustOk(detail);
            MustContain(inner, "textbox");
            MustContain(inner, "checkbox");
        }

        private static void Ac7()
        {
            string json = Page(10, header: false, split: false, hiddenSubmit: true);
            string tree = MustOk(BrowserAxTreeBuilder.BuildOverview(json));
            MustContain(tree, "提交");
            MustContain(tree, "(隐藏)");
        }

        private static void Ac8()
        {
            string json = Page(400, header: false, split: false, hiddenSubmit: false);
            BrowserAxBuildResult overview = BrowserAxTreeBuilder.BuildOverview(json);
            MustOk(overview);
            BrowserRefEntry seg = RepeatEntry(overview);
            BrowserAxBuildResult again = BrowserAxTreeBuilder.BuildOverview(json, Window(seg, 301, 400));
            string tree = MustOk(again);
            MustContain(tree, "共 400 项，当前 301-400");
            MustContain(tree, "\"Q301\"");
            MustContain(tree, "\"Q400\"");
            True(tree.IndexOf("\"Q300\"", StringComparison.Ordinal) < 0, "不出现第 300 项");
            MustRef(tree, "加入作业");
        }

        private static void Ac9()
        {
            string json = Page(400, header: false, split: false, hiddenSubmit: false);
            BrowserAxBuildResult overview = BrowserAxTreeBuilder.BuildOverview(json);
            MustOk(overview);
            BrowserRefEntry seg = RepeatEntry(overview);
            string wide = MustOk(BrowserAxTreeBuilder.BuildOverview(json, Window(seg, 1, 500)));
            MustContain(wide, "共 400 项，当前 1-300");
            string tail = MustOk(BrowserAxTreeBuilder.BuildOverview(json, Window(seg, 350, 1000)));
            MustContain(tail, "共 400 项，当前 350-400");
            MustContain(tail, "\"Q350\"");
            MustContain(tail, "\"Q400\"");
            True(tail.IndexOf("\"Q349\"", StringComparison.Ordinal) < 0, "350 之前不印");
        }

        private static void SignatureParagraphs()
        {
            string tree = MustOk(BrowserAxTreeBuilder.BuildOverview(VariablePage(paragraphs: true)));
            MustContain(tree, "共 3 项");
        }

        private static void SignatureAlternation()
        {
            string tree = MustOk(BrowserAxTreeBuilder.BuildOverview(VariablePage(paragraphs: false)));
            MustContain(tree, "共 3 项");
        }

        private static BrowserAxTreeBuilder.RepeatWindowRequest Window(BrowserRefEntry seg, int start, int end)
        {
            return new BrowserAxTreeBuilder.RepeatWindowRequest
            {
                ParentBackendId = seg.RepeatParentBackendId ?? -1,
                Signature = seg.RepeatSignature,
                Ordinal = seg.RepeatOrdinal,
                Start = start,
                End = end
            };
        }

        private static string Page(int count, bool header, bool split, bool hiddenSubmit)
        {
            var nodes = new JArray();
            var kids = new List<string>();
            if (header)
            {
                kids.Add("hdr");
                nodes.Add(Node("hdr", "row", "表头", 50, "h1", "h2", "h3"));
                nodes.Add(Node("h1", "columnheader", "列1", 51));
                nodes.Add(Node("h2", "columnheader", "列2", 52));
                nodes.Add(Node("h3", "columnheader", "列3", 53));
            }

            for (int i = 1; i <= count; i++)
            {
                bool extra = split && i == 5;
                string id = "q" + i;
                kids.Add(id);
                AddQuestion(nodes, id, extra ? "SPECIAL" : "Q" + i.ToString("000"), 1000 + i, extra);
            }

            kids.Add("join");
            nodes.Add(Node("join", "link", "加入作业", 9));
            if (hiddenSubmit)
            {
                kids.Add("sub");
                nodes.Add(Node("sub", "button", "提交", 8, ignored: true));
            }

            nodes.Add(Node("root", "RootWebArea", "页", 1, kids.ToArray()));
            return Wrap(nodes);
        }

        private static void AddQuestion(JArray nodes, string id, string name, int backend, bool extra)
        {
            var childIds = new List<string>();
            childIds.Add(id + "tb");
            childIds.Add(id + "cb");
            childIds.Add(id + "in");
            nodes.Add(Node(id + "tb", "textbox", "answer", backend + 100000));
            nodes.Add(Node(id + "cb", "checkbox", "opt", backend + 200000));
            nodes.Add(Node(id + "in", "StaticText", "INNER", 0));
            for (int p = 0; p < 9; p++)
            {
                string pid = id + "p" + p;
                childIds.Add(pid);
                nodes.Add(Node(pid, "StaticText", "pad", 0));
            }

            if (extra)
            {
                childIds.Add(id + "x");
                nodes.Add(Node(id + "x", "button", "extra-btn", backend + 300000));
            }

            nodes.Add(Node(id, "row", name, backend, childIds.ToArray()));
        }

        private static string VariablePage(bool paragraphs)
        {
            var nodes = new JArray();
            var kids = new List<string>();
            int[] counts = { 4, 6, 4 };
            for (int i = 0; i < counts.Length; i++)
            {
                string id = "v" + i;
                kids.Add(id);
                var childIds = new List<string>();
                for (int n = 0; n < 12; n++)
                {
                    string nid = id + "n" + n;
                    childIds.Add(nid);
                    nodes.Add(Node(nid, paragraphs ? "StaticText" : "LineBreak", "pad", 0));
                }

                for (int n = 0; n < counts[i]; n++)
                {
                    if (!paragraphs)
                    {
                        string sid = id + "s" + n;
                        childIds.Add(sid);
                        nodes.Add(Node(sid, "StaticText", "词", 0));
                    }

                    string pid = id + "g" + n;
                    childIds.Add(pid);
                    nodes.Add(Node(pid, "paragraph", "段", 0));
                }

                nodes.Add(Node(id, "row", "V" + i, 4000 + i, childIds.ToArray()));
            }

            nodes.Add(Node("root", "RootWebArea", "页", 1, kids.ToArray()));
            return Wrap(nodes);
        }

        private static string Wrap(JArray nodes)
        {
            return new JObject { ["nodes"] = nodes }.ToString();
        }

        private static JObject Node(string id, string role, string name, int backend, params string[] kids)
        {
            return Node(id, role, name, backend, false, kids);
        }

        private static JObject Node(string id, string role, string name, int backend, bool ignored, params string[] kids)
        {
            var o = new JObject
            {
                ["nodeId"] = id,
                ["role"] = role,
                ["name"] = name ?? "",
                ["childIds"] = new JArray(kids ?? new string[0])
            };
            if (backend > 0)
            {
                o["backendDOMNodeId"] = backend;
            }

            if (ignored)
            {
                o["ignored"] = true;
            }

            return o;
        }

        private static string MustOk(BrowserAxBuildResult built)
        {
            if (built == null || !built.Success)
            {
                throw new InvalidOperationException(built == null ? "无结果" : (built.Error ?? "失败"));
            }

            return built.TreeText ?? "";
        }

        private static void MustContain(string tree, string text)
        {
            if (tree.IndexOf(text, StringComparison.Ordinal) < 0)
            {
                throw new InvalidOperationException("缺少 " + text + "\n" + Head(tree));
            }
        }

        private static void MustRef(string tree, string name)
        {
            string line = LineWith(tree, "\"" + name + "\"");
            if (line.IndexOf("ref=", StringComparison.Ordinal) < 0)
            {
                throw new InvalidOperationException(name + " 没有 ref\n" + line);
            }
        }

        private static string RefOnLine(string tree, string name)
        {
            string line = LineWith(tree, "\"" + name + "\"");
            int at = line.IndexOf("ref=", StringComparison.Ordinal);
            if (at < 0)
            {
                throw new InvalidOperationException(name + " 没有 ref");
            }

            int start = at + 4;
            int end = start;
            while (end < line.Length && line[end] != ' ' && line[end] != '"')
            {
                end++;
            }

            return line.Substring(start, end - start);
        }

        private static string LineWith(string tree, string token)
        {
            string[] lines = tree.Split(new[] { '\n' }, StringSplitOptions.RemoveEmptyEntries);
            foreach (string line in lines)
            {
                if (line.IndexOf(token, StringComparison.Ordinal) >= 0)
                {
                    return line;
                }
            }

            throw new InvalidOperationException("没有行 " + token + "\n" + Head(tree));
        }

        private static BrowserRefEntry RepeatEntry(BrowserAxBuildResult built)
        {
            foreach (KeyValuePair<string, BrowserRefEntry> pair in built.Refs)
            {
                if (pair.Value != null && pair.Value.IsRepeat)
                {
                    return pair.Value;
                }
            }

            throw new InvalidOperationException("没有 repeat ref");
        }

        private static int CountOf(string tree, string token)
        {
            int n = 0;
            int at = 0;
            while ((at = tree.IndexOf(token, at, StringComparison.Ordinal)) >= 0)
            {
                n++;
                at += token.Length;
            }

            return n;
        }

        private static void Equal(int actual, int expect, string why)
        {
            if (actual != expect)
            {
                throw new InvalidOperationException(why + " 实际 " + actual + " 期望 " + expect);
            }
        }

        private static void True(bool ok, string why)
        {
            if (!ok)
            {
                throw new InvalidOperationException(why);
            }
        }

        private static string Head(string tree)
        {
            if (string.IsNullOrEmpty(tree) || tree.Length <= 800)
            {
                return tree ?? "";
            }

            return tree.Substring(0, 800);
        }
    }
}
