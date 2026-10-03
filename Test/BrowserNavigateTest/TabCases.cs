using System;
using WordAddIn1;

namespace BrowserNavigateTest
{
    /// <summary>第 4 波中间栏 Tab / 可点补全，只测 agent。</summary>
    internal static class TabCases
    {
        private static readonly string[] AllIds =
        {
            "tab-snap-member", "tab-snap-detail", "tab-snap-task", "tab-snap-side",
            "tab-snap-status-detail", "tab-snap-no-assign", "tab-snap-no-rows", "tab-snap-longcard",
            "tab-click-member", "tab-click-resnapshot", "tab-click-member-assign",
            "tab-click-member-rows", "tab-click-member-row-ref", "tab-click-member-table",
            "tab-side-after-member", "tab-tabs-after-member", "tab-longcard-after-click",
            "tab-use-new-tree", "tab-click-task-hides-rows", "tab-click-detail",
            "tab-switch-back-detail", "tab-click-side", "tab-click-assign",
            "tab-click-member-twice", "tab-hidden-click-member"
        };

        public static void Run(TestRun run, FixtureHttp http, int chromeBefore, int edgeBefore)
        {
            var s = new NavSession();
            RunMain(run, s, http);
            RunHidden(run, s, http);
            int chromeNow = IntCases.CountProc("chrome");
            int edgeNow = IntCases.CountProc("msedge");
            if (chromeNow < chromeBefore || edgeNow < edgeBefore)
            {
                Console.WriteLine(
                    "  [WARN] 用户浏览器进程变少 chrome "
                    + chromeBefore + "→" + chromeNow
                    + " edge " + edgeBefore + "→" + edgeNow);
            }
        }

        private static void RunMain(TestRun run, NavSession s, FixtureHttp http)
        {
            string id = IntCases.OpenOk(run, s, http.Url("center-tabs", "index.html"), "tab-snap-member");
            if (id == null)
            {
                IntCases.FailMany(run, "依赖 center-tabs", RestAfter("tab-snap-member"));
                return;
            }

            string snap = IntCases.WaitTree(s, id, false, "项目成员", "评教列表", "项目详情");
            AssertSnap(run, snap);

            string mem = TabRef(snap, "项目成员");
            if (string.IsNullOrEmpty(mem))
            {
                IntCases.FailMany(run, "空拍无项目成员 ref",
                    "tab-click-member", "tab-click-resnapshot", "tab-click-member-assign",
                    "tab-click-member-rows", "tab-click-member-row-ref", "tab-click-member-table",
                    "tab-side-after-member", "tab-tabs-after-member", "tab-longcard-after-click",
                    "tab-use-new-tree", "tab-click-task-hides-rows", "tab-click-detail",
                    "tab-switch-back-detail", "tab-click-side", "tab-click-assign",
                    "tab-click-member-twice");
                return;
            }

            ToolResult clicked = s.Interact(id, "click", mem);
            string after = clicked == null ? "" : IntCases.Tree(clicked);
            bool okClick = clicked != null && clicked.Success && after.Contains("项目成员");
            IntCases.Ok(run, "tab-click-member", okClick, clicked == null ? "无回包" : (clicked.Error ?? "状态未变"));
            IntCases.Ok(
                run,
                "tab-click-resnapshot",
                clicked != null && clicked.Success
                    && after.Length > 0
                    && IntCases.Flag(clicked, "resnapshot")
                    && !IntCases.Flag(clicked, "refs_invalidated"),
                clicked == null ? "无回包" : (clicked.Error ?? "未带回树"));

            if (!okClick)
            {
                IntCases.FailMany(run, "点项目成员失败",
                    "tab-click-member-assign", "tab-click-member-rows", "tab-click-member-row-ref",
                    "tab-click-member-table", "tab-side-after-member", "tab-tabs-after-member",
                    "tab-longcard-after-click", "tab-use-new-tree", "tab-click-task-hides-rows",
                    "tab-click-detail", "tab-switch-back-detail", "tab-click-side",
                    "tab-click-assign", "tab-click-member-twice");
                return;
            }

            IntCases.Ok(
                run,
                "tab-click-member-assign",
                !string.IsNullOrEmpty(TabRef(after, "项目成员分配")),
                "新树无项目成员分配 ref");
            bool hasHuang = after.Contains("黄浩然");
            bool hasLiu = after.Contains("刘翔");
            IntCases.Ok(run, "tab-click-member-rows", hasHuang && hasLiu,
                !hasHuang && !hasLiu ? "两行都无" : (!hasHuang ? "无黄浩然" : "无刘翔"));
            string huangRef = IntCases.FindRef(after, "黄浩然");
            string liuRef = IntCases.FindRef(after, "刘翔");
            IntCases.Ok(
                run,
                "tab-click-member-row-ref",
                !string.IsNullOrEmpty(huangRef) && !string.IsNullOrEmpty(liuRef),
                string.IsNullOrEmpty(huangRef) ? "黄浩然无 ref" : "刘翔无 ref");
            IntCases.Ok(
                run,
                "tab-click-member-table",
                after.Contains("学号") || after.Contains("学生数=2"),
                "无学号/学生数");
            IntCases.Ok(
                run,
                "tab-side-after-member",
                !string.IsNullOrEmpty(TabRef(after, "评教列表")),
                "侧栏评教列表丢了");
            IntCases.Ok(
                run,
                "tab-tabs-after-member",
                !string.IsNullOrEmpty(TabRef(after, "项目详情"))
                    && !string.IsNullOrEmpty(TabRef(after, "项目成员"))
                    && !string.IsNullOrEmpty(TabRef(after, "项目任务")),
                "点完后三个 Tab 缺 ref");
            IntCases.Ok(run, "tab-longcard-after-click", !LongCardIsShortButton(after), "长卡片收成短按钮");

            string taskRef = TabRef(after, "项目任务");
            ToolResult task = string.IsNullOrEmpty(taskRef) ? null : s.Interact(id, "click", taskRef);
            string taskTree = task == null ? "" : IntCases.Tree(task);
            IntCases.Ok(
                run,
                "tab-use-new-tree",
                task != null && task.Success && taskTree.Contains("项目任务") && taskTree.Contains("任务列表"),
                task == null ? "新树无项目任务 ref" : (task.Error ?? "未到任务"));
            IntCases.Ok(
                run,
                "tab-click-task-hides-rows",
                task != null && task.Success && !VisibleName(taskTree, "黄浩然") && !VisibleName(taskTree, "刘翔"),
                task == null ? "依赖点任务" : "切到任务后两行仍可见\n" + DumpNameLines(taskTree, "黄浩然", "刘翔"));

            string afterTask = task != null && task.Success ? taskTree : after;
            string detailRef = TabRef(afterTask, "项目详情");
            if (string.IsNullOrEmpty(detailRef))
            {
                afterTask = IntCases.WaitTree(s, id, false, "项目详情");
                detailRef = TabRef(afterTask, "项目详情");
            }

            ToolResult detail = string.IsNullOrEmpty(detailRef) ? null : s.Interact(id, "click", detailRef);
            string detailTree = detail == null ? "" : IntCases.Tree(detail);
            IntCases.Ok(
                run,
                "tab-click-detail",
                detail != null && detail.Success && detailTree.Contains("项目详情"),
                detail == null ? "无项目详情 ref" : (detail.Error ?? "状态未回详情"));
            IntCases.Ok(
                run,
                "tab-switch-back-detail",
                detail != null && detail.Success
                    && !VisibleName(detailTree, "项目成员分配")
                    && !VisibleName(detailTree, "黄浩然")
                    && !VisibleName(detailTree, "刘翔"),
                detail == null ? "依赖点详情" : "切回详情后成员仍可见\n" + DumpNameLines(detailTree, "项目成员分配", "黄浩然", "刘翔"));

            string sideTree = detail != null && detail.Success ? detailTree : afterTask;
            string sideRef = TabRef(sideTree, "评教列表");
            if (string.IsNullOrEmpty(sideRef))
            {
                sideTree = IntCases.WaitTree(s, id, false, "评教列表");
                sideRef = TabRef(sideTree, "评教列表");
            }

            ToolResult side = string.IsNullOrEmpty(sideRef) ? null : s.Interact(id, "click", sideRef);
            string sideGot = side == null ? "" : IntCases.Tree(side);
            IntCases.Ok(
                run,
                "tab-click-side",
                side != null && side.Success && sideGot.Contains("侧栏") && sideGot.Contains("评教列表"),
                side == null ? "无评教列表 ref" : (side.Error ?? "侧栏状态未变"));

            string memAgain = TabRef(side != null && side.Success ? sideGot : sideTree, "项目成员");
            if (string.IsNullOrEmpty(memAgain))
            {
                memAgain = TabRef(IntCases.WaitTree(s, id, false, "项目成员"), "项目成员");
            }

            ToolResult member2 = string.IsNullOrEmpty(memAgain) ? null : s.Interact(id, "click", memAgain);
            string member2Tree = member2 == null ? "" : IntCases.Tree(member2);
            string assignRef = TabRef(member2Tree, "项目成员分配");
            ToolResult assign = string.IsNullOrEmpty(assignRef) ? null : s.Interact(id, "click", assignRef);
            IntCases.Ok(
                run,
                "tab-click-assign",
                assign != null && assign.Success,
                assign == null ? "无项目成员分配 ref" : assign.Error);

            string twiceTree = assign != null && assign.Success ? IntCases.Tree(assign) : member2Tree;
            string mem3 = TabRef(twiceTree, "项目成员");
            if (string.IsNullOrEmpty(mem3))
            {
                mem3 = TabRef(IntCases.WaitTree(s, id, false, "项目成员"), "项目成员");
            }

            ToolResult twice = string.IsNullOrEmpty(mem3) ? null : s.Interact(id, "click", mem3);
            string twiceGot = twice == null ? "" : IntCases.Tree(twice);
            IntCases.Ok(
                run,
                "tab-click-member-twice",
                twice != null && twice.Success
                    && twiceGot.Contains("黄浩然")
                    && twiceGot.Contains("刘翔"),
                twice == null ? "无项目成员 ref" : (twice.Error ?? "再点后两行丢了"));
        }

        private static void RunHidden(TestRun run, NavSession s, FixtureHttp http)
        {
            string hid = IntCases.OpenOk(
                run, s, http.Url("center-tabs", "index.html"), "tab-hidden-click-member", visible: false);
            if (hid == null)
            {
                return;
            }

            string tree = IntCases.WaitTree(s, hid, false, "项目成员");
            string mem = TabRef(tree, "项目成员");
            ToolResult r = string.IsNullOrEmpty(mem) ? null : s.Interact(hid, "click", mem);
            string got = r == null ? "" : IntCases.Tree(r);
            IntCases.Ok(
                run,
                "tab-hidden-click-member",
                r != null && r.Success && got.Contains("项目成员")
                    && got.Contains("黄浩然") && got.Contains("刘翔"),
                r == null ? "隐藏窗无项目成员 ref" : (r.Error ?? "隐藏窗点完两行不齐"));
        }

        private static void AssertSnap(TestRun run, string snap)
        {
            IntCases.Ok(run, "tab-snap-member", !string.IsNullOrEmpty(TabRef(snap, "项目成员")), "无项目成员 ref");
            IntCases.Ok(run, "tab-snap-detail", !string.IsNullOrEmpty(TabRef(snap, "项目详情")), "无项目详情 ref");
            IntCases.Ok(run, "tab-snap-task", !string.IsNullOrEmpty(TabRef(snap, "项目任务")), "无项目任务 ref");
            IntCases.Ok(run, "tab-snap-side", !string.IsNullOrEmpty(TabRef(snap, "评教列表")), "无评教列表 ref");
            IntCases.Ok(run, "tab-snap-status-detail", snap.Contains("项目详情"), "默认状态不是项目详情");
            IntCases.Ok(run, "tab-snap-no-assign", !VisibleName(snap, "项目成员分配"), "未点 Tab 已见项目成员分配");
            IntCases.Ok(
                run,
                "tab-snap-no-rows",
                !VisibleName(snap, "黄浩然") && !VisibleName(snap, "刘翔"),
                "未点 Tab 已见表行\n" + DumpNameLines(snap, "黄浩然", "刘翔"));
            IntCases.Ok(run, "tab-snap-longcard", !LongCardIsShortButton(snap), "长卡片收成短按钮");
        }

        private static string TabRef(string tree, string label)
        {
            if (label == "项目成员")
            {
                return IntCases.FindRef(tree, "\"项目成员\"")
                    ?? FindLineRef(tree, "项目成员", skip: "分配");
            }

            if (label == "项目成员分配")
            {
                return IntCases.FindRef(tree, "项目成员分配");
            }

            return IntCases.FindRef(tree, "\"" + label + "\"")
                ?? IntCases.FindRef(tree, label);
        }

        private static string FindLineRef(string tree, string hint, string skip)
        {
            if (string.IsNullOrEmpty(tree))
            {
                return null;
            }

            foreach (string line in tree.Split('\n'))
            {
                if (line.IndexOf(hint, StringComparison.Ordinal) < 0)
                {
                    continue;
                }

                if (!string.IsNullOrEmpty(skip) && line.IndexOf(skip, StringComparison.Ordinal) >= 0)
                {
                    continue;
                }

                var m = System.Text.RegularExpressions.Regex.Match(line, @"ref=(e\d+)");
                if (m.Success)
                {
                    return m.Groups[1].Value;
                }
            }

            return null;
        }

        private static string DumpNameLines(string tree, params string[] names)
        {
            if (string.IsNullOrEmpty(tree) || names == null || names.Length == 0)
            {
                return "";
            }

            var sb = new System.Text.StringBuilder();
            foreach (string line in tree.Split('\n'))
            {
                for (int i = 0; i < names.Length; i++)
                {
                    if (!string.IsNullOrEmpty(names[i]) && line.IndexOf(names[i], StringComparison.Ordinal) >= 0)
                    {
                        sb.Append("         | ").Append(line.Trim()).AppendLine();
                        break;
                    }
                }
            }

            return sb.ToString();
        }

        private static bool VisibleName(string tree, string name)
        {
            if (string.IsNullOrEmpty(tree) || tree.IndexOf(name, StringComparison.Ordinal) < 0)
            {
                return false;
            }

            foreach (string line in tree.Split('\n'))
            {
                if (line.IndexOf(name, StringComparison.Ordinal) < 0)
                {
                    continue;
                }

                if (line.IndexOf("(隐藏)", StringComparison.Ordinal) >= 0)
                {
                    continue;
                }

                return true;
            }

            return false;
        }

        private static bool LongCardIsShortButton(string tree)
        {
            const string prefix = "这是一张很长的可点卡片说明文字";
            if (string.IsNullOrEmpty(tree))
            {
                return false;
            }

            foreach (string line in tree.Split('\n'))
            {
                if (line.IndexOf("button", StringComparison.OrdinalIgnoreCase) < 0)
                {
                    continue;
                }

                if (line.IndexOf(prefix, StringComparison.Ordinal) < 0)
                {
                    continue;
                }

                var q = System.Text.RegularExpressions.Regex.Match(line, "\"([^\"]+)\"");
                if (q.Success && q.Groups[1].Value.Length <= 16)
                {
                    return true;
                }
            }

            return false;
        }

        private static string[] RestAfter(string first)
        {
            int i = Array.IndexOf(AllIds, first);
            if (i < 0 || i + 1 >= AllIds.Length)
            {
                return new string[0];
            }

            var rest = new string[AllIds.Length - i - 1];
            Array.Copy(AllIds, i + 1, rest, 0, rest.Length);
            return rest;
        }
    }
}
