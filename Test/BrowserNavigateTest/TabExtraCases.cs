using System;
using WordAddIn1;

namespace BrowserNavigateTest
{
    internal static partial class TabCases
    {
        private static readonly string[] ExtraIds =
        {
            "tab-snap-has-dates", "tab-snap-no-tasklist", "tab-snap-no-count", "tab-snap-no-xuehao",
            "tab-snap-no-wang", "tab-snap-home", "tab-snap-process", "tab-snap-docs",
            "tab-snap-no-docs-body", "tab-snap-short16", "tab-snap-long17", "tab-snap-no-assign-done",
            "tab-click-task-direct", "tab-task-direct-list", "tab-task-direct-no-rows",
            "tab-task-direct-no-assign", "tab-task-direct-no-dates", "tab-task-direct-resnap",
            "tab-click-detail-again", "tab-detail-again-dates", "tab-click-docs", "tab-docs-body",
            "tab-docs-hides-dates", "tab-docs-hides-rows",
            "tab-member-wang", "tab-member-wang-ref", "tab-member-ids", "tab-member-teacher",
            "tab-member-no-dates", "tab-member-no-tasklist", "tab-member-count", "tab-member-no-docs-body",
            "tab-click-row-huang", "tab-click-row-liu", "tab-click-row-wang",
            "tab-round-task-member", "tab-round-detail-member", "tab-round-docs-member",
            "tab-assign-status", "tab-assign-keeps-rows",
            "tab-side-home", "tab-side-process", "tab-side-keeps-tabs", "tab-side-from-detail-no-rows",
            "tab-task-then-detail-dates", "tab-docs-resnap", "tab-short16-click",
            "tab-hidden-snap-no-rows", "tab-hidden-task-hides", "tab-hidden-roundtrip"
        };

        internal static void RunExtra(TestRun run, NavSession s, FixtureHttp http)
        {
            string id = IntCases.OpenOk(run, s, http.Url("center-tabs", "index.html"), "tab-snap-has-dates");
            if (id == null)
            {
                IntCases.FailMany(run, "依赖 center-tabs 巩固", ExtraAfter("tab-snap-has-dates"));
                return;
            }

            string snap = IntCases.WaitTree(s, id, false, "项目成员", "项目资料", "首页");
            AssertExtraSnap(run, snap);

            string taskRef = TabRef(snap, "项目任务");
            ToolResult task = Click(s, id, taskRef);
            string taskTree = TreeOf(task);
            IntCases.Ok(run, "tab-click-task-direct", OkStatus(task, "项目任务"), Why(task, "直点任务"));
            IntCases.Ok(run, "tab-task-direct-list", task != null && task.Success && taskTree.Contains("任务列表"), "无任务列表");
            IntCases.Ok(
                run,
                "tab-task-direct-no-rows",
                task != null && task.Success && !VisibleName(taskTree, "黄浩然") && !VisibleName(taskTree, "刘翔") && !VisibleName(taskTree, "王敏"),
                "直点任务仍见表行");
            IntCases.Ok(run, "tab-task-direct-no-assign", task != null && task.Success && !VisibleName(taskTree, "项目成员分配"), "直点任务已见分配");
            IntCases.Ok(run, "tab-task-direct-no-dates", task != null && task.Success && !VisibleName(taskTree, "立项时间"), "直点任务仍见立项时间");
            IntCases.Ok(
                run,
                "tab-task-direct-resnap",
                task != null && task.Success && IntCases.Flag(task, "resnapshot") && !IntCases.Flag(task, "refs_invalidated"),
                "直点任务未回树");

            string afterTask = task != null && task.Success ? taskTree : snap;
            string detailRef = TabRef(afterTask, "项目详情") ?? TabRef(IntCases.WaitTree(s, id, false, "项目详情"), "项目详情");
            ToolResult detail = Click(s, id, detailRef);
            string detailTree = TreeOf(detail);
            IntCases.Ok(run, "tab-click-detail-again", OkStatus(detail, "项目详情"), Why(detail, "再点详情"));
            IntCases.Ok(run, "tab-detail-again-dates", detail != null && detail.Success && VisibleName(detailTree, "立项时间"), "回详情无立项时间");

            string afterDetail = detail != null && detail.Success ? detailTree : afterTask;
            string docsRef = TabRef(afterDetail, "项目资料") ?? TabRef(IntCases.WaitTree(s, id, false, "项目资料"), "项目资料");
            ToolResult docs = Click(s, id, docsRef);
            string docsTree = TreeOf(docs);
            IntCases.Ok(run, "tab-click-docs", OkStatus(docs, "项目资料"), Why(docs, "点项目资料"));
            IntCases.Ok(run, "tab-docs-body", docs != null && docs.Success && docsTree.Contains("资料正文"), "无资料正文");
            IntCases.Ok(run, "tab-docs-hides-dates", docs != null && docs.Success && !VisibleName(docsTree, "立项时间"), "资料页仍见立项时间");
            IntCases.Ok(
                run,
                "tab-docs-hides-rows",
                docs != null && docs.Success && !VisibleName(docsTree, "黄浩然") && !VisibleName(docsTree, "王敏"),
                "资料页仍见表行");
            IntCases.Ok(
                run,
                "tab-docs-resnap",
                docs != null && docs.Success && IntCases.Flag(docs, "resnapshot") && !IntCases.Flag(docs, "refs_invalidated"),
                "点资料未回树");

            string afterDocs = docs != null && docs.Success ? docsTree : afterDetail;
            string mem = TabRef(afterDocs, "项目成员") ?? TabRef(IntCases.WaitTree(s, id, false, "项目成员"), "项目成员");
            ToolResult member = Click(s, id, mem);
            string memberTree = TreeOf(member);
            bool memberOk = member != null && member.Success && memberTree.Contains("项目成员");
            IntCases.Ok(run, "tab-member-wang", memberOk && VisibleName(memberTree, "王敏"), memberOk ? "无王敏" : Why(member, "点成员"));
            IntCases.Ok(run, "tab-member-wang-ref", memberOk && !string.IsNullOrEmpty(IntCases.FindRef(memberTree, "王敏")), "王敏无 ref");
            IntCases.Ok(
                run,
                "tab-member-ids",
                memberOk && memberTree.Contains("239990407") && memberTree.Contains("239990418") && memberTree.Contains("239990419"),
                "学号不齐");
            IntCases.Ok(run, "tab-member-teacher", memberOk && VisibleName(memberTree, "林嘉诚"), "无指导教师");
            IntCases.Ok(run, "tab-member-no-dates", memberOk && !VisibleName(memberTree, "立项时间"), "成员页仍见立项时间");
            IntCases.Ok(run, "tab-member-no-tasklist", memberOk && !VisibleName(memberTree, "任务列表"), "成员页仍见任务列表");
            IntCases.Ok(run, "tab-member-count", memberOk && memberTree.Contains("学生数=3"), "无学生数=3");
            IntCases.Ok(run, "tab-member-no-docs-body", memberOk && !VisibleName(memberTree, "资料正文"), "成员页仍见资料正文");

            if (!memberOk)
            {
                IntCases.FailMany(run, "点成员失败，后面巩固依赖成员树",
                    "tab-click-row-huang", "tab-click-row-liu", "tab-click-row-wang",
                    "tab-round-task-member", "tab-round-detail-member", "tab-round-docs-member",
                    "tab-assign-status", "tab-assign-keeps-rows");
            }
            else
            {
                ToolResult huang = Click(s, id, IntCases.FindRef(memberTree, "黄浩然"));
                IntCases.Ok(run, "tab-click-row-huang", huang != null && huang.Success, Why(huang, "点黄浩然行"));
                memberTree = huang != null && huang.Success ? TreeOf(huang) : memberTree;
                ToolResult liu = Click(s, id, IntCases.FindRef(memberTree, "刘翔"));
                IntCases.Ok(run, "tab-click-row-liu", liu != null && liu.Success, Why(liu, "点刘翔行"));
                memberTree = liu != null && liu.Success ? TreeOf(liu) : memberTree;
                ToolResult wang = Click(s, id, IntCases.FindRef(memberTree, "王敏"));
                IntCases.Ok(run, "tab-click-row-wang", wang != null && wang.Success, Why(wang, "点王敏行"));
                memberTree = wang != null && wang.Success ? TreeOf(wang) : memberTree;

                ToolResult toTask = Click(s, id, TabRef(memberTree, "项目任务"));
                string t2 = TreeOf(toTask);
                ToolResult backMem = Click(s, id, TabRef(t2, "项目成员") ?? TabRef(IntCases.WaitTree(s, id, false, "项目成员"), "项目成员"));
                string backMemTree = TreeOf(backMem);
                IntCases.Ok(
                    run,
                    "tab-round-task-member",
                    backMem != null && backMem.Success
                        && VisibleName(backMemTree, "黄浩然") && VisibleName(backMemTree, "刘翔") && VisibleName(backMemTree, "王敏"),
                    "任务再回成员两行不齐");

                ToolResult toDetail = Click(s, id, TabRef(backMemTree, "项目详情"));
                string d2 = TreeOf(toDetail);
                ToolResult backMem2 = Click(s, id, TabRef(d2, "项目成员") ?? TabRef(IntCases.WaitTree(s, id, false, "项目成员"), "项目成员"));
                string backMem2Tree = TreeOf(backMem2);
                IntCases.Ok(
                    run,
                    "tab-round-detail-member",
                    backMem2 != null && backMem2.Success && VisibleName(backMem2Tree, "黄浩然") && VisibleName(backMem2Tree, "王敏"),
                    "详情再回成员行丢了");

                ToolResult toDocs = Click(s, id, TabRef(backMem2Tree, "项目资料"));
                string docs2 = TreeOf(toDocs);
                ToolResult backMem3 = Click(s, id, TabRef(docs2, "项目成员") ?? TabRef(IntCases.WaitTree(s, id, false, "项目成员"), "项目成员"));
                string backMem3Tree = TreeOf(backMem3);
                IntCases.Ok(
                    run,
                    "tab-round-docs-member",
                    backMem3 != null && backMem3.Success && VisibleName(backMem3Tree, "刘翔") && VisibleName(backMem3Tree, "王敏"),
                    "资料再回成员行丢了");

                string assignRef = TabRef(backMem3Tree, "项目成员分配");
                ToolResult assign = Click(s, id, assignRef);
                string assignTree = assign != null && assign.Success ? TreeOf(assign) : backMem3Tree;
                IntCases.Ok(
                    run,
                    "tab-assign-keeps-rows",
                    assign != null && assign.Success && VisibleName(assignTree, "黄浩然") && VisibleName(assignTree, "王敏"),
                    "点分配后行丢了");
                string extraBtn = IntCases.FindRef(assignTree, "巩固点按")
                    ?? IntCases.FindRef(IntCases.WaitTree(s, id, false, "巩固点按"), "巩固点按");
                ToolResult extra = Click(s, id, extraBtn);
                IntCases.Ok(run, "tab-assign-status", extra != null && extra.Success && TreeOf(extra).Contains("已点分配"), Why(extra, "点巩固点按"));
            }

            string live = IntCases.WaitTree(s, id, false, "首页", "过程管理", "项目详情");
            ToolResult toDetail3 = Click(s, id, TabRef(live, "项目详情"));
            string detail3 = toDetail3 != null && toDetail3.Success ? TreeOf(toDetail3) : live;
            ToolResult home = Click(s, id, TabRef(detail3, "首页"));
            string homeTree = TreeOf(home);
            IntCases.Ok(run, "tab-side-home", home != null && home.Success && homeTree.Contains("侧栏") && homeTree.Contains("首页"), Why(home, "点首页"));
            string afterHome = home != null && home.Success ? homeTree : detail3;
            ToolResult proc = Click(s, id, TabRef(afterHome, "过程管理") ?? TabRef(IntCases.WaitTree(s, id, false, "过程管理"), "过程管理"));
            string procTree = TreeOf(proc);
            IntCases.Ok(run, "tab-side-process", proc != null && proc.Success && procTree.Contains("侧栏") && procTree.Contains("过程管理"), Why(proc, "点过程管理"));
            string afterProc = proc != null && proc.Success ? procTree : afterHome;
            IntCases.Ok(
                run,
                "tab-side-keeps-tabs",
                !string.IsNullOrEmpty(TabRef(afterProc, "项目详情"))
                    && !string.IsNullOrEmpty(TabRef(afterProc, "项目成员"))
                    && !string.IsNullOrEmpty(TabRef(afterProc, "项目任务"))
                    && !string.IsNullOrEmpty(TabRef(afterProc, "项目资料")),
                "点侧栏后 Tab 缺 ref");
            IntCases.Ok(
                run,
                "tab-side-from-detail-no-rows",
                !VisibleName(afterProc, "黄浩然") && !VisibleName(afterProc, "王敏"),
                "详情态点侧栏却见表行");

            ToolResult task2 = Click(s, id, TabRef(afterProc, "项目任务"));
            string task2Tree = TreeOf(task2);
            ToolResult detail4 = Click(s, id, TabRef(task2Tree, "项目详情") ?? TabRef(IntCases.WaitTree(s, id, false, "项目详情"), "项目详情"));
            IntCases.Ok(
                run,
                "tab-task-then-detail-dates",
                detail4 != null && detail4.Success && VisibleName(TreeOf(detail4), "立项时间") && VisibleName(TreeOf(detail4), "结项时间"),
                "任务再回详情无日期");

            string shortRef = IntCases.FindRef(TreeOf(detail4).Length > 0 ? TreeOf(detail4) : afterProc, "短标签巩固入口一二三四五六七八九")
                ?? IntCases.FindRef(IntCases.WaitTree(s, id, false, "短标签巩固入口"), "短标签巩固入口一二三四五六七八九");
            ToolResult shortClick = Click(s, id, shortRef);
            IntCases.Ok(
                run,
                "tab-short16-click",
                shortClick != null && shortClick.Success && TreeOf(shortClick).Contains("短标"),
                Why(shortClick, "点十六字短标签"));

            RunHiddenExtra(run, s, http);
        }

        private static void AssertExtraSnap(TestRun run, string snap)
        {
            IntCases.Ok(run, "tab-snap-has-dates", VisibleName(snap, "立项时间") && VisibleName(snap, "结项时间"), "默认详情无日期");
            IntCases.Ok(run, "tab-snap-no-tasklist", !VisibleName(snap, "任务列表"), "未点任务已见任务列表");
            IntCases.Ok(run, "tab-snap-no-count", !VisibleName(snap, "学生数=3"), "未点成员已见学生数");
            IntCases.Ok(run, "tab-snap-no-xuehao", !VisibleName(snap, "239990407"), "未点成员已见学号");
            IntCases.Ok(run, "tab-snap-no-wang", !VisibleName(snap, "王敏"), "未点成员已见王敏");
            IntCases.Ok(run, "tab-snap-home", !string.IsNullOrEmpty(TabRef(snap, "首页")), "无首页 ref");
            IntCases.Ok(run, "tab-snap-process", !string.IsNullOrEmpty(TabRef(snap, "过程管理")), "无过程管理 ref");
            IntCases.Ok(run, "tab-snap-docs", !string.IsNullOrEmpty(TabRef(snap, "项目资料")), "无项目资料 ref");
            IntCases.Ok(run, "tab-snap-no-docs-body", !VisibleName(snap, "资料正文"), "未点资料已见资料正文");
            IntCases.Ok(run, "tab-snap-short16", !string.IsNullOrEmpty(IntCases.FindRef(snap, "短标签巩固入口一二三四五六七八九")), "十六字短标签无 ref");
            IntCases.Ok(run, "tab-snap-long17", !LongTextIsShortButton(snap, "短标签巩固入口一二三四五六七八九十"), "十七字收成短按钮");
            IntCases.Ok(run, "tab-snap-no-assign-done", !VisibleName(snap, "已点分配"), "未点分配已见已点分配");
        }

        private static void RunHiddenExtra(TestRun run, NavSession s, FixtureHttp http)
        {
            string hid = IntCases.OpenOk(run, s, http.Url("center-tabs", "index.html"), "tab-hidden-snap-no-rows", visible: false);
            if (hid == null)
            {
                IntCases.FailMany(run, "依赖隐藏窗", "tab-hidden-task-hides", "tab-hidden-roundtrip");
                return;
            }

            string snap = IntCases.WaitTree(s, hid, false, "项目成员");
            IntCases.Ok(
                run,
                "tab-hidden-snap-no-rows",
                !VisibleName(snap, "黄浩然") && !VisibleName(snap, "王敏"),
                "隐藏窗空拍已见表行");
            ToolResult mem = Click(s, hid, TabRef(snap, "项目成员"));
            string memTree = TreeOf(mem);
            ToolResult task = Click(s, hid, TabRef(memTree, "项目任务"));
            string taskTree = TreeOf(task);
            IntCases.Ok(
                run,
                "tab-hidden-task-hides",
                task != null && task.Success && !VisibleName(taskTree, "黄浩然") && !VisibleName(taskTree, "王敏"),
                "隐藏窗切任务后行仍在");
            ToolResult back = Click(s, hid, TabRef(taskTree, "项目成员") ?? TabRef(IntCases.WaitTree(s, hid, false, "项目成员"), "项目成员"));
            IntCases.Ok(
                run,
                "tab-hidden-roundtrip",
                back != null && back.Success && VisibleName(TreeOf(back), "黄浩然") && VisibleName(TreeOf(back), "王敏"),
                "隐藏窗再回成员行不齐");
        }

        private static ToolResult Click(NavSession s, string id, string refId)
        {
            if (string.IsNullOrEmpty(id) || string.IsNullOrEmpty(refId))
            {
                return null;
            }

            return s.Interact(id, "click", refId);
        }

        private static string TreeOf(ToolResult r)
        {
            return r == null ? "" : IntCases.Tree(r);
        }

        private static bool OkStatus(ToolResult r, string text)
        {
            return r != null && r.Success && TreeOf(r).Contains(text);
        }

        private static string Why(ToolResult r, string what)
        {
            if (r == null)
            {
                return "无 " + what + " ref";
            }

            return r.Success ? what + " 状态未变" : r.Error;
        }

        private static bool LongTextIsShortButton(string tree, string text)
        {
            if (string.IsNullOrEmpty(tree) || string.IsNullOrEmpty(text))
            {
                return false;
            }

            foreach (string line in tree.Split('\n'))
            {
                if (line.IndexOf("button", StringComparison.OrdinalIgnoreCase) < 0)
                {
                    continue;
                }

                if (line.IndexOf(text, StringComparison.Ordinal) < 0)
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

        private static string[] ExtraAfter(string first)
        {
            int i = Array.IndexOf(ExtraIds, first);
            if (i < 0 || i + 1 >= ExtraIds.Length)
            {
                return new string[0];
            }

            var rest = new string[ExtraIds.Length - i - 1];
            Array.Copy(ExtraIds, i + 1, rest, 0, rest.Length);
            return rest;
        }
    }
}
