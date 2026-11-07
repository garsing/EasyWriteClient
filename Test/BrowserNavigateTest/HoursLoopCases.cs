using System;
using System.Threading;
using WordAddIn1;

namespace BrowserNavigateTest
{
    /// <summary>工时填报绕圈短套。只跑易写浏览 agent，不重跑已绿 snp / int。</summary>
    internal static class HoursLoopCases
    {
        private static readonly string[] All =
        {
            "hrs-number-commit",
            "hrs-text-plain",
            "hrs-detail-keep",
            "hrs-xpath-keep",
            "hrs-nav-clear",
            "hrs-sticky-scroll",
            "hrs-sticky-stuck",
            "hrs-shade-close",
            "hrs-detail-miss",
        };

        public static void Run(TestRun run, FixtureHttp http, string host, string filter)
        {
            IntCases.NavHost = host;
            if (!string.Equals(host, "yiwrite", StringComparison.OrdinalIgnoreCase))
            {
                foreach (string id in All)
                {
                    if (Want(filter, id))
                    {
                        run.CaseFail(id, "本套只覆盖易写浏览 agent");
                    }
                }

                return;
            }

            var s = new NavSession();
            if (Want(filter, "hrs-number-commit"))
            {
                NumberCommit(run, s, http);
            }

            if (Want(filter, "hrs-text-plain"))
            {
                TextPlain(run, s, http);
            }

            if (Want(filter, "hrs-detail-keep"))
            {
                DetailKeep(run, s, http);
            }

            if (Want(filter, "hrs-xpath-keep"))
            {
                XPathKeep(run, s, http);
            }

            if (Want(filter, "hrs-nav-clear"))
            {
                NavClear(run, s, http);
            }

            if (Want(filter, "hrs-sticky-scroll"))
            {
                StickyScroll(run, s, http);
            }

            if (Want(filter, "hrs-sticky-stuck"))
            {
                StickyStuck(run, s, http);
            }

            if (Want(filter, "hrs-shade-close"))
            {
                ShadeClose(run, s, http);
            }

            if (Want(filter, "hrs-detail-miss"))
            {
                DetailMiss(run, s, http);
            }
        }

        private static void NumberCommit(TestRun run, NavSession s, FixtureHttp http)
        {
            const string idCase = "hrs-number-commit";
            string id = Open(run, s, http, "form.html", idCase);
            if (id == null)
            {
                return;
            }

            string tree = IntCases.WaitTree(s, id, true, "工时");
            string hours = IntCases.FindRef(tree, "工时", "spinbutton") ?? IntCases.FindRef(tree, "工时");
            if (string.IsNullOrEmpty(hours))
            {
                run.CaseFail(idCase, "无工时 ref");
                return;
            }

            ToolResult typed = s.Interact(id, "type", hours, text: "8");
            Thread.Sleep(250);
            string after = IntCases.Tree(s.Snap(id, domSupplement: "input"));
            if (string.IsNullOrEmpty(after))
            {
                after = IntCases.Tree(typed);
            }

            IntCases.Ok(
                run,
                idCase,
                typed != null && typed.Success
                    && after.IndexOf("工时已写入 8", StringComparison.Ordinal) >= 0
                    && after.IndexOf("工时回退到0", StringComparison.Ordinal) < 0,
                typed == null ? "无回包" : (typed.Success ? after : typed.Error));
        }

        private static void TextPlain(TestRun run, NavSession s, FixtureHttp http)
        {
            const string idCase = "hrs-text-plain";
            string id = Open(run, s, http, "form.html", idCase);
            if (id == null)
            {
                return;
            }

            string tree = IntCases.WaitTree(s, id, true, "备注");
            string note = IntCases.FindRef(tree, "备注", "textbox") ?? IntCases.FindRef(tree, "备注");
            if (string.IsNullOrEmpty(note))
            {
                run.CaseFail(idCase, "无备注 ref");
                return;
            }

            ToolResult typed = s.Interact(id, "type", note, text: "甲");
            Thread.Sleep(200);
            string after = IntCases.Tree(s.Snap(id, domSupplement: "input"));
            IntCases.Ok(
                run,
                idCase,
                typed != null && typed.Success && after.IndexOf("备注已写入 甲", StringComparison.Ordinal) >= 0,
                typed == null ? "无回包" : (typed.Success ? after : typed.Error));
        }

        private static void DetailKeep(TestRun run, NavSession s, FixtureHttp http)
        {
            const string idCase = "hrs-detail-keep";
            string id = Open(run, s, http, "form.html", idCase);
            if (id == null)
            {
                return;
            }

            string tree = WaitAll(s, id, "工时", "工作内容", "提交工时");
            string hours = IntCases.FindRef(tree, "工时", "spinbutton") ?? IntCases.FindRef(tree, "工时");
            string content = IntCases.FindRef(tree, "工作内容");
            string submit = IntCases.FindRef(tree, "提交工时");
            if (string.IsNullOrEmpty(hours) || string.IsNullOrEmpty(content) || string.IsNullOrEmpty(submit))
            {
                run.CaseFail(idCase, "缺号 工时=" + hours + " 内容=" + content + " 提交=" + submit);
                return;
            }

            ToolResult detail = s.Snap(id, hours);
            string detailTree = IntCases.Tree(detail);
            bool treeOk = detail != null && detail.Success
                && detailTree.IndexOf("工时", StringComparison.Ordinal) >= 0
                && detailTree.IndexOf("工作内容", StringComparison.Ordinal) < 0
                && detailTree.IndexOf("提交工时", StringComparison.Ordinal) < 0;
            if (!treeOk)
            {
                run.CaseFail(idCase, detail != null && detail.Success ? detailTree : (detail == null ? "无回包" : detail.Error));
                return;
            }

            ToolResult typed = s.Interact(id, "type", content, text: "填报说明");
            Thread.Sleep(200);
            string afterType = IntCases.Tree(s.Snap(id));
            if (typed == null || !typed.Success || afterType.IndexOf("内容已写入", StringComparison.Ordinal) < 0)
            {
                run.CaseFail(idCase, typed == null ? "内容无回包" : (typed.Success ? afterType : typed.Error));
                return;
            }

            ToolResult clicked = s.Interact(id, "click", submit);
            Thread.Sleep(200);
            string afterClick = IntCases.Tree(s.Snap(id));
            IntCases.Ok(
                run,
                idCase,
                clicked != null && clicked.Success && afterClick.IndexOf("已点提交", StringComparison.Ordinal) >= 0,
                clicked == null ? "提交无回包" : (clicked.Success ? afterClick : clicked.Error));
        }

        private static void XPathKeep(TestRun run, NavSession s, FixtureHttp http)
        {
            const string idCase = "hrs-xpath-keep";
            string id = Open(run, s, http, "form.html", idCase);
            if (id == null)
            {
                return;
            }

            string tree = WaitAll(s, id, "工时", "天数", "换节点");
            string hours = IntCases.FindRef(tree, "工时", "spinbutton") ?? IntCases.FindRef(tree, "工时");
            string days = IntCases.FindRef(tree, "天数", "spinbutton") ?? IntCases.FindRef(tree, "天数");
            string swap = IntCases.FindRef(tree, "换节点");
            if (string.IsNullOrEmpty(hours) || string.IsNullOrEmpty(days) || string.IsNullOrEmpty(swap))
            {
                run.CaseFail(idCase, "缺号 工时=" + hours + " 天数=" + days + " 换=" + swap);
                return;
            }

            ToolResult clicked = s.Interact(id, "click", swap);
            if (clicked == null || !clicked.Success)
            {
                run.CaseFail(idCase, clicked == null ? "换节点无回包" : clicked.Error);
                return;
            }

            string again = WaitAll(s, id, "已换节点", "工时", "天数");
            string hours2 = IntCases.FindRef(again, "工时", "spinbutton") ?? IntCases.FindRef(again, "工时");
            string days2 = IntCases.FindRef(again, "天数", "spinbutton") ?? IntCases.FindRef(again, "天数");
            if (!string.Equals(hours, hours2, StringComparison.OrdinalIgnoreCase)
                || !string.Equals(days, days2, StringComparison.OrdinalIgnoreCase))
            {
                run.CaseFail(idCase, "号变了 工时 " + hours + "→" + hours2 + " 天数 " + days + "→" + days2);
                return;
            }

            ToolResult typed = s.Interact(id, "type", hours, text: "9");
            Thread.Sleep(250);
            string after = IntCases.Tree(s.Snap(id, domSupplement: "input"));
            IntCases.Ok(
                run,
                idCase,
                typed != null && typed.Success
                    && after.IndexOf("工时已写入 9", StringComparison.Ordinal) >= 0
                    && after.IndexOf("天数已写入", StringComparison.Ordinal) < 0,
                typed == null ? "无回包" : (typed.Success ? after : typed.Error));
        }

        private static void NavClear(TestRun run, NavSession s, FixtureHttp http)
        {
            const string idCase = "hrs-nav-clear";
            string id = Open(run, s, http, "form.html", idCase);
            if (id == null)
            {
                return;
            }

            string tree = IntCases.WaitTree(s, id, true, "工时");
            string hours = IntCases.FindRef(tree, "工时", "spinbutton") ?? IntCases.FindRef(tree, "工时");
            if (string.IsNullOrEmpty(hours))
            {
                run.CaseFail(idCase, "无工时 ref");
                return;
            }

            ToolResult jump = s.Nav(http.Url("hours-loop", "other.html"), host: IntCases.NavHost, channelId: id);
            if (jump == null || !jump.Success)
            {
                run.CaseFail(idCase, jump == null ? "跳转无回包" : jump.Error);
                return;
            }

            string landed = IntCases.WaitTree(s, id, false, "另一页标记");
            if (landed.IndexOf("另一页标记", StringComparison.Ordinal) < 0)
            {
                run.CaseFail(idCase, landed);
                return;
            }

            ToolResult typed = s.Interact(id, "type", hours, text: "8");
            string err = typed == null ? "" : (typed.Error ?? "");
            IntCases.Ok(
                run,
                idCase,
                typed != null && !typed.Success
                    && (err.IndexOf("过期", StringComparison.Ordinal) >= 0
                        || err.IndexOf("无效", StringComparison.Ordinal) >= 0),
                typed == null ? "无回包" : (typed.Success ? "旧号仍能输入" : err));
        }

        private static void StickyScroll(TestRun run, NavSession s, FixtureHttp http)
        {
            const string idCase = "hrs-sticky-scroll";
            string id = Open(run, s, http, "sticky-scroll.html", idCase);
            if (id == null)
            {
                return;
            }

            string tree = IntCases.WaitTree(s, id, false, "查询");
            string query = IntCases.FindRef(tree, "查询", "button") ?? IntCases.FindRef(tree, "查询");
            if (string.IsNullOrEmpty(query))
            {
                run.CaseFail(idCase, "无查询 ref");
                return;
            }

            ToolResult clicked = s.Interact(id, "click", query);
            string err = clicked == null ? "" : (clicked.Error ?? "");
            if (clicked == null || !clicked.Success
                || err.IndexOf("蒙层", StringComparison.Ordinal) >= 0
                || err.IndexOf("关闭", StringComparison.Ordinal) >= 0)
            {
                run.CaseFail(idCase, clicked == null ? "无回包" : (clicked.Success ? err : err));
                return;
            }

            Thread.Sleep(200);
            string after = IntCases.Tree(s.Snap(id));
            IntCases.Ok(
                run,
                idCase,
                after.IndexOf("已查询完成", StringComparison.Ordinal) >= 0,
                after);
        }

        private static void StickyStuck(TestRun run, NavSession s, FixtureHttp http)
        {
            const string idCase = "hrs-sticky-stuck";
            string id = Open(run, s, http, "sticky-stuck.html", idCase);
            if (id == null)
            {
                return;
            }

            string tree = IntCases.WaitTree(s, id, false, "查询");
            string query = IntCases.FindRef(tree, "查询", "button") ?? IntCases.FindRef(tree, "查询");
            ToolResult clicked = string.IsNullOrEmpty(query) ? null : s.Interact(id, "click", query);
            string err = clicked == null ? "" : (clicked.Error ?? "");
            IntCases.Ok(
                run,
                idCase,
                clicked != null && !clicked.Success
                    && err.IndexOf("固定条", StringComparison.Ordinal) >= 0
                    && err.IndexOf("蒙层", StringComparison.Ordinal) < 0
                    && err.IndexOf("关闭", StringComparison.Ordinal) < 0,
                string.IsNullOrEmpty(query) ? "无查询 ref" : (clicked != null && clicked.Success ? "应失败却成功" : err));
        }

        private static void ShadeClose(TestRun run, NavSession s, FixtureHttp http)
        {
            const string idCase = "hrs-shade-close";
            string id = Open(run, s, http, "shade.html", idCase);
            if (id == null)
            {
                return;
            }

            string tree = IntCases.WaitTree(s, id, false, "底层查询");
            string under = IntCases.FindRef(tree, "底层查询");
            ToolResult clicked = string.IsNullOrEmpty(under) ? null : s.Interact(id, "click", under);
            string err = clicked == null ? "" : (clicked.Error ?? "");
            IntCases.Ok(
                run,
                idCase,
                clicked != null && !clicked.Success
                    && err.IndexOf("蒙层", StringComparison.Ordinal) >= 0
                    && err.IndexOf("固定条", StringComparison.Ordinal) < 0,
                string.IsNullOrEmpty(under) ? "无底层查询 ref" : (clicked != null && clicked.Success ? "应失败却成功" : err));
        }

        private static void DetailMiss(TestRun run, NavSession s, FixtureHttp http)
        {
            const string idCase = "hrs-detail-miss";
            string id = Open(run, s, http, "form.html", idCase);
            if (id == null)
            {
                return;
            }

            string tree = IntCases.WaitTree(s, id, false, "提交工时");
            string submit = IntCases.FindRef(tree, "提交工时");
            if (string.IsNullOrEmpty(submit))
            {
                run.CaseFail(idCase, "无提交 ref");
                return;
            }

            ToolResult missed = s.Snap(id, "e99999");
            string back = IntCases.Tree(missed);
            string message = IntCases.Str(missed, "message");
            bool kept = missed != null && missed.Success
                && message.IndexOf("这次没有展开", StringComparison.Ordinal) >= 0
                && back.IndexOf("提交工时", StringComparison.Ordinal) >= 0
                && back.IndexOf("请重新", StringComparison.Ordinal) < 0;
            if (!kept)
            {
                run.CaseFail(
                    idCase,
                    missed == null
                        ? "无回包"
                        : (missed.Success ? (message + " / " + back) : missed.Error));
                return;
            }

            ToolResult clicked = s.Interact(id, "click", submit);
            Thread.Sleep(200);
            string after = IntCases.Tree(s.Snap(id));
            IntCases.Ok(
                run,
                idCase,
                clicked != null && clicked.Success && after.IndexOf("已点提交", StringComparison.Ordinal) >= 0,
                clicked == null ? "提交无回包" : (clicked.Success ? after : clicked.Error));
        }

        private static string WaitAll(NavSession s, string id, params string[] hints)
        {
            string last = "";
            for (int i = 0; i < 8; i++)
            {
                last = IntCases.Tree(s.Snap(id, domSupplement: "input"));
                if (HasAll(last, hints))
                {
                    return last;
                }

                Thread.Sleep(150);
            }

            return last;
        }

        private static bool HasAll(string tree, string[] hints)
        {
            if (string.IsNullOrEmpty(tree) || hints == null)
            {
                return false;
            }

            for (int i = 0; i < hints.Length; i++)
            {
                if (tree.IndexOf(hints[i], StringComparison.Ordinal) < 0)
                {
                    return false;
                }
            }

            return true;
        }

        private static string Open(TestRun run, NavSession s, FixtureHttp http, string file, string failId)
        {
            return IntCases.OpenOk(run, s, http.Url("hours-loop", file), failId);
        }

        private static bool Want(string filter, string id)
        {
            if (string.IsNullOrEmpty(filter)
                || string.Equals(filter, "hrs", StringComparison.OrdinalIgnoreCase)
                || string.Equals(filter, "hrs-", StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            return string.Equals(filter, id, StringComparison.OrdinalIgnoreCase)
                || id.StartsWith(filter, StringComparison.OrdinalIgnoreCase);
        }
    }
}
