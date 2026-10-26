using System;
using System.Text.RegularExpressions;
using WordAddIn1;

namespace BrowserNavigateTest
{
    /// <summary>重复结构收起的真浏览器操作条。经 Desktop 管道调已开的易写。</summary>
    internal static class RepeatLiveCases
    {
        public static void Run(TestRun run, FixtureHttp http, string navHost, string filter)
        {
            if (string.IsNullOrEmpty(navHost) || navHost == "agent")
            {
                navHost = "yiwrite";
            }

            var s = new NavSession();
            string small = http.Url("repeat-rows", "index.html");
            string big = small + "?n=400";
            bool all = string.Equals(filter, "rpt-", StringComparison.OrdinalIgnoreCase)
                || string.Equals(filter, "rpt", StringComparison.OrdinalIgnoreCase);

            string id = Open(run, s, small, navHost, "rpt-op-overview");
            if (id == null)
            {
                FailRest(run, "依赖开页", all, filter,
                    "rpt-op-row-detail", "rpt-op-name40", "rpt-op-seq",
                    "rpt-op-repeat-no-range", "rpt-op-one-bound", "rpt-op-range-no-ref",
                    "rpt-op-range-on-link", "rpt-op-range-bad", "rpt-op-click-repeat",
                    "rpt-op-click-join", "rpt-op-resnap", "rpt-op-legacy");
            }
            else
            {
                ToolResult ov = s.Snap(id);
                string tree = Tree(ov);
                string joinRef = RefOn(tree, "加入作业");
                bool show = ov.Success
                    && Str(ov, "mode") == "overview"
                    && tree.IndexOf("repeat", StringComparison.Ordinal) >= 0
                    && tree.IndexOf("共 ", StringComparison.Ordinal) >= 0
                    && tree.IndexOf("当前 ", StringComparison.Ordinal) >= 0
                    && tree.IndexOf("INNER", StringComparison.Ordinal) >= 0
                    && !string.IsNullOrEmpty(joinRef)
                    && tree.IndexOf("repeat", StringComparison.Ordinal) < tree.IndexOf("加入作业", StringComparison.Ordinal);
                Ok(run, "rpt-op-overview", Want(all, filter, "rpt-op-overview"), show, show ? "" : Head(tree, ov));

                if (Want(all, filter, "rpt-op-row-detail"))
                {
                    string rowTree = tree;
                    string rowRef = RefOn(rowTree, "Q003");
                    if (string.IsNullOrEmpty(rowRef))
                    {
                        ToolResult again = s.Snap(id);
                        rowTree = Tree(again);
                        rowRef = RefOn(rowTree, "Q003");
                    }

                    if (string.IsNullOrEmpty(rowRef))
                    {
                        run.CaseFail("rpt-op-row-detail", "收起行没有 ref\n" + Head(rowTree, ov));
                    }
                    else
                    {
                        ToolResult dt = s.Snap(id, rowRef);
                        string inner = Tree(dt);
                        Ok(run, "rpt-op-row-detail", true,
                            dt.Success
                            && Str(dt, "mode") == "detail"
                            && Str(dt, "ref") == rowRef
                            && (inner.IndexOf("textbox", StringComparison.OrdinalIgnoreCase) >= 0
                                || inner.IndexOf("checkbox", StringComparison.OrdinalIgnoreCase) >= 0
                                || inner.IndexOf("answer", StringComparison.Ordinal) >= 0),
                            dt.Success ? Head(inner, dt) : dt.Error);
                    }
                }

                if (Want(all, filter, "rpt-op-name40"))
                {
                    Ok(run, "rpt-op-name40", true,
                        tree.IndexOf("LONG", StringComparison.Ordinal) >= 0
                        && tree.IndexOf("…", StringComparison.Ordinal) >= 0
                        && tree.IndexOf("尾部标记TAIL", StringComparison.Ordinal) < 0,
                        Head(tree, ov));
                }

                if (Want(all, filter, "rpt-op-seq"))
                {
                    Ok(run, "rpt-op-seq", true,
                        tree.IndexOf("第9题", StringComparison.Ordinal) >= 0
                        && Regex.IsMatch(tree, @"当前 1-\d+"),
                        Head(tree, ov));
                }

                if (Want(all, filter, "rpt-op-repeat-no-range"))
                {
                    ToolResult fresh = s.Snap(id);
                    string freshRepeat = RefOfRole(Tree(fresh), "repeat");
                    if (string.IsNullOrEmpty(freshRepeat))
                    {
                        run.CaseFail("rpt-op-repeat-no-range", "没有 repeat ref\n" + Head(Tree(fresh), fresh));
                    }
                    else
                    {
                        ExpectFail(run, "rpt-op-repeat-no-range", s.Snap(id, freshRepeat), "重复段请同时传入");
                    }
                }

                if (Want(all, filter, "rpt-op-one-bound"))
                {
                    ExpectFail(run, "rpt-op-one-bound", s.Snap(id, rangeStart: 1), "要两个一起传");
                }

                if (Want(all, filter, "rpt-op-range-no-ref"))
                {
                    ExpectFail(run, "rpt-op-range-no-ref", s.Snap(id, rangeStart: 1, rangeEnd: 3), "范围只对重复段");
                }

                if (Want(all, filter, "rpt-op-range-on-link"))
                {
                    ToolResult fresh = s.Snap(id);
                    string freshJoin = RefOn(Tree(fresh), "加入作业");
                    if (string.IsNullOrEmpty(freshJoin))
                    {
                        run.CaseFail("rpt-op-range-on-link", "加入作业没有 ref");
                    }
                    else
                    {
                        ExpectFail(run, "rpt-op-range-on-link", s.Snap(id, freshJoin, rangeStart: 1, rangeEnd: 3), "范围只对重复段");
                    }
                }

                if (Want(all, filter, "rpt-op-range-bad"))
                {
                    ToolResult fresh = s.Snap(id);
                    string freshRepeat = RefOfRole(Tree(fresh), "repeat");
                    ToolResult bad = string.IsNullOrEmpty(freshRepeat)
                        ? s.Snap(id, "e1", rangeStart: 0, rangeEnd: 1)
                        : s.Snap(id, freshRepeat, rangeStart: 0, rangeEnd: 1);
                    ExpectFail(run, "rpt-op-range-bad", bad, "序号范围无效");
                }

                if (Want(all, filter, "rpt-op-click-repeat"))
                {
                    ToolResult fresh = s.Snap(id);
                    string freshRepeat = RefOfRole(Tree(fresh), "repeat");
                    if (string.IsNullOrEmpty(freshRepeat))
                    {
                        run.CaseFail("rpt-op-click-repeat", "没有 repeat ref");
                    }
                    else
                    {
                        ToolResult clicked = s.Interact(id, "click", freshRepeat);
                        string clickErr = clicked.Error ?? "";
                        bool refused = !clicked.Success
                            && (clickErr.IndexOf("可点区域", StringComparison.Ordinal) >= 0
                                || clickErr.IndexOf("缺少可操作节点定位", StringComparison.Ordinal) >= 0);
                        if (refused)
                        {
                            run.CaseOk("rpt-op-click-repeat");
                        }
                        else
                        {
                            run.CaseFail("rpt-op-click-repeat", clicked.Success ? "应失败却成功" : clickErr);
                        }
                    }
                }

                if (Want(all, filter, "rpt-op-click-join") || Want(all, filter, "rpt-op-resnap"))
                {
                    ToolResult fresh = s.Snap(id);
                    string freshJoin = RefOn(Tree(fresh), "加入作业");
                    if (string.IsNullOrEmpty(freshJoin))
                    {
                        if (Want(all, filter, "rpt-op-click-join"))
                        {
                            run.CaseFail("rpt-op-click-join", "加入作业没有 ref");
                        }

                        if (Want(all, filter, "rpt-op-resnap"))
                        {
                            run.CaseFail("rpt-op-resnap", "加入作业没有 ref");
                        }
                    }
                    else
                    {
                        ToolResult hit = s.Interact(id, "click", freshJoin);
                        string back = Tree(hit);
                        if (Want(all, filter, "rpt-op-click-join"))
                        {
                            Ok(run, "rpt-op-click-join", true,
                                hit.Success && back.IndexOf("已加入", StringComparison.Ordinal) >= 0,
                                hit.Success ? Head(back, hit) : hit.Error);
                        }

                        if (Want(all, filter, "rpt-op-resnap"))
                        {
                            string againJoin = RefOn(back, "加入作业");
                            Ok(run, "rpt-op-resnap", true,
                                hit.Success
                                && back.IndexOf("repeat", StringComparison.Ordinal) >= 0
                                && !string.IsNullOrEmpty(againJoin)
                                && back.IndexOf("repeat", StringComparison.Ordinal) < back.IndexOf("加入作业", StringComparison.Ordinal),
                                hit.Success ? Head(back, hit) : hit.Error);
                        }
                    }
                }

                if (Want(all, filter, "rpt-op-legacy"))
                {
                    run.CaseSkip("rpt-op-legacy", "本机开的是易写指针轨，这条只对非指针轨");
                }
            }

            if (!(Want(all, filter, "rpt-op-range-page")
                || Want(all, filter, "rpt-op-range-cap")
                || Want(all, filter, "rpt-op-range-tail")
                || Want(all, filter, "rpt-op-stale")))
            {
                return;
            }

            string bigId = Open(run, s, big, navHost, "rpt-op-range-page");
            if (bigId == null)
            {
                FailRest(run, "400 题页没打开", all, filter,
                    "rpt-op-range-page", "rpt-op-range-cap", "rpt-op-range-tail", "rpt-op-stale");
                return;
            }

            ToolResult bigOv = s.Snap(bigId);
            string bigTree = Tree(bigOv);
            string seg = RefOfRole(bigTree, "repeat");
            if (string.IsNullOrEmpty(seg))
            {
                FailRest(run, "400 题页没有 repeat\n" + Head(bigTree, bigOv), all, filter,
                    "rpt-op-range-page", "rpt-op-range-cap", "rpt-op-range-tail", "rpt-op-stale");
                return;
            }

            if (Want(all, filter, "rpt-op-range-page"))
            {
                ToolResult page = s.Snap(bigId, seg, rangeStart: 301, rangeEnd: 400);
                string pageTree = Tree(page);
                Ok(run, "rpt-op-range-page", true,
                    page.Success
                    && Str(page, "mode") == "overview"
                    && !string.IsNullOrEmpty(Str(page, "ref"))
                    && pageTree.IndexOf("当前 301-400", StringComparison.Ordinal) >= 0
                    && pageTree.IndexOf("Q301", StringComparison.Ordinal) >= 0
                    && pageTree.IndexOf("Q400", StringComparison.Ordinal) >= 0
                    && pageTree.IndexOf("Q300", StringComparison.Ordinal) < 0
                    && RefOn(pageTree, "加入作业").Length > 0,
                    page.Success ? Head(pageTree, page) : page.Error);
            }

            if (Want(all, filter, "rpt-op-range-cap"))
            {
                ToolResult cap = s.Snap(bigId, seg, rangeStart: 1, rangeEnd: 500);
                string capTree = Tree(cap);
                Ok(run, "rpt-op-range-cap", true,
                    cap.Success && capTree.IndexOf("当前 1-300", StringComparison.Ordinal) >= 0,
                    cap.Success ? Head(capTree, cap) : cap.Error);
                seg = RefOfRole(capTree, "repeat");
                if (string.IsNullOrEmpty(seg))
                {
                    seg = RefOfRole(Tree(s.Snap(bigId)), "repeat");
                }
            }

            if (Want(all, filter, "rpt-op-range-tail"))
            {
                if (string.IsNullOrEmpty(seg))
                {
                    run.CaseFail("rpt-op-range-tail", "翻段后没有 repeat ref");
                }
                else
                {
                    ToolResult tail = s.Snap(bigId, seg, rangeStart: 350, rangeEnd: 1000);
                    string tailTree = Tree(tail);
                    Ok(run, "rpt-op-range-tail", true,
                        tail.Success && tailTree.IndexOf("当前 350-400", StringComparison.Ordinal) >= 0,
                        tail.Success ? Head(tailTree, tail) : tail.Error);
                    seg = RefOfRole(tailTree, "repeat");
                }
            }

            if (Want(all, filter, "rpt-op-stale"))
            {
                ToolResult before = s.Snap(bigId);
                string oldRow = RefOn(Tree(before), "Q010");
                if (string.IsNullOrEmpty(oldRow))
                {
                    run.CaseFail("rpt-op-stale", "空拍没有 Q010 的 ref\n" + Head(Tree(before), before));
                }
                else
                {
                    ToolResult moved = s.Snap(bigId, RefOfRole(Tree(before), "repeat"), rangeStart: 301, rangeEnd: 400);
                    if (!moved.Success)
                    {
                        run.CaseFail("rpt-op-stale", moved.Error ?? "翻段失败");
                    }
                    else
                    {
                        ToolResult again = s.Snap(bigId, oldRow);
                        if (!again.Success)
                        {
                            ExpectFail(run, "rpt-op-stale", again, "过期");
                        }
                        else if (Tree(again).IndexOf("Q010", StringComparison.Ordinal) >= 0)
                        {
                            run.CaseFail("rpt-op-stale", "翻段后旧 ref 仍打开 Q010\n" + Head(Tree(again), again));
                        }
                        else
                        {
                            run.CaseOk("rpt-op-stale");
                        }
                    }
                }
            }
        }

        private static bool Want(bool all, string filter, string id)
        {
            return all || filter.IndexOf(id, StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private static string Open(TestRun run, NavSession s, string url, string host, string failId)
        {
            ToolResult r = s.Nav(url, host: host);
            if (r.Success && NavSession.IsShortId(Str(r, "channel_id")))
            {
                return Str(r, "channel_id");
            }

            run.CaseFail(failId, r.Success ? "无短 channel_id" : (r.Error ?? "打开失败"));
            return null;
        }

        private static void FailRest(TestRun run, string reason, bool all, string filter, params string[] ids)
        {
            for (int i = 0; i < ids.Length; i++)
            {
                if (Want(all, filter, ids[i]))
                {
                    run.CaseFail(ids[i], reason);
                }
            }
        }

        private static void Ok(TestRun run, string id, bool wanted, bool pass, string reason)
        {
            if (!wanted)
            {
                return;
            }

            if (pass)
            {
                run.CaseOk(id);
            }
            else
            {
                run.CaseFail(id, string.IsNullOrEmpty(reason) ? "断言失败" : reason);
            }
        }

        private static void ExpectFail(TestRun run, string id, ToolResult r, string mustContain)
        {
            if (r != null && !r.Success)
            {
                string err = r.Error ?? "";
                if (err.IndexOf(mustContain, StringComparison.Ordinal) >= 0)
                {
                    run.CaseOk(id);
                    return;
                }

                run.CaseFail(id, "失败文案不含「" + mustContain + "」: " + err);
                return;
            }

            run.CaseFail(id, r != null && r.Success ? "应失败却成功" : "无回包");
        }

        private static string Str(ToolResult r, string key)
        {
            var d = r?.Data as System.Collections.Generic.Dictionary<string, object>;
            if (d == null || !d.ContainsKey(key) || d[key] == null)
            {
                return "";
            }

            return Convert.ToString(d[key]) ?? "";
        }

        private static string Tree(ToolResult r)
        {
            return Str(r, "snapshot");
        }

        private static string RefOn(string tree, string token)
        {
            foreach (string line in (tree ?? "").Split('\n'))
            {
                if (line.IndexOf(token, StringComparison.Ordinal) < 0)
                {
                    continue;
                }

                Match m = Regex.Match(line, @"ref=(e\d+)");
                if (m.Success)
                {
                    return m.Groups[1].Value;
                }
            }

            return "";
        }

        private static string RefOfRole(string tree, string role)
        {
            foreach (string line in (tree ?? "").Split('\n'))
            {
                if (!Regex.IsMatch(line, @"(^|\s)-\s+" + Regex.Escape(role) + @"\b"))
                {
                    continue;
                }

                Match m = Regex.Match(line, @"ref=(e\d+)");
                if (m.Success)
                {
                    return m.Groups[1].Value;
                }
            }

            return "";
        }

        private static string Head(string tree, ToolResult r)
        {
            string err = r != null && !r.Success ? (r.Error ?? "") : "";
            string text = tree ?? "";
            if (text.Length > 900)
            {
                text = text.Substring(0, 900);
            }

            return err + "\n" + text;
        }
    }
}
