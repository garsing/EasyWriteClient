using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Text.RegularExpressions;
using WordAddIn1;

namespace BrowserNavigateTest
{
    /// <summary>第 3 波 interact，只测 agent。</summary>
    internal static class IntCases
    {
        public static void Run(TestRun run, FixtureHttp http, int chromeBefore, int edgeBefore) // 不关用户 Chrome/Edge
        {
            var s = new NavSession();
            RunNeg(run, s, http);
            RunD18(run, s, http);
            RunClick(run, s, http);
            RunHover(run, s, http);
            RunType(run, s, http);
            RunScroll(run, s, http);
            RunPress(run, s, http);
            RunSelect(run, s, http);
            RunForm(run, s, http);
            int chromeNow = CountProc("chrome");
            int edgeNow = CountProc("msedge");
            if (chromeNow < chromeBefore || edgeNow < edgeBefore)
            {
                Console.WriteLine(
                    "  [WARN] 用户浏览器进程变少 chrome "
                    + chromeBefore + "→" + chromeNow
                    + " edge " + edgeBefore + "→" + edgeNow);
            }
        }

        private static void RunNeg(TestRun run, NavSession s, FixtureHttp http)
        {
            ExpectFail(run, "int-neg-no-action", s.Interact(channelId: "b1"), "action");
            string id = OpenOk(run, s, http.Url("overlay-cookie", "index.html"), "int-neg-bad-action");
            if (id == null)
            {
                FailMany(run, "依赖开页",
                    "int-neg-no-channel", "int-neg-wrong-kind",
                    "int-neg-unknown-id", "int-neg-click-no-ref", "int-neg-type-no-text",
                    "int-neg-press-no-key", "int-neg-press-bad-key", "int-neg-select-no-option",
                    "int-neg-bad-ref", "int-neg-stale-ref", "int-neg-type-iframe",
                    "int-neg-hover-iframe");
                return;
            }

            ExpectFail(run, "int-neg-bad-action", s.Interact(id, "foo"), "action");

            ToolResult noCh = s.Interact(action: "click");
            if (!noCh.Success)
            {
                string err = noCh.Error ?? "";
                if (ContainsAny(err, "渠道", "打开网页", "ref"))
                {
                    run.CaseOk("int-neg-no-channel");
                }
                else
                {
                    run.CaseFail("int-neg-no-channel", err);
                }
            }
            else
            {
                run.CaseOk("int-neg-no-channel");
            }

            ExpectFail(run, "int-neg-wrong-kind", s.Interact("w1", "click"), "渠道");
            ExpectFail(run, "int-neg-unknown-id", s.Interact("b999", "click"), "渠道");
            ExpectFail(run, "int-neg-click-no-ref", s.Interact(id, "click"), "ref");

            string tree = TreeWithInput(s, id);
            string inputRef = FindRef(tree, "普通按钮");
            ExpectFail(
                run,
                "int-neg-type-no-text",
                s.Interact(id, "type", string.IsNullOrEmpty(inputRef) ? "e1" : inputRef),
                "text");
            ExpectFail(run, "int-neg-press-no-key", s.Interact(id, "press"), "key");
            ExpectFail(run, "int-neg-press-bad-key", s.Interact(id, "press", key: "Control"), "白名单");

            string selId = OpenOk(run, s, http.Url("native-select", "index.html"), "int-neg-select-no-option");
            if (selId == null)
            {
                run.CaseFail("int-neg-select-no-option", "依赖 native-select");
            }
            else
            {
                string selTree = TreeWithInput(s, selId);
                string selRef = FindRef(selTree, "第一作者类型", "combobox")
                    ?? FindRef(selTree, "第一作者类型");
                ExpectFail(
                    run,
                    "int-neg-select-no-option",
                    s.Interact(selId, "select", string.IsNullOrEmpty(selRef) ? "e1" : selRef),
                    "option");
            }

            ExpectFail(run, "int-neg-bad-ref", s.Interact(id, "click", "e9999"), "过期");

            ToolResult ov = s.Snap(id);
            string oldRef = FindRef(Tree(ov), "普通按钮");
            ToolResult jump = s.Nav(http.Url("public-example", "index.html"), host: "yiwrite", channelId: id);
            if (!jump.Success)
            {
                run.CaseFail("int-neg-stale-ref", jump.Error ?? "跳转失败");
            }
            else if (string.IsNullOrEmpty(oldRef))
            {
                run.CaseFail("int-neg-stale-ref", "跳转前没有 eN");
            }
            else
            {
                ExpectFail(run, "int-neg-stale-ref", s.Interact(id, "click", oldRef), "过期");
            }

            string ifr = OpenOk(run, s, http.Url("iframe", "outer.html"), "int-neg-type-iframe");
            if (ifr == null)
            {
                FailMany(run, "依赖 outer", "int-neg-type-iframe", "int-neg-hover-iframe");
                return;
            }

            string ifrTree = Tree(s.Snap(ifr));
            string frameRef = FindRef(ifrTree, "教师工作台") ?? FindIframeRef(ifrTree);
            if (string.IsNullOrEmpty(frameRef))
            {
                FailMany(run, "无 iframe 行", "int-neg-type-iframe", "int-neg-hover-iframe");
                return;
            }

            ExpectFail(run, "int-neg-type-iframe", s.Interact(ifr, "type", frameRef, text: "x"), "iframe");
            ExpectFail(run, "int-neg-hover-iframe", s.Interact(ifr, "hover", frameRef), "click");
        }

        private static void RunD18(TestRun run, NavSession s, FixtureHttp http)
        {
            string id = OpenOk(run, s, http.Url("d18-risk", "index.html"), "int-d18-type-password");
            if (id == null)
            {
                FailMany(run, "依赖 d18-risk",
                    "int-d18-click-login", "int-d18-click-pay",
                    "int-d18-click-delete", "int-d18-press-enter-login");
                return;
            }

            string tree = TreeWithInput(s, id);
            string pwd = FindRef(tree, "密码");
            string login = FindRef(tree, "登录", "button") ?? FindRef(tree, "登录");
            string pay = FindRef(tree, "立即支付");
            string del = FindRef(tree, "删除");
            string safe = FindRef(tree, "安全按钮");

            ToolResult typed = string.IsNullOrEmpty(pwd)
                ? null
                : s.Interact(id, "type", pwd, text: "secret");
            if (typed != null && !typed.Success)
            {
                string err = typed.Error ?? "";
                if (err.IndexOf("secret", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    run.CaseFail("int-d18-type-password", "错误回显了 secret");
                }
                else if (ContainsAny(err, "拒绝", "密码"))
                {
                    run.CaseOk("int-d18-type-password");
                }
                else
                {
                    run.CaseFail("int-d18-type-password", err);
                }
            }
            else
            {
                run.CaseFail("int-d18-type-password", pwd == null ? "无密码 ref" : "应失败却成功");
            }

            ExpectFailAndUnchanged(run, s, id, "int-d18-click-login", login, "click", null, "拒绝");
            ExpectFailAndUnchanged(run, s, id, "int-d18-click-pay", pay, "click", null, "拒绝");
            ExpectFailAndUnchanged(run, s, id, "int-d18-click-delete", del, "click", null, "拒绝");

            ToolResult press = string.IsNullOrEmpty(pwd)
                ? null
                : s.Interact(id, "press", pwd, key: "Enter");
            if (press != null && !press.Success && ContainsAny(press.Error, "拒绝", "密码", "提交"))
            {
                run.CaseOk("int-d18-press-enter-login");
            }
            else
            {
                run.CaseFail(
                    "int-d18-press-enter-login",
                    pwd == null ? "无密码 ref" : (press != null && press.Success ? "应失败却成功" : press?.Error));
            }

            if (!string.IsNullOrEmpty(safe))
            {
                s.Interact(id, "click", safe);
            }
        }

        private static void RunClick(TestRun run, NavSession s, FixtureHttp http)
        {
            string id = OpenOk(run, s, http.Url("overlay-cookie", "index.html"), "int-click-plain");
            if (id == null)
            {
                FailMany(run, "依赖 overlay-cookie",
                    "int-ok-snapshot", "int-click-hidden",
                    "int-click-blocked", "int-click-new-ref");
            }
            else
            {
                string tree = Tree(s.Snap(id));
                string plain = FindRef(tree, "普通按钮");
                ToolResult clicked = string.IsNullOrEmpty(plain) ? null : s.Interact(id, "click", plain);
                bool okClick = clicked != null && clicked.Success
                    && (Tree(clicked).Contains("普通按钮已点") || Tree(clicked).Contains("已点"));
                Ok(run, "int-click-plain", okClick, clicked == null ? "无普通按钮 ref" : (clicked.Error ?? "状态未变"));
                Ok(
                    run,
                    "int-ok-snapshot",
                    okClick
                    && Flag(clicked, "resnapshot")
                    && !Flag(clicked, "refs_invalidated")
                    && Tree(clicked).Length > 0
                    && !HasLeak(clicked),
                    clicked == null ? "无回包" : "回树/泄漏不符");

                string hid = OpenOk(run, s, http.Url("overlay-cookie", "index.html"), "int-click-hidden", visible: false);
                if (hid == null)
                {
                    run.CaseFail("int-click-hidden", "隐藏窗打开失败");
                }
                else
                {
                    string ht = Tree(s.Snap(hid));
                    string href = FindRef(ht, "普通按钮");
                    ToolResult hc = string.IsNullOrEmpty(href) ? null : s.Interact(hid, "click", href);
                    Ok(run, "int-click-hidden", hc != null && hc.Success, hc == null ? "无 ref" : hc.Error);
                }

                string blkId = OpenOk(run, s, http.Url("overlay-cookie", "index.html"), "int-click-blocked");
                if (blkId == null)
                {
                    run.CaseFail("int-click-new-ref", "依赖 overlay-cookie");
                }
                else
                {
                    string bt = Tree(s.Snap(blkId));
                    string under = FindRef(bt, "查看结果");
                    string keep = FindRef(bt, "普通按钮");
                    ToolResult blocked = string.IsNullOrEmpty(under) ? null : s.Interact(blkId, "click", under);
                    bool blockedOk = blocked != null && !blocked.Success
                        && (blocked.Error ?? "").IndexOf("挡住", StringComparison.Ordinal) >= 0;
                    if (!blockedOk)
                    {
                        run.CaseFail(
                            "int-click-blocked",
                            under == null ? "无查看结果 ref" : (blocked != null && blocked.Success ? "应挡住却成功" : blocked?.Error));
                    }
                    else if (string.IsNullOrEmpty(keep))
                    {
                        run.CaseFail("int-click-blocked", "挡住对了但无普通按钮旧 eN");
                    }
                    else
                    {
                        ToolResult still = s.Interact(blkId, "click", keep);
                        Ok(run, "int-click-blocked", still.Success, still.Error ?? "挡住后旧 eN 不能点");
                    }

                    string fresh = OpenOk(run, s, http.Url("overlay-cookie", "index.html"), "int-click-new-ref");
                    if (fresh == null)
                    {
                        run.CaseFail("int-click-new-ref", "开页失败");
                    }
                    else
                    {
                        string ft = Tree(s.Snap(fresh));
                        string oldBtn = FindRef(ft, "普通按钮");
                        ToolResult first = string.IsNullOrEmpty(oldBtn) ? null : s.Interact(fresh, "click", oldBtn);
                        if (first == null || !first.Success)
                        {
                            run.CaseFail("int-click-new-ref", first == null ? "无 ref" : first.Error);
                        }
                        else
                        {
                            ToolResult stale = s.Interact(fresh, "click", oldBtn);
                            string neu = FindRef(Tree(first), "普通按钮");
                            ToolResult again = string.IsNullOrEmpty(neu) ? null : s.Interact(fresh, "click", neu);
                            Ok(
                                run,
                                "int-click-new-ref",
                                stale != null && !stale.Success && again != null && again.Success,
                                stale != null && stale.Success
                                    ? "旧 eN 仍能点（表未 Replace）"
                                    : (again == null ? "新树无普通按钮" : again.Error));
                        }
                    }
                }
            }

            string ifr = OpenOk(run, s, http.Url("iframe", "outer.html"), "int-click-iframe-expand");
            if (ifr == null)
            {
                FailMany(run, "依赖 outer",
                    "int-click-iframe-inner", "int-expand-not-shell");
            }
            else
            {
                string t0 = Tree(s.Snap(ifr));
                string frameRef = FindRef(t0, "教师工作台") ?? FindIframeRef(t0);
                ToolResult exp = string.IsNullOrEmpty(frameRef) ? null : s.Interact(ifr, "click", frameRef);
                string et = exp == null ? "" : Tree(exp);
                bool expanded = exp != null && exp.Success
                    && (et.Contains("我的得分") || et.Contains("评教结果"))
                    && Flag(exp, "expanded_frame");
                Ok(run, "int-click-iframe-expand", expanded, exp == null ? "无 iframe ref" : (exp.Error ?? "未展开"));
                Ok(
                    run,
                    "int-expand-not-shell",
                    expanded && !LooksLikeOuterShellOnly(et),
                    "展开后仍像外壳");

                string score = FindRef(et, "我的得分");
                ToolResult inner = string.IsNullOrEmpty(score) ? null : s.Interact(ifr, "click", score);
                Ok(
                    run,
                    "int-click-iframe-inner",
                    inner != null && inner.Success && Tree(inner).Contains("已点我的得分"),
                    inner == null ? "无我的得分 ref" : (inner.Error ?? "状态未变"));
            }

            string card = OpenOk(run, s, http.Url("hover-card", "index.html"), "int-click-image-card");
            if (card == null)
            {
                run.CaseFail("int-click-image-card", "开页失败");
            }
            else
            {
                string ct = Tree(s.Snap(card));
                string img = FindRef(ct, "机器学习课程设计", "image")
                    ?? FindRef(ct, "机器学习课程设计", "img")
                    ?? FindRef(ct, "机器学习课程设计");
                ToolResult go = string.IsNullOrEmpty(img) ? null : s.Interact(card, "click", img);
                string got = go == null ? "" : (Tree(go) + " " + Str(go, "url"));
                Ok(
                    run,
                    "int-click-image-card",
                    go != null && go.Success && (got.Contains("已进入详情") || got.Contains("#/detail")),
                    go == null ? "无封面图 ref" : (go.Error ?? "未进详情"));
            }

            string tabs = OpenOk(run, s, http.Url("center-tabs", "index.html"), "int-click-tab-member");
            if (tabs == null)
            {
                run.CaseFail("int-click-tab-member", "开页失败");
            }
            else
            {
                string tt = Tree(s.Snap(tabs));
                string mem = FindRef(tt, "项目成员");
                ToolResult tab = string.IsNullOrEmpty(mem) ? null : s.Interact(tabs, "click", mem);
                Ok(
                    run,
                    "int-click-tab-member",
                    tab != null && tab.Success && Tree(tab).Contains("项目成员"),
                    tab == null ? "无项目成员 ref" : (tab.Error ?? "状态未变"));
            }

            string file = OpenOk(run, s, http.Url("file-chooser", "index.html"), "int-click-file");
            if (file == null)
            {
                run.CaseFail("int-click-file", "开页失败");
            }
            else
            {
                string ft = TreeWithInput(s, file);
                string pick = FindRef(ft, "选择文件");
                string view = FindRef(ft, "查看结果");
                ToolResult denied = string.IsNullOrEmpty(pick) ? null : s.Interact(file, "click", pick);
                bool denyOk = denied != null && !denied.Success && ContainsAny(denied.Error, "文件选择", "选文件");
                if (denyOk && !string.IsNullOrEmpty(view))
                {
                    ToolResult still = s.Interact(file, "click", view);
                    Ok(run, "int-click-file", still.Success, still.Success ? "" : "失败后旧 eN 不能点查看结果: " + still.Error);
                }
                else
                {
                    run.CaseFail("int-click-file", pick == null ? "无选择文件 ref" : (denied != null && denied.Success ? "应拒绝却成功" : denied?.Error));
                }
            }

            string nsel = OpenOk(run, s, http.Url("native-select", "index.html"), "int-click-native-select");
            if (nsel == null)
            {
                run.CaseFail("int-click-native-select", "开页失败");
            }
            else
            {
                string nt = TreeWithInput(s, nsel);
                string combo = FindRef(nt, "第一作者类型", "combobox") ?? FindRef(nt, "第一作者类型");
                ToolResult clk = string.IsNullOrEmpty(combo) ? null : s.Interact(nsel, "click", combo);
                Ok(
                    run,
                    "int-click-native-select",
                    clk != null && !clk.Success && ContainsAny(clk.Error, "select", "下拉"),
                    combo == null ? "无下拉 ref" : (clk != null && clk.Success ? "应拒绝 click" : clk?.Error));
            }
        }

        private static void RunHover(TestRun run, NavSession s, FixtureHttp http)
        {
            string id = OpenOk(run, s, http.Url("hover-card", "index.html"), "int-hover-card");
            if (id == null)
            {
                return;
            }

            string tree = Tree(s.Snap(id));
            string card = FindRef(tree, "机器学习课程设计") ?? FindRef(tree, "查看详情");
            ToolResult hv = string.IsNullOrEmpty(card) ? null : s.Interact(id, "hover", card);
            Ok(
                run,
                "int-hover-card",
                hv != null && hv.Success && Flag(hv, "resnapshot")
                && (Tree(hv).Contains("查看详情") || Tree(hv).Contains("机器学习课程设计")),
                hv == null ? "无卡片 ref" : (hv.Error ?? "hover 后树对不上"));
        }

        private static void RunType(TestRun run, NavSession s, FixtureHttp http)
        {
            string id = OpenOk(run, s, http.Url("native-select", "index.html"), "int-type-input");
            if (id == null)
            {
                run.CaseFail("int-type-hidden", "依赖 native-select");
            }
            else
            {
                string tree = TreeWithInput(s, id);
                string box = FindRef(tree, "论文题目") ?? FindRef(tree, "lwtm");
                ToolResult typed = string.IsNullOrEmpty(box) ? null : s.Interact(id, "type", box, text: "已填");
                bool ok = typed != null && typed.Success
                    && (Tree(typed).Contains("已填") || Tree(typed).IndexOf("value=\"已填\"", StringComparison.Ordinal) >= 0);
                if (ok && (typed.Error ?? "").IndexOf("已填", StringComparison.Ordinal) >= 0)
                {
                    run.CaseFail("int-type-input", "成功路径 Error 回显了 text");
                }
                else
                {
                    Ok(run, "int-type-input", ok, typed == null ? "无论文题目 ref" : (typed.Error ?? "value 未变"));
                }

                string hid = OpenOk(run, s, http.Url("native-select", "index.html"), "int-type-hidden", visible: false);
                if (hid == null)
                {
                    run.CaseFail("int-type-hidden", "隐藏窗打开失败");
                }
                else
                {
                    string ht = TreeWithInput(s, hid);
                    string href = FindRef(ht, "论文题目") ?? FindRef(ht, "lwtm");
                    ToolResult th = string.IsNullOrEmpty(href) ? null : s.Interact(hid, "type", href, text: "隐填");
                    Ok(run, "int-type-hidden", th != null && th.Success, th == null ? "无 ref" : th.Error);
                }
            }

            string dis = OpenOk(run, s, http.Url("disabled-input", "index.html"), "int-type-disabled");
            if (dis == null)
            {
                FailMany(run, "依赖 disabled-input",
                    "int-type-readonly",
                    "int-type-fieldset-disabled", "int-type-aria-disabled");
                return;
            }

            string dt = TreeWithInput(s, dis);
            ExpectFailKeep(
                run, s, dis, "int-type-disabled", FindRef(dt, "禁用框"), "type", "改", "不可编辑");
            ExpectFailKeep(
                run, s, dis, "int-type-readonly", FindRef(dt, "只读框"), "type", "改", "不可编辑");
            ExpectFailKeep(
                run, s, dis, "int-type-fieldset-disabled", FindRef(dt, "组内禁用"), "type", "改", "不可编辑");
            ExpectFailKeep(
                run, s, dis, "int-type-aria-disabled", FindRef(dt, "无障碍禁用"), "type", "改", "不可编辑");
        }

        private static void RunScroll(TestRun run, NavSession s, FixtureHttp http)
        {
            string id = OpenOk(run, s, http.Url("overflow-scroll", "index.html"), "int-scroll-window");
            if (id == null)
            {
                run.CaseFail("int-scroll-inner", "依赖 overflow-scroll");
                return;
            }

            ToolResult win = s.Interact(id, "scroll", direction: "down");
            bool noPx = win != null && win.Success
                && string.IsNullOrEmpty(Str(win, "pixels"))
                && string.IsNullOrEmpty(Str(win, "delta"));
            Ok(run, "int-scroll-window", noPx, win == null ? "无回包" : (win.Error ?? "回包带了像素"));

            string tree = Tree(s.Snap(id));
            string bottom = FindRef(tree, "列表底项");
            if (string.IsNullOrEmpty(bottom))
            {
                run.CaseFail("int-scroll-inner", "无列表底项 ref");
                return;
            }

            ToolResult sc = s.Interact(id, "scroll", bottom);
            if (!sc.Success)
            {
                run.CaseFail("int-scroll-inner", sc.Error ?? "scroll 失败");
                return;
            }

            string neu = FindRef(Tree(sc), "列表底项") ?? bottom;
            ToolResult clk = s.Interact(id, "click", neu);
            Ok(
                run,
                "int-scroll-inner",
                clk.Success && Tree(clk).Contains("已点底项"),
                clk.Success ? "滚了但点不到底项" : clk.Error);
        }

        private static void RunPress(TestRun run, NavSession s, FixtureHttp http)
        {
            string id = OpenOk(run, s, http.Url("native-select", "index.html"), "int-press-enter-search");
            if (id == null)
            {
                run.CaseFail("int-press-tab", "依赖 native-select");
                return;
            }

            string tree = TreeWithInput(s, id);
            string box = FindRef(tree, "论文题目") ?? FindRef(tree, "lwtm");
            if (string.IsNullOrEmpty(box))
            {
                FailMany(run, "无论文题目 ref", "int-press-enter-search", "int-press-tab");
                return;
            }

            ToolResult typed = s.Interact(id, "type", box, text: "词");
            string focus = typed.Success ? (FindRef(Tree(typed), "论文题目") ?? FindRef(Tree(typed), "lwtm") ?? box) : box;
            ToolResult ent = s.Interact(id, "press", focus, key: "Enter");
            Ok(run, "int-press-enter-search", typed.Success && ent.Success, ent.Success ? typed.Error : ent.Error);

            ToolResult tab = s.Interact(id, "press", key: "Tab");
            Ok(run, "int-press-tab", tab.Success, tab.Error);
        }

        private static void RunSelect(TestRun run, NavSession s, FixtureHttp http)
        {
            string id = OpenOk(run, s, http.Url("native-select", "index.html"), "int-select-text");
            if (id == null)
            {
                FailMany(run, "依赖 native-select", "int-select-missing", "int-select-hidden");
                return;
            }

            string tree = TreeWithInput(s, id);
            string first = FindRef(tree, "第一作者类型", "combobox") ?? FindRef(tree, "第一作者类型");
            string cat = FindRef(tree, "统计类别", "combobox") ?? FindRef(tree, "统计类别");
            ToolResult ok = string.IsNullOrEmpty(first)
                ? null
                : s.Interact(id, "select", first, option: "本校教师");
            Ok(
                run,
                "int-select-text",
                ok != null && ok.Success && Tree(ok).Contains("本校教师"),
                first == null ? "无下拉 ref" : (ok.Error ?? "状态未变"));

            ToolResult miss = string.IsNullOrEmpty(cat)
                ? null
                : s.Interact(id, "select", cat, option: "应用研究");
            Ok(
                run,
                "int-select-missing",
                miss != null && !miss.Success
                && (miss.Error ?? "").Contains("自然科学类")
                && (miss.Error ?? "").Contains("人文社科类"),
                cat == null ? "无统计类别 ref" : (miss != null && miss.Success ? "应失败" : miss?.Error));

            string hid = OpenOk(run, s, http.Url("native-select", "index.html"), "int-select-hidden", visible: false);
            if (hid == null)
            {
                return;
            }

            string ht = TreeWithInput(s, hid);
            string href = FindRef(ht, "第一作者类型", "combobox") ?? FindRef(ht, "第一作者类型");
            ToolResult hs = string.IsNullOrEmpty(href)
                ? null
                : s.Interact(hid, "select", href, option: "本校教师");
            Ok(run, "int-select-hidden", hs != null && hs.Success, hs == null ? "无 ref" : hs.Error);
        }

        private static void RunForm(TestRun run, NavSession s, FixtureHttp http)
        {
            string id = OpenOk(run, s, http.Url("form-controls", "index.html"), "int-type-date-iso");
            if (id == null)
            {
                FailMany(run, "依赖 form-controls",
                    "int-type-date-today", "int-type-date-yesterday",
                    "int-type-date-bad", "int-type-datetime-local", "int-type-month",
                    "int-type-week", "int-type-time", "int-type-textarea",
                    "int-type-contenteditable", "int-click-checkbox", "int-click-switch",
                    "int-click-radio", "int-type-number");
                return;
            }

            string tree = TreeWithInput(s, id);
            TypeValue(run, s, id, ref tree, "int-type-date-iso", DateHint(tree), "2026-10-03", "2026-10-03");
            TypeValue(
                run, s, id, ref tree, "int-type-date-today", DateHint(tree), "今天",
                DateTime.Now.Date.ToString("yyyy-MM-dd"));
            TypeValue(
                run, s, id, ref tree, "int-type-date-yesterday", DateHint(tree), "昨天",
                DateTime.Now.Date.AddDays(-1).ToString("yyyy-MM-dd"));

            string dateRef = DateHint(tree);
            ToolResult bad = string.IsNullOrEmpty(dateRef)
                ? null
                : s.Interact(id, "type", dateRef, text: "不是日期");
            Ok(
                run,
                "int-type-date-bad",
                bad != null && !bad.Success && ContainsAny(bad.Error, "日期", "时间", "格式")
                && (Tree(s.Snap(id)).IndexOf("不是日期", StringComparison.Ordinal) < 0),
                dateRef == null ? "无立项日 ref" : (bad != null && bad.Success ? "乱填却成功" : bad?.Error));

            TypeValue(
                run, s, id, ref tree, "int-type-datetime-local",
                FindRef(tree, "答辩时刻") ?? FindTemporal(tree, "datetime"),
                "2026-10-03T14:30", "2026-10-03T14:30");
            TypeValue(
                run, s, id, ref tree, "int-type-month",
                FindRef(tree, "所属月") ?? FindTemporal(tree, "月"),
                "2026-10", "2026-10");
            TypeValue(
                run, s, id, ref tree, "int-type-week",
                FindRef(tree, "教学周") ?? FindTemporal(tree, "周"),
                "2026-W40", "2026-W40");
            TypeValue(
                run, s, id, ref tree, "int-type-time",
                FindRef(tree, "开始时刻") ?? FindTemporal(tree, "时刻"),
                "14:30", "14:30");
            TypeValue(
                run, s, id, ref tree, "int-type-textarea",
                FindRef(tree, "多行说明"),
                "两行说明", "两行说明");
            TypeValue(
                run, s, id, ref tree, "int-type-contenteditable",
                FindRef(tree, "可编辑摘要"),
                "摘要词", "摘要词");

            ClickStatus(run, s, id, ref tree, "int-click-checkbox", FindRef(tree, "同意条款"), "已勾选");
            ClickStatus(run, s, id, ref tree, "int-click-switch", FindRef(tree, "夜间模式"), "夜间已开");
            ClickStatus(
                run, s, id, ref tree, "int-click-radio",
                FindRef(tree, "教师"), "身份=tea", "教师");
            TypeValue(run, s, id, ref tree, "int-type-number", FindRef(tree, "份数"), "3", "3");
        }

        private static void TypeValue(
            TestRun run,
            NavSession s,
            string id,
            ref string tree,
            string caseId,
            string refId,
            string text,
            string expect)
        {
            if (string.IsNullOrEmpty(refId))
            {
                run.CaseFail(caseId, "无控件 ref");
                return;
            }

            ToolResult r = s.Interact(id, "type", refId, text: text);
            string got = r == null ? "" : Tree(r);
            bool pass = r != null && r.Success && got.IndexOf(expect, StringComparison.Ordinal) >= 0;
            Ok(run, caseId, pass, r == null ? "无回包" : (r.Success ? "value/状态未含 " + expect : r.Error));
            if (r != null && r.Success && got.Length > 0)
            {
                tree = got;
            }
        }

        private static void ClickStatus(
            TestRun run,
            NavSession s,
            string id,
            ref string tree,
            string caseId,
            string refId,
            string expect,
            string alt = null)
        {
            if (string.IsNullOrEmpty(refId))
            {
                run.CaseFail(caseId, "无控件 ref");
                return;
            }

            ToolResult r = s.Interact(id, "click", refId);
            string got = r == null ? "" : Tree(r);
            bool pass = r != null && r.Success
                && (got.Contains(expect) || (!string.IsNullOrEmpty(alt) && got.Contains(alt)));
            Ok(run, caseId, pass, r == null ? "无回包" : (r.Success ? "状态未变" : r.Error));
            if (r != null && r.Success && got.Length > 0)
            {
                tree = got;
            }
        }

        private static string DateHint(string tree)
        {
            return FindRef(tree, "立项日")
                ?? FindTemporal(tree, "年")
                ?? FindTemporal(tree, "日");
        }

        private static string FindTemporal(string tree, string hint)
        {
            return FindRef(tree, hint, "spinbutton")
                ?? FindRef(tree, hint, "textbox")
                ?? FindRef(tree, hint);
        }

        private static void ExpectFailAndUnchanged(
            TestRun run,
            NavSession s,
            string id,
            string caseId,
            string refId,
            string action,
            string text,
            string must)
        {
            if (string.IsNullOrEmpty(refId))
            {
                run.CaseFail(caseId, "无 ref");
                return;
            }

            ToolResult r = s.Interact(id, action, refId, text: text);
            string after = Tree(s.Snap(id));
            Ok(
                run,
                caseId,
                r != null && !r.Success && ContainsAny(r.Error, must, "拒绝") && after.Contains("未操作"),
                r != null && r.Success ? "应失败却成功" : (r?.Error ?? "无回包"));
        }

        private static void ExpectFailKeep(
            TestRun run,
            NavSession s,
            string id,
            string caseId,
            string refId,
            string action,
            string text,
            string must)
        {
            if (string.IsNullOrEmpty(refId))
            {
                run.CaseFail(caseId, "无 ref");
                return;
            }

            ToolResult r = s.Interact(id, action, refId, text: text);
            if (r != null && !r.Success && ContainsAny(r.Error, must))
            {
                run.CaseOk(caseId);
                return;
            }

            run.CaseFail(caseId, r != null && r.Success ? "应失败却成功" : (r?.Error ?? "无回包"));
        }

        private static string OpenOk(
            TestRun run,
            NavSession s,
            string url,
            string failId,
            object visible = null)
        {
            ToolResult r = s.Nav(url, host: "yiwrite", visible: visible);
            if (r.Success && NavSession.IsShortId(Str(r, "channel_id")))
            {
                return Str(r, "channel_id");
            }

            run.CaseFail(failId, r.Success ? "无短 channel_id" : (r.Error ?? "打开失败"));
            return null;
        }

        private static string TreeWithInput(NavSession s, string id)
        {
            string t = Tree(s.Snap(id));
            if (HasAnyRef(t) && (t.Contains("论文题目") || t.Contains("立项日") || t.Contains("密码")
                || t.Contains("禁用框") || t.Contains("选择文件") || t.Contains("份数")
                || t.Contains("普通按钮")))
            {
                return t;
            }

            string extra = Tree(s.Snap(id, domSupplement: "input"));
            return extra.Length > t.Length ? extra : t;
        }

        private static bool LooksLikeOuterShellOnly(string tree)
        {
            if (string.IsNullOrEmpty(tree))
            {
                return true;
            }

            return tree.Contains("外壳")
                && tree.IndexOf("我的得分", StringComparison.Ordinal) < 0
                && tree.IndexOf("评教结果", StringComparison.Ordinal) < 0;
        }

        private static bool HasLeak(ToolResult r)
        {
            string dump = Tree(r);
            Dictionary<string, object> d = Data(r);
            if (d != null)
            {
                dump += string.Join(",", new List<string>(d.Keys).ToArray());
            }

            return dump.IndexOf("backendNodeId", StringComparison.Ordinal) >= 0
                || dump.IndexOf("cssPath", StringComparison.Ordinal) >= 0
                || dump.IndexOf("click_mode", StringComparison.Ordinal) >= 0
                || dump.IndexOf("<html", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private static void Ok(TestRun run, string id, bool pass, string reason)
        {
            if (pass)
            {
                run.CaseOk(id);
            }
            else
            {
                run.CaseFail(id, reason ?? "断言失败");
            }
        }

        private static void FailMany(TestRun run, string reason, params string[] ids)
        {
            for (int i = 0; i < ids.Length; i++)
            {
                run.CaseFail(ids[i], reason);
            }
        }

        private static void ExpectFail(TestRun run, string id, ToolResult r, string mustContain)
        {
            if (r != null && !r.Success)
            {
                string err = r.Error ?? "";
                if (string.IsNullOrEmpty(mustContain)
                    || err.IndexOf(mustContain, StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    run.CaseOk(id);
                    return;
                }

                run.CaseFail(id, "失败文案不含「" + mustContain + "」: " + err);
                return;
            }

            run.CaseFail(id, r != null && r.Success ? "应失败却成功" : "无回包");
        }

        private static bool ContainsAny(string text, params string[] parts)
        {
            if (string.IsNullOrEmpty(text) || parts == null)
            {
                return false;
            }

            for (int i = 0; i < parts.Length; i++)
            {
                if (!string.IsNullOrEmpty(parts[i])
                    && text.IndexOf(parts[i], StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    return true;
                }
            }

            return false;
        }

        private static Dictionary<string, object> Data(ToolResult r)
        {
            return r?.Data as Dictionary<string, object>;
        }

        private static string Str(ToolResult r, string key)
        {
            Dictionary<string, object> d = Data(r);
            if (d == null || !d.ContainsKey(key) || d[key] == null)
            {
                return "";
            }

            return Convert.ToString(d[key]) ?? "";
        }

        private static bool Flag(ToolResult r, string key)
        {
            Dictionary<string, object> d = Data(r);
            if (d == null || !d.ContainsKey(key) || d[key] == null)
            {
                return false;
            }

            object v = d[key];
            if (v is bool b)
            {
                return b;
            }

            string s = Convert.ToString(v);
            return string.Equals(s, "True", StringComparison.OrdinalIgnoreCase)
                || string.Equals(s, "true", StringComparison.OrdinalIgnoreCase);
        }

        private static string Tree(ToolResult r)
        {
            return Str(r, "snapshot");
        }

        private static bool HasAnyRef(string tree)
        {
            return Regex.IsMatch(tree ?? "", @"ref=e\d+");
        }

        private static string FindRef(string tree, string hint, string preferRole = null)
        {
            if (string.IsNullOrEmpty(tree) || string.IsNullOrEmpty(hint))
            {
                return null;
            }

            string fallback = null;
            foreach (string line in tree.Split('\n'))
            {
                if (line.IndexOf(hint, StringComparison.Ordinal) < 0)
                {
                    continue;
                }

                Match m = Regex.Match(line, @"ref=(e\d+)");
                if (!m.Success)
                {
                    continue;
                }

                if (!string.IsNullOrEmpty(preferRole)
                    && line.IndexOf(preferRole, StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    return m.Groups[1].Value;
                }

                if (fallback == null)
                {
                    fallback = m.Groups[1].Value;
                }
            }

            return fallback;
        }

        private static string FindIframeRef(string tree)
        {
            foreach (string line in (tree ?? "").Split('\n'))
            {
                if (!Regex.IsMatch(line, @"iframe|IFrame|frame", RegexOptions.IgnoreCase))
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
