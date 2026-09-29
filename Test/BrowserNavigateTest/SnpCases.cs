using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Text.RegularExpressions;
using WordAddIn1;

namespace BrowserNavigateTest
{
    internal static class SnpCases
    {
        public static void Run(TestRun run, FixtureHttp http, bool attach, int chromeBefore, int edgeBefore)
        {
            var s = new NavSession();
            string navHost = attach ? "chrome" : "yiwrite";
            RunNeg(run, s, http, navHost);
            RunOverview(run, s, http, navHost);
            RunRealSlim(run, s, http, navHost);
            RunRealSaved(run, s, http, navHost);
            RunNoQuit(run, chromeBefore, edgeBefore);
        }

        private static void RunNeg(TestRun run, NavSession s, FixtureHttp http, string navHost)
        {
            ToolResult noCh = s.Snap();
            if (!noCh.Success)
            {
                ExpectFail(run, "snp-neg-no-channel", noCh, "渠道");
            }
            else
            {
                run.CaseOk("snp-neg-no-channel");
            }

            ExpectFail(run, "snp-neg-wrong-kind", s.Snap("w1"), "渠道");
            ExpectFail(run, "snp-neg-unknown-id", s.Snap("b999"), "渠道");

            string id = OpenOk(run, s, http.Url("iframe", "outer.html"), navHost, "snp-neg-bad-ref");
            if (id == null)
            {
                run.CaseFail("snp-neg-stale-ref", "依赖开页");
                run.CaseFail("snp-neg-dom-all", "依赖开页");
                run.CaseFail("snp-neg-dom-default", "依赖开页");
                run.CaseFail("snp-neg-dom-unknown", "依赖开页");
                return;
            }

            ToolResult ov = s.Snap(id);
            ExpectFail(run, "snp-neg-bad-ref", s.Snap(id, "e9999"), "过期");

            string oldRef = FirstRef(Tree(ov));
            ToolResult jump = s.Nav(http.Url("public-example", "index.html"), host: navHost, channelId: id);
            if (!jump.Success)
            {
                run.CaseFail("snp-neg-stale-ref", jump.Error ?? "跳转失败");
            }
            else if (string.IsNullOrEmpty(oldRef))
            {
                run.CaseFail("snp-neg-stale-ref", "壳上没有 eN");
            }
            else
            {
                ExpectFail(run, "snp-neg-stale-ref", s.Snap(id, oldRef), "过期");
            }

            ExpectFail(run, "snp-neg-dom-all", s.Snap(id, domSupplement: "all"), "未知");
            ExpectFail(run, "snp-neg-dom-default", s.Snap(id, domSupplement: "default"), "未知");
            ExpectFail(run, "snp-neg-dom-unknown", s.Snap(id, domSupplement: "foo"), "未知");
        }

        private static void RunOverview(TestRun run, NavSession s, FixtureHttp http, string navHost)
        {
            string id = OpenOk(run, s, http.Url("iframe", "outer.html"), navHost, "snp-ov-mode");
            if (id == null)
            {
                FailMany(run, "依赖 outer",
                    "snp-ov-mode", "snp-ov-refs", "snp-ov-id-short", "snp-ov-kind",
                    "snp-ov-no-html", "snp-ov-no-backend", "snp-ov-url-title",
                    "snp-ov-dom-off", "snp-dt-mode", "snp-dt-echo-ref", "snp-dt-subtree",
                    "snp-dt-ignore-dom", "snp-dt-rebuild",
                    "snp-ifr-shell", "snp-ifr-row", "snp-ifr-enter", "snp-ifr-hidden-row");
                RunHidden(run, s, http);
                RunDomInput(run, s, http, navHost);
                RunNested(run, s, http, navHost);
                RunSandbox(run, s, http, navHost);
                return;
            }

            ToolResult ov = s.Snap(id);
            string tree = Tree(ov);
            string mode = Str(ov, "mode");
            string kind = Str(ov, "kind");
            string track = Str(ov, "track");
            Ok(run, "snp-ov-mode", ov.Success && mode == "overview", ov.Success ? "mode=" + mode : ov.Error);
            Ok(run, "snp-ov-refs", HasAnyRef(tree), "无 ref=e");
            Ok(run, "snp-ov-id-short", NavSession.IsShortId(Str(ov, "channel_id")), Str(ov, "channel_id"));
            Ok(run, "snp-ov-kind", kind == "browser" && !string.IsNullOrEmpty(track), "kind=" + kind);
            Ok(run, "snp-ov-no-html",
                ov.Success
                && tree.IndexOf("<!doctype", StringComparison.OrdinalIgnoreCase) < 0
                && tree.IndexOf("<html", StringComparison.OrdinalIgnoreCase) < 0,
                "树像 HTML");
            string dump = tree + string.Join(",", Data(ov) == null ? new string[0] : new List<string>(Data(ov).Keys).ToArray());
            Ok(run, "snp-ov-no-backend",
                dump.IndexOf("backendNodeId", StringComparison.Ordinal) < 0
                && dump.IndexOf("frameId", StringComparison.Ordinal) < 0
                && dump.IndexOf("cssPath", StringComparison.Ordinal) < 0,
                "泄漏内部 id");
            Ok(run, "snp-ov-url-title",
                ov.Success && Uri.TryCreate(Str(ov, "url"), UriKind.Absolute, out _),
                Str(ov, "url"));
            string dom = Str(ov, "dom_supplement");
            Ok(run, "snp-ov-dom-off",
                tree.IndexOf("(DOM补全)", StringComparison.Ordinal) < 0
                && (string.IsNullOrEmpty(dom) || dom == "off"),
                "dom=" + dom);

            Ok(run, "snp-ifr-shell",
                (tree.Contains("外壳") || tree.Contains("根页")) && tree.IndexOf("我的得分", StringComparison.Ordinal) < 0,
                "壳/子页字不符");
            string workRef = IframeRef(tree, "教师工作台");
            Ok(run, "snp-ifr-row", !string.IsNullOrEmpty(workRef), "无工作台 iframe ref");
            Ok(run, "snp-ifr-hidden-row", HasIframeHint(tree, "ad") || HasIframeHint(tree, "hidden") || HasAnyIframeRef(tree),
                "没有隐藏/广告 iframe 行");

            if (!string.IsNullOrEmpty(workRef))
            {
                ToolResult entered = s.Snap(id, workRef);
                string et = Tree(entered);
                Ok(run, "snp-ifr-enter",
                    entered.Success && (et.Contains("我的得分") || et.Contains("评教结果")),
                    entered.Success ? "进了但无子页字" : entered.Error);
            }
            else
            {
                run.CaseFail("snp-ifr-enter", "无 iframe ref");
            }

            string leaf = FirstNonIframeRef(tree);
            if (string.IsNullOrEmpty(leaf))
            {
                FailMany(run, "壳上无非 iframe eN",
                    "snp-dt-mode", "snp-dt-echo-ref", "snp-dt-subtree", "snp-dt-ignore-dom");
            }
            else
            {
                ToolResult dt = s.Snap(id, leaf);
                Ok(run, "snp-dt-mode", dt.Success && Str(dt, "mode") == "detail", dt.Error ?? Str(dt, "mode"));
                Ok(run, "snp-dt-echo-ref", Str(dt, "ref") == leaf, Str(dt, "ref"));
                Ok(run, "snp-dt-subtree", Tree(dt).Length > 0, "空树");
                ToolResult ign = s.Snap(id, leaf, "input");
                Ok(run, "snp-dt-ignore-dom", ign.Success, ign.Error);
            }

            ToolResult again = s.Snap(id);
            Ok(run, "snp-dt-rebuild", again.Success && Str(again, "mode") == "overview", again.Error);

            RunHidden(run, s, http);
            RunDomInput(run, s, http, navHost);
            RunNested(run, s, http, navHost);
            RunSandbox(run, s, http, navHost);
        }

        private static void RunHidden(TestRun run, NavSession s, FixtureHttp http)
        {
            ToolResult n = s.Nav(http.Url("iframe", "outer.html"), host: "yiwrite", visible: false);
            if (!n.Success)
            {
                run.CaseFail("snp-ov-hidden", n.Error ?? "打开失败");
                return;
            }

            ToolResult ov = s.Snap(Str(n, "channel_id"));
            Ok(run, "snp-ov-hidden", ov.Success, ov.Error);
        }

        private static void RunDomInput(TestRun run, NavSession s, FixtureHttp http, string navHost)
        {
            string id = OpenOk(run, s, http.Url("native-select", "index.html"), navHost, "snp-ov-dom-input");
            if (id == null)
            {
                return;
            }

            ToolResult r = s.Snap(id, domSupplement: "input");
            string t = Tree(r);
            Ok(run, "snp-ov-dom-input",
                r.Success && (t.Contains("论文题目") || t.Contains("(DOM补全)") || t.Contains("lwtm")),
                r.Success ? "无输入痕迹" : r.Error);
        }

        private static void RunNested(TestRun run, NavSession s, FixtureHttp http, string navHost)
        {
            string id = OpenOk(run, s, http.Url("iframe", "nested.html"), navHost, "snp-ifr-nested");
            if (id == null)
            {
                return;
            }

            ToolResult ov = s.Snap(id);
            string t0 = Tree(ov);
            if (t0.Contains("最里层按钮"))
            {
                run.CaseFail("snp-ifr-nested", "空拍就看见最里层");
                return;
            }

            string fr = IframeRef(t0, null);
            if (string.IsNullOrEmpty(fr))
            {
                run.CaseFail("snp-ifr-nested", "无 iframe 行");
                return;
            }

            ToolResult mid = s.Snap(id, fr);
            string t1 = Tree(mid);
            Ok(run, "snp-ifr-nested",
                mid.Success && (t1.Contains("应用层") || t1.Contains("成绩报表") || HasAnyIframeRef(t1))
                && t1.IndexOf("最里层按钮", StringComparison.Ordinal) < 0,
                mid.Success ? "一层后字不符或已见最里层" : mid.Error);
        }

        private static void RunSandbox(TestRun run, NavSession s, FixtureHttp http, string navHost)
        {
            string id = OpenOk(run, s, http.Url("iframe", "sandbox-xorigin.html"), navHost, "snp-ifr-sandbox");
            if (id == null)
            {
                return;
            }

            ToolResult ov = s.Snap(id);
            string fr = IframeRef(Tree(ov), "sandbox");
            if (string.IsNullOrEmpty(fr))
            {
                fr = IframeRef(Tree(ov), null);
            }

            if (string.IsNullOrEmpty(fr))
            {
                run.CaseFail("snp-ifr-sandbox", "无 iframe 行");
                return;
            }

            ToolResult ex = s.Snap(id, fr);
            string inner = Tree(ex);
            if (ex.Success)
            {
                Ok(run, "snp-ifr-sandbox", LooksLikeExampleCom(inner), "展开成功但回的是外壳树");
                return;
            }

            ExpectFail(run, "snp-ifr-sandbox", ex, "无法读取");
        }

        private static void RunRealSlim(TestRun run, NavSession s, FixtureHttp http, string navHost)
        {
            SnapHas(run, s, http.Url("public-example", "index.html"), navHost, "snp-real-example",
                t => t.Contains("Example Domain") && HasTextRef(t, "More information"));

            string tabs = OpenOk(run, s, http.Url("center-tabs", "index.html"), navHost, "snp-real-haima-tab");
            if (tabs == null)
            {
                FailMany(run, "依赖 center-tabs",
                    "snp-real-haima-tab", "snp-real-haima-side", "snp-real-haima-longcard");
            }
            else
            {
                string t = Tree(s.Snap(tabs));
                Ok(run, "snp-real-haima-tab", HasTextRef(t, "项目成员"), "无项目成员 ref");
                Ok(run, "snp-real-haima-side", HasTextRef(t, "评教列表"), "无评教列表 ref");
                Ok(run, "snp-real-haima-longcard", !IsShortButton(t, "这是一张很长的可点卡片说明文字"),
                    "长卡片收成短按钮");
            }

            string card = OpenOk(run, s, http.Url("hover-card", "index.html"), navHost, "snp-real-pj-card");
            if (card == null)
            {
                run.CaseFail("snp-real-pj-detail", "依赖 hover-card");
            }
            else
            {
                string t = Tree(s.Snap(card));
                Ok(run, "snp-real-pj-card", HasTextRef(t, "机器学习课程设计") || t.Contains("机器学习课程设计"),
                    "无课程卡");
                Ok(run, "snp-real-pj-detail", HasTextRef(t, "查看详情"), "无查看详情 ref");
            }

            string sel = OpenOk(run, s, http.Url("native-select", "index.html"), navHost, "snp-real-select");
            if (sel != null)
            {
                string t = Tree(s.Snap(sel));
                bool good = t.Contains("第一作者类型") || t.Contains("本校教师") || t.Contains("options=");
                bool badName = Regex.IsMatch(t, @"textbox[^\n]*不能为空[^\n]*ref=e|不能为空[^\n]*ref=e")
                    && !t.Contains("论文题目");
                Ok(run, "snp-real-select", good && !badName, good ? "校验文案当了 name" : "无下拉字");
            }

            SnapHas(run, s, http.Url("overlay-cookie", "index.html"), navHost, "snp-real-cookie",
                t => HasTextRef(t, "查看结果") && HasTextRef(t, "接受"));
        }

        private static void RunRealSaved(TestRun run, NavSession s, FixtureHttp http, string navHost)
        {
            string train = OpenOk(run, s, http.Real("train"), navHost, "snp-real-train-member");
            if (train == null)
            {
                FailMany(run, "依赖 项目实战去实训.html",
                    "snp-real-train-member", "snp-real-train-detail-tab", "snp-real-train-assign",
                    "snp-real-haima-menu");
            }
            else
            {
                string t = Tree(s.Snap(train));
                Ok(run, "snp-real-train-member", HasTextRef(t, "项目成员"), "无项目成员 ref");
                Ok(run, "snp-real-train-detail-tab", HasTextRef(t, "项目详情"), "无项目详情 ref");
                Ok(run, "snp-real-train-assign",
                    HasTextRef(t, "项目成员分配")
                    || (!t.Contains("项目成员分配") && HasTextRef(t, "项目成员")),
                    "分配按钮既无 ref、成员 Tab 也无 ref");
                Ok(run, "snp-real-haima-menu", HasTextRef(t, "过程管理") || HasTextRef(t, "首页"),
                    "无侧栏短项");
            }

            SnapHas(run, s, http.Real("square"), navHost, "snp-real-square",
                t => (HasTextRef(t, "项目审核") || HasTextRef(t, "刷新")) && t.Contains("项目广场"));
            SnapHas(run, s, http.Real("type"), navHost, "snp-real-type",
                t => HasTextRef(t, "查询") || HasTextRef(t, "项目类别") || t.Contains("项目类别"));
            SnapHas(run, s, http.Real("jd-home"), navHost, "snp-real-jd-cart",
                t => HasTextRef(t, "购物车") || t.Contains("购物车"));
            SnapHas(run, s, http.Real("jd-mine"), navHost, "snp-real-jd-mine",
                t => HasTextRef(t, "我的京东") || t.Contains("我的京东"));
        }

        private static void RunNoQuit(TestRun run, int chromeBefore, int edgeBefore)
        {
            int chromeNow = CountProc("chrome");
            int edgeNow = CountProc("msedge");
            Ok(run, "snp-no-quit", chromeNow >= chromeBefore && edgeNow >= edgeBefore,
                "chrome " + chromeBefore + "→" + chromeNow + " edge " + edgeBefore + "→" + edgeNow);
        }

        private static void SnapHas(
            TestRun run,
            NavSession s,
            string url,
            string navHost,
            string id,
            Func<string, bool> pred)
        {
            string ch = OpenOk(run, s, url, navHost, id);
            if (ch == null)
            {
                return;
            }

            ToolResult r = s.Snap(ch);
            Ok(run, id, r.Success && pred(Tree(r)), r.Success ? "树未命中" : r.Error);
        }

        private static string OpenOk(TestRun run, NavSession s, string url, string host, string failId)
        {
            ToolResult r = s.Nav(url, host: host);
            if (r.Success && NavSession.IsShortId(Str(r, "channel_id")))
            {
                return Str(r, "channel_id");
            }

            run.CaseFail(failId, r.Success ? "无短 channel_id" : (r.Error ?? "打开失败"));
            return null;
        }

        private static bool LooksLikeExampleCom(string tree)
        {
            if (string.IsNullOrEmpty(tree))
            {
                return false;
            }

            return tree.IndexOf("Example Domain", StringComparison.Ordinal) >= 0
                || tree.IndexOf("More information", StringComparison.Ordinal) >= 0
                || tree.IndexOf("documentation examples", StringComparison.OrdinalIgnoreCase) >= 0
                || tree.IndexOf("This domain is for use", StringComparison.Ordinal) >= 0;
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

        private static string Tree(ToolResult r)
        {
            return Str(r, "snapshot");
        }

        private static bool HasAnyRef(string tree)
        {
            return Regex.IsMatch(tree ?? "", @"ref=e\d+");
        }

        private static bool HasTextRef(string tree, string text)
        {
            if (string.IsNullOrEmpty(tree) || string.IsNullOrEmpty(text))
            {
                return false;
            }

            foreach (string line in tree.Split('\n'))
            {
                if (line.IndexOf(text, StringComparison.Ordinal) >= 0
                    && Regex.IsMatch(line, @"ref=e\d+"))
                {
                    return true;
                }
            }

            return false;
        }

        private static bool IsShortButton(string tree, string longText)
        {
            if (string.IsNullOrEmpty(tree) || string.IsNullOrEmpty(longText))
            {
                return false;
            }

            string head = longText.Length > 16 ? longText.Substring(0, 16) : longText;
            foreach (string line in tree.Split('\n'))
            {
                if (line.IndexOf(head, StringComparison.Ordinal) < 0 || !Regex.IsMatch(line, @"ref=e\d+"))
                {
                    continue;
                }

                if (line.IndexOf(longText, StringComparison.Ordinal) < 0 && line.Length < 80)
                {
                    return true;
                }
            }

            return false;
        }

        private static string FirstRef(string tree)
        {
            Match m = Regex.Match(tree ?? "", @"ref=(e\d+)");
            return m.Success ? m.Groups[1].Value : "";
        }

        private static string FirstNonIframeRef(string tree)
        {
            foreach (string line in (tree ?? "").Split('\n'))
            {
                if (Regex.IsMatch(line, @"iframe|frame", RegexOptions.IgnoreCase))
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

        private static string IframeRef(string tree, string hint)
        {
            string first = "";
            foreach (string line in (tree ?? "").Split('\n'))
            {
                if (!Regex.IsMatch(line, @"iframe|IFrame|frame", RegexOptions.IgnoreCase))
                {
                    continue;
                }

                Match m = Regex.Match(line, @"ref=(e\d+)");
                if (!m.Success)
                {
                    continue;
                }

                if (!string.IsNullOrEmpty(hint) && line.IndexOf(hint, StringComparison.Ordinal) >= 0)
                {
                    return m.Groups[1].Value;
                }

                if (first.Length == 0)
                {
                    first = m.Groups[1].Value;
                }
            }

            return first;
        }

        private static bool HasAnyIframeRef(string tree)
        {
            return !string.IsNullOrEmpty(IframeRef(tree, null));
        }

        private static bool HasIframeHint(string tree, string hint)
        {
            foreach (string line in (tree ?? "").Split('\n'))
            {
                if (line.IndexOf(hint, StringComparison.OrdinalIgnoreCase) >= 0
                    && Regex.IsMatch(line, @"iframe|frame", RegexOptions.IgnoreCase)
                    && Regex.IsMatch(line, @"ref=e\d+"))
                {
                    return true;
                }
            }

            return false;
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
