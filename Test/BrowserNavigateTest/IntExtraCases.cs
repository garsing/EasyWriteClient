using System;
using System.Globalization;
using WordAddIn1;

namespace BrowserNavigateTest
{
    internal static partial class IntCases
    {
        internal static void RunExtra(TestRun run, NavSession s, FixtureHttp http)
        {
            RunExtraDate(run, s, http);
            RunExtraInput(run, s, http);
            RunExtraCheckSelect(run, s, http);
            RunExtraPressScroll(run, s, http);
            RunExtraD18Neg(run, s, http);
            RunExtraTreeIframe(run, s, http);
        }

        private static void RunExtraDate(TestRun run, NavSession s, FixtureHttp http)
        {
            string id = OpenOk(run, s, http.Url("form-controls", "index.html"), "int-type-date-tomorrow");
            if (id == null)
            {
                FailMany(run, "依赖 form-controls", ExtraDateIdsExceptFirst());
                return;
            }

            string tree = WaitTree(s, id, true, "立项日");
            DateTime today = DateTime.Now.Date;
            TypeValue(run, s, id, ref tree, "int-type-date-tomorrow", DateHint(tree), "明天", today.AddDays(1).ToString("yyyy-MM-dd"));
            tree = WaitTree(s, id, true, "立项日");
            TypeValue(run, s, id, ref tree, "int-type-date-tomorrow-en", DateHint(tree), "tomorrow", today.AddDays(1).ToString("yyyy-MM-dd"));
            tree = WaitTree(s, id, true, "立项日");
            TypeValue(run, s, id, ref tree, "int-type-date-today-en", DateHint(tree), "today", today.ToString("yyyy-MM-dd"));
            tree = WaitTree(s, id, true, "立项日");
            TypeValue(run, s, id, ref tree, "int-type-date-yesterday-en", DateHint(tree), "yesterday", today.AddDays(-1).ToString("yyyy-MM-dd"));
            tree = WaitTree(s, id, true, "立项日");
            TypeValue(run, s, id, ref tree, "int-type-date-slash", DateHint(tree), "2026/10/3", "2026-10-03");
            tree = WaitTree(s, id, true, "立项日");
            TypeValue(run, s, id, ref tree, "int-type-date-cn", DateHint(tree), "2026年10月3日", "2026-10-03");
            tree = WaitTree(s, id, true, "立项日");
            TypeValue(run, s, id, ref tree, "int-type-date-compact", DateHint(tree), "20261003", "2026-10-03");
            tree = WaitTree(s, id, true, "立项日");
            TypeValue(run, s, id, ref tree, "int-type-date-again", DateHint(tree), "2026-11-01", "2026-11-01");

            ExpectFail(
                run,
                "int-neg-type-empty",
                s.Interact(id, "type", DateHint(WaitTree(s, id, true, "立项日")), text: ""),
                "text");

            tree = WaitTree(s, id, true, "答辩时刻");
            TypeValue(run, s, id, ref tree, "int-type-datetime-space", FindRef(tree, "答辩时刻"), "2026-10-03 14:30", "2026-10-03T14:30");
            tree = WaitTree(s, id, true, "答辩时刻");
            TypeValue(run, s, id, ref tree, "int-type-datetime-today", FindRef(tree, "答辩时刻"), "今天", today.ToString("yyyy-MM-dd") + "T00:00");

            tree = WaitTree(s, id, true, "所属月");
            TypeValue(run, s, id, ref tree, "int-type-month-today", FindRef(tree, "所属月"), "今天", today.ToString("yyyy-MM"));
            tree = WaitTree(s, id, true, "所属月");
            TypeValue(run, s, id, ref tree, "int-type-month-slash", FindRef(tree, "所属月"), "2026/10", "2026-10");
            tree = WaitTree(s, id, true, "所属月");
            TypeValue(run, s, id, ref tree, "int-type-month-tomorrow", FindRef(tree, "所属月"), "明天", today.AddDays(1).ToString("yyyy-MM"));

            tree = WaitTree(s, id, true, "教学周");
            TypeValue(run, s, id, ref tree, "int-type-week-today", FindRef(tree, "教学周"), "今天", IsoWeek(today));

            tree = WaitTree(s, id, true, "开始时刻");
            TypeValue(run, s, id, ref tree, "int-type-time-sec", FindRef(tree, "开始时刻"), "14:30:00", "14:30:00");
            tree = WaitTree(s, id, true, "开始时刻");
            TypeValue(run, s, id, ref tree, "int-type-time-short", FindRef(tree, "开始时刻"), "9:05", "09:05");

            ExpectFmtFail(run, s, id, "int-type-time-bad", "开始时刻", "不是时间");
            ExpectFmtFail(run, s, id, "int-type-month-bad", "所属月", "不是月份");
            ExpectFmtFail(run, s, id, "int-type-week-bad", "教学周", "不是周");
            ExpectFmtFail(run, s, id, "int-type-datetime-bad", "答辩时刻", "不是日期时间");
        }

        private static string[] ExtraDateIdsExceptFirst()
        {
            return new[]
            {
                "int-type-date-tomorrow-en", "int-type-date-today-en", "int-type-date-yesterday-en", "int-type-date-slash",
                "int-type-date-cn", "int-type-date-compact", "int-type-date-again",
                "int-neg-type-empty", "int-type-datetime-space", "int-type-datetime-today",
                "int-type-month-today", "int-type-month-slash", "int-type-month-tomorrow",
                "int-type-week-today", "int-type-time-sec", "int-type-time-short",
                "int-type-time-bad", "int-type-month-bad", "int-type-week-bad",
                "int-type-datetime-bad"
            };
        }

        private static void ExpectFmtFail(TestRun run, NavSession s, string id, string caseId, string hint, string text)
        {
            string tree = WaitTree(s, id, true, hint);
            string r = FindRef(tree, hint);
            ToolResult got = string.IsNullOrEmpty(r) ? null : s.Interact(id, "type", r, text: text);
            Ok(
                run,
                caseId,
                got != null && !got.Success && ContainsAny(got.Error, "日期", "时间", "格式"),
                r == null ? "无 " + hint : (got != null && got.Success ? "乱填却成功" : got?.Error));
        }

        private static void RunExtraInput(TestRun run, NavSession s, FixtureHttp http)
        {
            string id = OpenOk(run, s, http.Url("form-extra", "index.html"), "int-type-email");
            if (id == null)
            {
                FailMany(run, "依赖 form-extra",
                    "int-type-tel", "int-type-url", "int-type-search", "int-type-datalist",
                    "int-type-overwrite", "int-type-chinese", "int-type-spaces", "int-type-long",
                    "int-type-twice", "int-type-email-hidden", "int-type-search-enter",
                    "int-type-on-button", "int-hover-plain", "int-select-multi");
            }
            else
            {
                string tree = WaitTree(s, id, true, "邮箱", "检索", "课程名");
                TypeValue(run, s, id, ref tree, "int-type-email", FindRef(tree, "邮箱"), "a@b.com", "a@b.com");
                tree = WaitTree(s, id, true, "电话");
                TypeValue(run, s, id, ref tree, "int-type-tel", FindRef(tree, "电话"), "13800138000", "13800138000");
                tree = WaitTree(s, id, true, "网址");
                TypeValue(run, s, id, ref tree, "int-type-url", FindRef(tree, "网址"), "https://example.com", "https://example.com");
                tree = WaitTree(s, id, true, "检索");
                TypeOk(run, s, id, "int-type-search", FindRef(tree, "检索"), "检索词");
                tree = WaitTree(s, id, true, "课程名");
                TypeValue(run, s, id, ref tree, "int-type-datalist", FindRef(tree, "课程名"), "机器学习", "机器学习");
                tree = WaitTree(s, id, true, "邮箱");
                TypeValue(run, s, id, ref tree, "int-type-overwrite", FindRef(tree, "邮箱"), "c@d.com", "c@d.com");
                tree = WaitTree(s, id, true, "检索");
                TypeOk(run, s, id, "int-type-chinese", FindRef(WaitTree(s, id, true, "检索"), "检索"), "中文词");
                TypeOk(run, s, id, "int-type-spaces", FindRef(WaitTree(s, id, true, "检索"), "检索"), "含空 格");
                TypeOk(run, s, id, "int-type-long", FindRef(WaitTree(s, id, true, "检索"), "检索"), new string('字', 40));
                tree = WaitTree(s, id, true, "邮箱");
                string em = FindRef(tree, "邮箱");
                ToolResult a = string.IsNullOrEmpty(em) ? null : s.Interact(id, "type", em, text: "one@x.com");
                string em2 = a != null && a.Success ? (FindRef(Tree(a), "邮箱") ?? em) : em;
                ToolResult b = string.IsNullOrEmpty(em2) ? null : s.Interact(id, "type", em2, text: "two@x.com");
                Ok(run, "int-type-twice", b != null && b.Success && Tree(b).Contains("two@x.com"), b == null ? "无邮箱" : b.Error);

                string hid = OpenOk(run, s, http.Url("form-extra", "index.html"), "int-type-email-hidden", visible: false);
                if (hid == null)
                {
                    run.CaseFail("int-type-search-enter", "依赖隐藏窗");
                }
                else
                {
                    string ht = WaitTree(s, hid, true, "邮箱", "检索");
                    TypeValue(run, s, hid, ref ht, "int-type-email-hidden", FindRef(ht, "邮箱"), "h@h.com", "h@h.com");
                    string se = FindRef(WaitTree(s, hid, true, "检索"), "检索");
                    ToolResult typed = string.IsNullOrEmpty(se) ? null : s.Interact(hid, "type", se, text: "隐搜");
                    string focus = typed != null && typed.Success ? (FindRef(Tree(typed), "检索") ?? se) : se;
                    ToolResult ent = s.Interact(hid, "press", focus, key: "Enter");
                    Ok(run, "int-type-search-enter", typed != null && typed.Success && ent.Success, ent.Error ?? typed?.Error);
                }

                string btn = FindRef(WaitTree(s, id, false, "巩固按钮"), "巩固按钮");
                ToolResult onBtn = s.Interact(id, "click", btn ?? "e1");
                Ok(run, "int-type-on-button", onBtn.Success && Tree(onBtn).Contains("巩固按钮已点"), onBtn.Error);
                ToolResult hv = s.Interact(id, "hover", btn ?? "e1");
                Ok(run, "int-hover-plain", hv.Success && Flag(hv, "resnapshot"), hv.Error);

                string multi = FindRef(WaitTree(s, id, true, "多选类别"), "多选类别", "listbox")
                    ?? FindRef(WaitTree(s, id, true, "多选类别"), "多选类别");
                ToolResult ms = string.IsNullOrEmpty(multi) ? null : s.Interact(id, "select", multi, option: "甲类");
                Ok(run, "int-select-multi", ms != null && ms.Success && Tree(ms).Contains("甲类"), multi == null ? "无多选" : (ms?.Error ?? "状态未变"));
            }

            string form = OpenOk(run, s, http.Url("form-controls", "index.html"), "int-type-textarea-again");
            if (form == null)
            {
                FailMany(run, "依赖 form-controls",
                    "int-type-ce-again", "int-type-number-dec", "int-type-number-neg",
                    "int-type-on-checkbox", "int-type-ta-hidden");
            }
            else
            {
                string tree = WaitTree(s, form, true, "多行说明", "份数");
                TypeValue(run, s, form, ref tree, "int-type-textarea-again", FindRef(tree, "多行说明"), "第二稿", "第二稿");
                tree = WaitTree(s, form, true, "可编辑摘要");
                TypeValue(run, s, form, ref tree, "int-type-ce-again", FindRef(tree, "可编辑摘要"), "再写摘要", "再写摘要");
                tree = WaitTree(s, form, true, "份数");
                TypeValue(run, s, form, ref tree, "int-type-number-dec", FindRef(tree, "份数"), "3.5", "3.5");
                tree = WaitTree(s, form, true, "份数");
                TypeValue(run, s, form, ref tree, "int-type-number-neg", FindRef(tree, "份数"), "-2", "-2");
                string ck = FindRef(WaitTree(s, form, true, "同意条款"), "同意条款");
                ToolResult onCk = s.Interact(form, "click", ck ?? "e1");
                Ok(run, "int-type-on-checkbox", onCk.Success && ContainsAny(Tree(onCk), "已勾选", "已取消勾选"), onCk.Error);

                string hid = OpenOk(run, s, http.Url("form-controls", "index.html"), "int-type-ta-hidden", visible: false);
                if (hid != null)
                {
                    string ht = WaitTree(s, hid, true, "多行说明");
                    TypeValue(run, s, hid, ref ht, "int-type-ta-hidden", FindRef(ht, "多行说明"), "隐多行", "隐多行");
                }
            }

            string okb = OpenOk(run, s, http.Url("disabled-input", "index.html"), "int-type-ok-box");
            if (okb != null)
            {
                string t = WaitTree(s, okb, true, "可填框");
                TypeValue(run, s, okb, ref t, "int-type-ok-box", FindRef(t, "可填框"), "可填了", "可填了");
            }
        }

        private static void RunExtraCheckSelect(TestRun run, NavSession s, FixtureHttp http)
        {
            string id = OpenOk(run, s, http.Url("form-controls", "index.html"), "int-click-checkbox-off");
            if (id == null)
            {
                FailMany(run, "依赖 form-controls",
                    "int-click-checkbox-twice", "int-click-radio-stu", "int-click-radio-then-tea",
                    "int-click-switch-off", "int-click-switch-twice", "int-click-label-check",
                    "int-click-radio-again", "int-click-ck-hidden");
            }
            else
            {
                string tree = WaitTree(s, id, true, "同意条款", "夜间模式", "学生");
                string ck = FindRef(tree, "同意条款");
                ToolResult on = string.IsNullOrEmpty(ck) ? null : s.Interact(id, "click", ck);
                string ck2 = on != null && on.Success ? (FindRef(Tree(on), "同意条款") ?? ck) : ck;
                ToolResult off = string.IsNullOrEmpty(ck2) ? null : s.Interact(id, "click", ck2);
                Ok(run, "int-click-checkbox-off", off != null && off.Success && Tree(off).Contains("已取消勾选"), off == null ? "无勾选" : (off.Error ?? "未取消"));

                tree = WaitTree(s, id, true, "同意条款");
                ck = FindRef(tree, "同意条款");
                ToolResult c1 = string.IsNullOrEmpty(ck) ? null : s.Interact(id, "click", ck);
                string n1 = c1 != null && c1.Success ? (FindRef(Tree(c1), "同意条款") ?? ck) : ck;
                ToolResult c2 = string.IsNullOrEmpty(n1) ? null : s.Interact(id, "click", n1);
                Ok(run, "int-click-checkbox-twice", c1 != null && c1.Success && c2 != null && c2.Success, c2?.Error ?? c1?.Error);

                tree = WaitTree(s, id, true, "学生");
                ClickStatus(run, s, id, ref tree, "int-click-radio-stu", FindRef(tree, "学生"), "身份=stu", "学生");
                tree = WaitTree(s, id, true, "教师");
                ClickStatus(run, s, id, ref tree, "int-click-radio-then-tea", FindRef(tree, "教师"), "身份=tea", "教师");
                tree = WaitTree(s, id, true, "教师");
                ClickStatus(run, s, id, ref tree, "int-click-radio-again", FindRef(tree, "教师"), "身份=tea", "教师");

                tree = WaitTree(s, id, true, "夜间模式");
                string sw = FindRef(tree, "夜间模式");
                ToolResult swOn = string.IsNullOrEmpty(sw) ? null : s.Interact(id, "click", sw);
                string sw2 = swOn != null && swOn.Success ? (FindRef(Tree(swOn), "夜间模式") ?? sw) : sw;
                ToolResult swOff = string.IsNullOrEmpty(sw2) ? null : s.Interact(id, "click", sw2);
                Ok(run, "int-click-switch-off", swOff != null && swOff.Success && Tree(swOff).Contains("夜间已关"), swOff == null ? "无开关" : (swOff.Error ?? "未关"));
                tree = WaitTree(s, id, true, "夜间模式");
                sw = FindRef(tree, "夜间模式");
                ToolResult s1 = string.IsNullOrEmpty(sw) ? null : s.Interact(id, "click", sw);
                string s2r = s1 != null && s1.Success ? (FindRef(Tree(s1), "夜间模式") ?? sw) : sw;
                ToolResult s2 = string.IsNullOrEmpty(s2r) ? null : s.Interact(id, "click", s2r);
                Ok(run, "int-click-switch-twice", s1 != null && s1.Success && s2 != null && s2.Success, s2?.Error ?? s1?.Error);

                tree = WaitTree(s, id, true, "同意条款");
                ClickStatus(run, s, id, ref tree, "int-click-label-check", FindRef(tree, "同意条款"), "已勾选", "已取消勾选");

                string hid = OpenOk(run, s, http.Url("form-controls", "index.html"), "int-click-ck-hidden", visible: false);
                if (hid != null)
                {
                    string ht = WaitTree(s, hid, true, "同意条款");
                    ClickStatus(run, s, hid, ref ht, "int-click-ck-hidden", FindRef(ht, "同意条款"), "已勾选");
                }
            }

            string sel = OpenOk(run, s, http.Url("native-select", "index.html"), "int-select-value");
            if (sel == null)
            {
                FailMany(run, "依赖 native-select",
                    "int-select-student", "int-select-contain", "int-select-science",
                    "int-select-first-missing", "int-select-hidden-stu", "int-click-select-cat",
                    "int-select-resnapshot");
            }
            else
            {
                string tree = WaitTree(s, sel, true, "第一作者类型", "统计类别");
                string first = FindRef(tree, "第一作者类型", "combobox") ?? FindRef(tree, "第一作者类型");
                string cat = FindRef(tree, "统计类别", "combobox") ?? FindRef(tree, "统计类别");
                ToolResult byVal = string.IsNullOrEmpty(first) ? null : s.Interact(sel, "select", first, option: "01");
                Ok(run, "int-select-value", byVal != null && byVal.Success && Tree(byVal).Contains("本校教师"), first == null ? "无下拉" : byVal?.Error);
                tree = WaitTree(s, sel, true, "第一作者类型");
                first = FindRef(tree, "第一作者类型", "combobox") ?? FindRef(tree, "第一作者类型");
                ToolResult stu = string.IsNullOrEmpty(first) ? null : s.Interact(sel, "select", first, option: "本校学生");
                Ok(run, "int-select-student", stu != null && stu.Success && Tree(stu).Contains("本校学生"), stu?.Error ?? "无下拉");
                tree = WaitTree(s, sel, true, "第一作者类型");
                first = FindRef(tree, "第一作者类型", "combobox") ?? FindRef(tree, "第一作者类型");
                ToolResult contain = string.IsNullOrEmpty(first) ? null : s.Interact(sel, "select", first, option: "教师");
                Ok(run, "int-select-contain", contain != null && contain.Success && Tree(contain).Contains("本校教师"), contain?.Error ?? "无下拉");
                tree = WaitTree(s, sel, true, "统计类别");
                cat = FindRef(tree, "统计类别", "combobox") ?? FindRef(tree, "统计类别");
                ToolResult sci = string.IsNullOrEmpty(cat) ? null : s.Interact(sel, "select", cat, option: "自然科学类");
                Ok(run, "int-select-science", sci != null && sci.Success, sci?.Error ?? "无统计类别");
                Ok(run, "int-select-resnapshot", sci != null && sci.Success && Flag(sci, "resnapshot") && !Flag(sci, "refs_invalidated"), "回树不符");
                tree = WaitTree(s, sel, true, "第一作者类型");
                first = FindRef(tree, "第一作者类型", "combobox") ?? FindRef(tree, "第一作者类型");
                ExpectFail(run, "int-select-first-missing", s.Interact(sel, "select", first ?? "e1", option: "应用研究"), "可选");
                ExpectFail(
                    run,
                    "int-click-select-cat",
                    s.Interact(sel, "click", cat ?? FindRef(WaitTree(s, sel, true, "统计类别"), "统计类别") ?? "e1"),
                    "select");

                string hid = OpenOk(run, s, http.Url("native-select", "index.html"), "int-select-hidden-stu", visible: false);
                if (hid != null)
                {
                    string ht = WaitTree(s, hid, true, "第一作者类型");
                    string href = FindRef(ht, "第一作者类型", "combobox") ?? FindRef(ht, "第一作者类型");
                    ToolResult hs = string.IsNullOrEmpty(href) ? null : s.Interact(hid, "select", href, option: "本校学生");
                    Ok(run, "int-select-hidden-stu", hs != null && hs.Success, hs == null ? "无 ref" : hs.Error);
                }
            }

            string dis = OpenOk(run, s, http.Url("disabled-input", "index.html"), "int-select-disabled");
            if (dis != null)
            {
                string t = WaitTree(s, dis, true, "禁用下拉");
                ExpectFail(run, "int-select-disabled", s.Interact(dis, "select", FindRef(t, "禁用下拉") ?? "e1", option: "甲"), "不可编辑");
            }
        }

        private static void RunExtraPressScroll(TestRun run, NavSession s, FixtureHttp http)
        {
            string id = OpenOk(run, s, http.Url("native-select", "index.html"), "int-press-escape");
            if (id == null)
            {
                FailMany(run, "依赖 native-select",
                    "int-press-esc-alias", "int-press-arrow-down", "int-press-arrow-up", "int-press-arrow-left",
                    "int-press-arrow-right", "int-press-backspace", "int-press-delete-key",
                    "int-press-home", "int-press-end", "int-press-page-down", "int-press-page-up",
                    "int-press-enter-noref");
            }
            else
            {
                Ok(run, "int-press-escape", s.Interact(id, "press", key: "Escape").Success, "Escape");
                Ok(run, "int-press-esc-alias", s.Interact(id, "press", key: "Esc").Success, "Esc");
                Ok(run, "int-press-arrow-down", s.Interact(id, "press", key: "ArrowDown").Success, "ArrowDown");
                Ok(run, "int-press-arrow-up", s.Interact(id, "press", key: "ArrowUp").Success, "ArrowUp");
                Ok(run, "int-press-arrow-left", s.Interact(id, "press", key: "ArrowLeft").Success, "ArrowLeft");
                Ok(run, "int-press-arrow-right", s.Interact(id, "press", key: "ArrowRight").Success, "ArrowRight");
                Ok(run, "int-press-backspace", s.Interact(id, "press", key: "Backspace").Success, "Backspace");
                Ok(run, "int-press-delete-key", s.Interact(id, "press", key: "Delete").Success, "Delete键");
                Ok(run, "int-press-home", s.Interact(id, "press", key: "Home").Success, "Home");
                Ok(run, "int-press-end", s.Interact(id, "press", key: "End").Success, "End");
                Ok(run, "int-press-page-down", s.Interact(id, "press", key: "PageDown").Success, "PageDown");
                Ok(run, "int-press-page-up", s.Interact(id, "press", key: "PageUp").Success, "PageUp");
                Ok(run, "int-press-enter-noref", s.Interact(id, "press", key: "Enter").Success, "Enter无ref");
            }

            string sc = OpenOk(run, s, http.Url("overflow-scroll", "index.html"), "int-scroll-up");
            if (sc == null)
            {
                FailMany(run, "依赖 overflow-scroll", "int-scroll-left", "int-scroll-right");
            }
            else
            {
                Ok(run, "int-scroll-up", s.Interact(sc, "scroll", direction: "up").Success, "up");
                Ok(run, "int-scroll-left", s.Interact(sc, "scroll", direction: "left").Success, "left");
                Ok(run, "int-scroll-right", s.Interact(sc, "scroll", direction: "right").Success, "right");
            }

            string card = OpenOk(run, s, http.Url("hover-card", "index.html"), "int-hover-then-click");
            if (card != null)
            {
                string t = WaitTree(s, card, false, "机器学习课程设计");
                string href = FindRef(t, "机器学习课程设计", "image")
                    ?? FindRef(t, "机器学习课程设计");
                ToolResult hv = string.IsNullOrEmpty(href) ? null : s.Interact(card, "hover", href);
                string neu = hv != null && hv.Success
                    ? (FindRef(Tree(hv), "机器学习课程设计", "image")
                        ?? FindRef(Tree(hv), "机器学习课程设计")
                        ?? href)
                    : href;
                ToolResult clk = string.IsNullOrEmpty(neu) ? null : s.Interact(card, "click", neu);
                Ok(
                    run,
                    "int-hover-then-click",
                    hv != null && hv.Success && clk != null && clk.Success
                    && (Tree(clk).Contains("已进入详情") || Str(clk, "url").Contains("detail")),
                    clk == null ? "无卡片" : (clk.Error ?? hv?.Error));
                Ok(run, "int-hover-resnapshot", hv != null && hv.Success && Flag(hv, "resnapshot"), "hover 回树");
            }
            else
            {
                run.CaseFail("int-hover-resnapshot", "依赖 hover-card");
            }
        }

        private static void RunExtraD18Neg(TestRun run, NavSession s, FixtureHttp http)
        {
            string id = OpenOk(run, s, http.Url("d18-risk", "index.html"), "int-d18-click-submit");
            if (id == null)
            {
                FailMany(run, "依赖 d18-risk",
                    "int-d18-click-pay2", "int-d18-click-signin", "int-d18-click-delete-en",
                    "int-d18-click-safe", "int-d18-press-enter-login-btn");
            }
            else
            {
                string tree = WaitTree(s, id, true, "提交", "付款", "安全按钮");
                ExpectFailAndUnchanged(run, s, id, "int-d18-click-pay2", FindRef(tree, "付款"), "click", null, "拒绝");
                ExpectFailAndUnchanged(run, s, id, "int-d18-click-signin", FindRef(tree, "Sign in"), "click", null, "拒绝");
                ExpectFailAndUnchanged(run, s, id, "int-d18-click-delete-en", FindRef(tree, "Delete"), "click", null, "拒绝");
                string submit = FindRef(tree, "提交");
                ToolResult submitted = string.IsNullOrEmpty(submit) ? null : s.Interact(id, "click", submit);
                Ok(
                    run,
                    "int-d18-click-submit",
                    submitted != null && submitted.Success && Tree(submitted).Contains("提交已点"),
                    submit == null ? "无提交 ref" : (submitted == null ? "无回包" : submitted.Error));
                tree = submitted != null && submitted.Success ? Tree(submitted) : tree;
                string safe = FindRef(tree, "安全按钮");
                ToolResult ok = string.IsNullOrEmpty(safe) ? null : s.Interact(id, "click", safe);
                Ok(run, "int-d18-click-safe", ok != null && ok.Success && Tree(ok).Contains("安全按钮已点"), ok == null ? "无安全按钮" : ok.Error);
                string login = FindRef(WaitTree(s, id, false, "登录"), "登录", "button") ?? FindRef(WaitTree(s, id, false, "登录"), "登录");
                ExpectFail(run, "int-d18-press-enter-login-btn", s.Interact(id, "press", login, key: "Enter"), "拒绝");
            }

            string ov = OpenOk(run, s, http.Url("overlay-cookie", "index.html"), "int-neg-action-upper");
            if (ov == null)
            {
                FailMany(run, "依赖 overlay-cookie",
                    "int-neg-hover-no-ref", "int-neg-type-stale", "int-click-accept-cookie",
                    "int-click-cookie-then-under", "int-click-plain-twice", "int-type-resnapshot");
            }
            else
            {
                string tree = WaitTree(s, ov, false, "普通按钮");
                string plain = FindRef(tree, "普通按钮");
                ToolResult up = s.Interact(ov, "CLICK", plain);
                Ok(run, "int-neg-action-upper", up.Success, up.Error);
                ExpectFail(run, "int-neg-hover-no-ref", s.Interact(ov, "hover"), "ref");

                string staleId = OpenOk(run, s, http.Url("overlay-cookie", "index.html"), "int-neg-type-stale");
                if (staleId != null)
                {
                    string oldBox = FindRef(WaitTree(s, staleId, false, "普通按钮"), "普通按钮");
                    s.Nav(http.Url("public-example", "index.html"), host: NavHost, channelId: staleId);
                    ToolResult stale = s.Interact(staleId, "type", oldBox, text: "x");
                    Ok(run, "int-neg-type-stale", stale != null && !stale.Success, stale != null && stale.Success ? "跳转后旧 eN 仍能 type" : stale?.Error);
                }

                string ck = OpenOk(run, s, http.Url("overlay-cookie", "index.html"), "int-click-accept-cookie");
                if (ck != null)
                {
                    string ct = WaitTree(s, ck, false, "接受", "查看结果");
                    string acc = FindRef(ct, "接受");
                    ToolResult accR = string.IsNullOrEmpty(acc) ? null : s.Interact(ck, "click", acc);
                    Ok(run, "int-click-accept-cookie", accR != null && accR.Success && Tree(accR).Contains("已接受 Cookie"), accR == null ? "无接受" : accR.Error);
                    string under = FindRef(accR != null && accR.Success ? Tree(accR) : ct, "查看结果");
                    ToolResult und = string.IsNullOrEmpty(under) ? null : s.Interact(ck, "click", under);
                    Ok(run, "int-click-cookie-then-under", und != null && und.Success && Tree(und).Contains("查看结果已点"), und == null ? "无查看结果" : und.Error);
                }

                string tw = OpenOk(run, s, http.Url("overlay-cookie", "index.html"), "int-click-plain-twice");
                if (tw != null)
                {
                    string tt = WaitTree(s, tw, false, "普通按钮");
                    string p1 = FindRef(tt, "普通按钮");
                    ToolResult x1 = string.IsNullOrEmpty(p1) ? null : s.Interact(tw, "click", p1);
                    string p2 = x1 != null && x1.Success ? FindRef(Tree(x1), "普通按钮") : null;
                    ToolResult x2 = string.IsNullOrEmpty(p2) ? null : s.Interact(tw, "click", p2);
                    Ok(run, "int-click-plain-twice", x1 != null && x1.Success && x2 != null && x2.Success, x2?.Error ?? x1?.Error);
                }

                string ns = OpenOk(run, s, http.Url("native-select", "index.html"), "int-type-resnapshot");
                if (ns != null)
                {
                    string nt = WaitTree(s, ns, true, "论文题目");
                    ToolResult typed = s.Interact(ns, "type", FindRef(nt, "论文题目"), text: "回树词");
                    Ok(run, "int-type-resnapshot", typed.Success && Flag(typed, "resnapshot") && !Flag(typed, "refs_invalidated") && !HasLeak(typed), typed.Error);
                }
            }

            string ifr = OpenOk(run, s, http.Url("iframe", "outer.html"), "int-neg-scroll-iframe");
            if (ifr == null)
            {
                FailMany(run, "依赖 outer", "int-neg-press-iframe", "int-neg-select-iframe");
            }
            else
            {
                string frameRef = FindRef(WaitTree(s, ifr, false, "教师工作台"), "教师工作台")
                    ?? FindIframeRef(WaitTree(s, ifr, false, "教师工作台"));
                ExpectFail(run, "int-neg-scroll-iframe", s.Interact(ifr, "scroll", frameRef), "iframe");
                ExpectFail(run, "int-neg-press-iframe", s.Interact(ifr, "press", frameRef, key: "Tab"), "iframe");
                ExpectFail(run, "int-neg-select-iframe", s.Interact(ifr, "select", frameRef, option: "x"), "iframe");
            }

            ToolResult spaceKey = s.Interact(action: "press", key: " ");
            Ok(
                run,
                "int-neg-press-bad-space",
                spaceKey != null && !spaceKey.Success && ContainsAny(spaceKey.Error, "白名单", "key"),
                spaceKey != null && spaceKey.Success ? "空格键应失败" : spaceKey?.Error);
        }

        private static void RunExtraTreeIframe(TestRun run, NavSession s, FixtureHttp http)
        {
            string ifr = OpenOk(run, s, http.Url("iframe", "outer.html"), "int-type-iframe-inner");
            if (ifr == null)
            {
                FailMany(run, "依赖 outer", "int-click-iframe-result");
            }
            else
            {
                string t0 = WaitTree(s, ifr, false, "教师工作台");
                string frameRef = FindRef(t0, "教师工作台") ?? FindIframeRef(t0);
                ToolResult exp = string.IsNullOrEmpty(frameRef) ? null : s.Interact(ifr, "click", frameRef);
                string et = exp == null ? "" : Tree(exp);
                string note = FindRef(et, "子页备注");
                if (string.IsNullOrEmpty(note))
                {
                    et = WaitTree(s, ifr, true, "子页备注");
                    note = FindRef(et, "子页备注");
                }
                TypeValue(run, s, ifr, ref et, "int-type-iframe-inner", note, "内页已填", "内页已填");
                string link = FindRef(et, "评教结果") ?? FindRef(WaitTree(s, ifr, false, "评教结果"), "评教结果");
                ToolResult go = string.IsNullOrEmpty(link) ? null : s.Interact(ifr, "click", link);
                Ok(
                    run,
                    "int-click-iframe-result",
                    go != null && go.Success && (Tree(go).Contains("评教已打开") || Str(go, "title").Contains("评教") || Str(go, "url").Contains("clicked")),
                    go == null ? "无评教结果" : (go.Error ?? "未打开"));
            }

            string tabs = OpenOk(run, s, http.Url("center-tabs", "index.html"), "int-click-tab-detail");
            if (tabs == null)
            {
                run.CaseFail("int-click-tab-task", "依赖 center-tabs");
            }
            else
            {
                string t = WaitTree(s, tabs, false, "项目详情", "项目任务");
                ClickStatus(run, s, tabs, ref t, "int-click-tab-detail", FindRef(t, "项目详情"), "项目详情");
                t = WaitTree(s, tabs, false, "项目任务");
                ClickStatus(run, s, tabs, ref t, "int-click-tab-task", FindRef(t, "项目任务"), "项目任务");
            }

            string hid = OpenOk(run, s, http.Url("form-controls", "index.html"), "int-type-date-hidden", visible: false);
            if (hid != null)
            {
                string t = WaitTree(s, hid, true, "立项日");
                TypeValue(run, s, hid, ref t, "int-type-date-hidden", DateHint(t), "2026-12-01", "2026-12-01");
            }
        }

        private static void TypeOk(TestRun run, NavSession s, string id, string caseId, string refId, string text)
        {
            if (string.IsNullOrEmpty(refId))
            {
                run.CaseFail(caseId, "无控件 ref");
                return;
            }

            ToolResult r = s.Interact(id, "type", refId, text: text);
            Ok(run, caseId, r != null && r.Success, r == null ? "无回包" : r.Error);
        }

        private static string IsoWeek(DateTime date)
        {
            int offset = ((int)date.DayOfWeek + 6) % 7;
            DateTime thursday = date.AddDays(3 - offset);
            int week = CultureInfo.InvariantCulture.Calendar.GetWeekOfYear(
                thursday,
                CalendarWeekRule.FirstFourDayWeek,
                DayOfWeek.Monday);
            return thursday.Year.ToString("0000", CultureInfo.InvariantCulture)
                + "-W" + week.ToString("00", CultureInfo.InvariantCulture);
        }
    }
}
