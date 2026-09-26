using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Threading;
using System.Threading.Tasks;
using WordAddIn1;
using WordAddIn1.OpenFiles;
using WordAddIn1.PresentationHost;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PptHtmlContentRoundtripTest
{
    /// <summary>F_ppt_transition。前缀 trn-。</summary>
    internal static class TransitionRunner
    {
        private static readonly HashSet<string> Scored = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        private static readonly string[] ComIds =
        {
            "trn-neg-list-slide", "trn-neg-list-slide-ids", "trn-neg-get-no-slide",
            "trn-neg-set-no-slide", "trn-neg-clear-no-slide", "trn-neg-set-no-effect",
            "trn-neg-bad-effect", "trn-neg-push-no-dir", "trn-neg-cover-no-dir",
            "trn-neg-uncover-no-dir", "trn-neg-bad-dir", "trn-neg-wipe-no-dir",
            "trn-list-none", "trn-set-fade", "trn-get-fade", "trn-set-push",
            "trn-set-wipe", "trn-set-cover", "trn-set-uncover", "trn-set-cut",
            "trn-set-none", "trn-set-dir-down", "trn-list-count", "trn-list-after-two",
            "trn-other-untouched", "trn-set-twice", "trn-get-after-clear",
            "trn-clear-none", "trn-clear-keeps-advance", "trn-set-advance",
            "trn-set-advance-click", "trn-advance-omit-keeps", "trn-set-duration",
            "trn-hidden-ok", "trn-hidden-list", "trn-missing-slide", "trn-sid-prefix",
            "trn-duration-neg", "trn-advance-neg", "trn-echo-kind", "trn-echo-slide",
            "trn-echo-display", "trn-action-upper", "trn-action-spaces",
            "trn-clear-not-anim", "trn-three-list"
        };

        public static bool Wants(IList<string> nameFilters)
        {
            if (nameFilters == null || nameFilters.Count == 0)
            {
                return false;
            }

            for (int i = 0; i < nameFilters.Count; i++)
            {
                string f = nameFilters[i] ?? "";
                if (f.IndexOf("trn-", StringComparison.OrdinalIgnoreCase) >= 0
                    || string.Equals(f, "trn", StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
            }

            return false;
        }

        public static void Run(TestRun run, IList<string> nameFilters, bool useWpp)
        {
            Scored.Clear();
            Console.WriteLine(useWpp ? "换页过渡 · WPP" : "换页过渡 · PPT");
            RunParse(run, nameFilters);
            if (NeedCom(nameFilters))
            {
                RunCom(run, nameFilters, useWpp);
            }

            Console.WriteLine("换页过渡 通过 " + run.CasesPassed
                + "  失败 " + run.CasesFailed
                + "  跳过 " + run.CasesSkipped);
        }

        private static bool NeedCom(IList<string> filters)
        {
            if (IsBroad(filters))
            {
                return true;
            }

            for (int i = 0; i < ComIds.Length; i++)
            {
                if (Want(filters, ComIds[i]))
                {
                    return true;
                }
            }

            return false;
        }

        private static bool IsBroad(IList<string> filters)
        {
            if (filters == null || filters.Count == 0)
            {
                return true;
            }

            for (int i = 0; i < filters.Count; i++)
            {
                string f = filters[i] ?? "";
                if (string.Equals(f, "trn", StringComparison.OrdinalIgnoreCase)
                    || string.Equals(f, "trn-", StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
            }

            return false;
        }

        private static bool Want(IList<string> filters, string id)
        {
            if (IsBroad(filters))
            {
                return true;
            }

            for (int i = 0; i < filters.Count; i++)
            {
                string f = filters[i] ?? "";
                if (f.Length > 0
                    && (id.IndexOf(f, StringComparison.OrdinalIgnoreCase) >= 0
                        || f.IndexOf(id, StringComparison.OrdinalIgnoreCase) >= 0))
                {
                    return true;
                }
            }

            return false;
        }

        private static void Finish(TestRun run, string id, List<string> fails)
        {
            if (Scored.Contains(id))
            {
                return;
            }

            Scored.Add(id);
            if (fails == null || fails.Count == 0)
            {
                run.CaseOk(id);
            }
            else
            {
                run.CaseFail(id, string.Join("；", fails));
            }
        }

        private static void RunParse(TestRun run, IList<string> filters)
        {
            Console.WriteLine("--- A 渠道 / 互斥（stub）---");
            ChannelRegistry.Register(new KindStub("trn-stub-ppt", ChannelKind.Ppt), false);
            ChannelRegistry.Register(new KindStub("trn-stub-word", ChannelKind.Word), false);
            ChannelRegistry.Register(new KindStub("trn-stub-excel", ChannelKind.Excel), false);

            if (Want(filters, "trn-neg-no-channel"))
            {
                ExpectFail(run, "trn-neg-no-channel", new Dictionary<string, object>(), null);
            }

            if (Want(filters, "trn-neg-no-action"))
            {
                ExpectFail(run, "trn-neg-no-action", Args("trn-stub-ppt"), "list|get|set|clear");
            }

            if (Want(filters, "trn-neg-bad-action"))
            {
                var a = Args("trn-stub-ppt");
                a["action"] = "add";
                ExpectFail(run, "trn-neg-bad-action", a, null);
            }

            if (Want(filters, "trn-neg-apply-alias"))
            {
                var a = Args("trn-stub-ppt");
                a["action"] = "apply";
                a["slide_id"] = "1";
                a["effect"] = "fade";
                ExpectFail(run, "trn-neg-apply-alias", a, null);
            }

            if (Want(filters, "trn-neg-word"))
            {
                var a = Args("trn-stub-word");
                a["action"] = "list";
                ExpectFail(run, "trn-neg-word", a, "渠道");
            }

            if (Want(filters, "trn-neg-excel"))
            {
                var a = Args("trn-stub-excel");
                a["action"] = "list";
                ExpectFail(run, "trn-neg-excel", a, "渠道");
            }
        }

        private static void ExpectFail(
            TestRun run,
            string id,
            Dictionary<string, object> args,
            string contains)
        {
            var fails = new List<string>();
            ToolResult r = Invoke(args);
            if (r == null)
            {
                fails.Add("工具不可调");
            }
            else if (r.Success)
            {
                fails.Add("应失败却成功");
            }
            else if (contains != null && !ContainsAny(r.Error, contains))
            {
                fails.Add("文案应含「" + contains + "」实际 " + (r.Error ?? ""));
            }

            Finish(run, id, fails);
        }

        private static void RunCom(TestRun run, IList<string> filters, bool useWpp)
        {
            Console.WriteLine(useWpp ? "--- A Host / B COM（WPP）---" : "--- A Host / B COM ---");
            if (useWpp)
            {
                RunComWpp(run, filters);
            }
            else
            {
                RunComPpt(run, filters);
            }
        }

        private static void RunComPpt(TestRun run, IList<string> filters)
        {
            PowerPoint.Application app = null;
            PowerPoint.Presentation pres = null;
            try
            {
                Thread.Sleep(600);
                app = new PowerPoint.Application();
                app.Visible = Office.MsoTriState.msoTrue;
                try
                {
                    app.DisplayAlerts = PowerPoint.PpAlertLevel.ppAlertsNone;
                }
                catch
                {
                }

                pres = AddPpt(app);
                PptChannel ch = ChannelRegistry.CreateOrGetPpt(pres, "", false);
                if (!ch.TryGetLivePresentation(out PowerPoint.Presentation live) || live == null)
                {
                    FailNeeded(run, filters, "渠道拿不到活稿");
                    return;
                }

                RunBlocks(run, filters, new Deck(ch, live, wpp: false));
            }
            catch (Exception ex)
            {
                FailNeeded(run, filters, "COM 失败: " + ex.Message);
            }
            finally
            {
                try
                {
                    if (pres != null)
                    {
                        pres.Saved = Office.MsoTriState.msoTrue;
                        pres.Close();
                    }
                }
                catch
                {
                }

                try
                {
                    if (app != null && app.Presentations.Count == 0)
                    {
                        app.Quit();
                    }
                }
                catch
                {
                }
            }
        }

        private static PowerPoint.Presentation AddPpt(PowerPoint.Application app)
        {
            Exception last = null;
            for (int i = 0; i < 5; i++)
            {
                try
                {
                    return app.Presentations.Add(Office.MsoTriState.msoTrue);
                }
                catch (Exception ex)
                {
                    last = ex;
                    Thread.Sleep(600);
                }
            }

            throw last ?? new InvalidOperationException("Presentations.Add 失败");
        }

        private static void RunComWpp(TestRun run, IList<string> filters)
        {
            object app = null;
            object pres = null;
            try
            {
                Thread.Sleep(600);
                if (!WppCom.TryCreateApplication(out app, out string startErr))
                {
                    FailNeeded(run, filters, startErr);
                    return;
                }

                object presentations = WppCom.GetProperty(app, "Presentations");
                Exception last = null;
                for (int i = 0; i < 5; i++)
                {
                    try
                    {
                        try
                        {
                            pres = WppCom.Invoke(presentations, "Add", true);
                        }
                        catch
                        {
                            pres = WppCom.Invoke(presentations, "Add");
                        }

                        if (pres != null)
                        {
                            break;
                        }
                    }
                    catch (Exception ex)
                    {
                        last = ex;
                        Thread.Sleep(400);
                    }
                }

                if (pres == null)
                {
                    FailNeeded(run, filters, last == null ? "WPP Add 失败" : last.Message);
                    return;
                }

                WppChannel ch = ChannelRegistry.CreateOrGetWpp(pres, "", false);
                if (!ch.TryGetLivePresentation(out object live) || live == null)
                {
                    FailNeeded(run, filters, "渠道拿不到活稿");
                    return;
                }

                RunBlocks(run, filters, new Deck(ch, live, wpp: true));
            }
            catch (Exception ex)
            {
                FailNeeded(run, filters, "COM 失败: " + ex.Message);
            }
            finally
            {
                try
                {
                    if (pres != null)
                    {
                        WppCom.TrySetProperty(pres, "Saved", true);
                        WppCom.Invoke(pres, "Close");
                    }
                }
                catch
                {
                }

                try
                {
                    if (app != null)
                    {
                        object presentations = WppCom.GetProperty(app, "Presentations");
                        int n = Convert.ToInt32(WppCom.GetProperty(presentations, "Count") ?? 0);
                        if (n == 0)
                        {
                            WppCom.Invoke(app, "Quit");
                        }
                    }
                }
                catch
                {
                }
            }
        }

        private static void RunBlocks(TestRun run, IList<string> filters, Deck deck)
        {
            if (!deck.EnsureCount(2, out string err))
            {
                FailNeeded(run, filters, err ?? "铺页失败");
                return;
            }

            try
            {
                RunHostNeg(run, filters, deck);
                RunWrite(run, filters, deck);
            }
            catch (Exception ex)
            {
                FailNeeded(run, filters, "COM 中断: " + ex.Message);
            }
        }

        private static void RunHostNeg(TestRun run, IList<string> filters, Deck deck)
        {
            if (Want(filters, "trn-neg-list-slide"))
            {
                var a = deck.Args();
                a["action"] = "list";
                a["slide_id"] = deck.Id1;
                ExpectFail(run, "trn-neg-list-slide", a, "slide_id|请用 get");
            }

            if (Want(filters, "trn-neg-list-slide-ids"))
            {
                var a = deck.Args();
                a["action"] = "list";
                a["slide_ids"] = new[] { deck.Id1 };
                ExpectFail(run, "trn-neg-list-slide-ids", a, "slide_ids");
            }

            if (Want(filters, "trn-neg-get-no-slide"))
            {
                var a = deck.Args();
                a["action"] = "get";
                ExpectFail(run, "trn-neg-get-no-slide", a, "slide_id");
            }

            if (Want(filters, "trn-neg-set-no-slide"))
            {
                var a = deck.Args();
                a["action"] = "set";
                a["effect"] = "fade";
                ExpectFail(run, "trn-neg-set-no-slide", a, "slide_id");
            }

            if (Want(filters, "trn-neg-clear-no-slide"))
            {
                var a = deck.Args();
                a["action"] = "clear";
                ExpectFail(run, "trn-neg-clear-no-slide", a, "slide_id");
            }

            if (Want(filters, "trn-neg-set-no-effect"))
            {
                var a = deck.Set(deck.Id1, null, null);
                a.Remove("effect");
                ExpectFail(run, "trn-neg-set-no-effect", a, "effect");
            }

            if (Want(filters, "trn-neg-bad-effect"))
            {
                ExpectFail(run, "trn-neg-bad-effect", deck.Set(deck.Id1, "morph", null), "不支持|白名单");
            }

            if (Want(filters, "trn-neg-push-no-dir"))
            {
                ExpectFail(run, "trn-neg-push-no-dir", deck.Set(deck.Id1, "push", null), "dir");
            }

            if (Want(filters, "trn-neg-cover-no-dir"))
            {
                ExpectFail(run, "trn-neg-cover-no-dir", deck.Set(deck.Id1, "cover", null), "dir");
            }

            if (Want(filters, "trn-neg-uncover-no-dir"))
            {
                ExpectFail(run, "trn-neg-uncover-no-dir", deck.Set(deck.Id1, "uncover", null), "dir");
            }

            if (Want(filters, "trn-neg-bad-dir"))
            {
                ExpectFail(run, "trn-neg-bad-dir", deck.Set(deck.Id1, "push", "diagonal"), "left|right|up|down");
            }

            if (Want(filters, "trn-neg-wipe-no-dir"))
            {
                ExpectFail(run, "trn-neg-wipe-no-dir", deck.Set(deck.Id1, "wipe", null), "dir");
            }
        }

        private static void RunWrite(TestRun run, IList<string> filters, Deck deck)
        {
            if (Want(filters, "trn-list-none"))
            {
                var f = new List<string>();
                ToolResult r = Invoke(deck.List());
                if (!Ok(r, f))
                {
                    Finish(run, "trn-list-none", f);
                }
                else
                {
                    if (HasTopSlide(r))
                    {
                        f.Add("list 不应有顶层 slide_id");
                    }

                    if (CountTr(r) != deck.Count())
                    {
                        f.Add("transitions 长度应为 " + deck.Count() + " 实际 " + CountTr(r));
                    }

                    Finish(run, "trn-list-none", f);
                }
            }

            if (Want(filters, "trn-set-fade"))
            {
                Finish(run, "trn-set-fade", CheckSet(Invoke(deck.Set(deck.Id1, "fade", null)), "fade", null));
            }

            if (Want(filters, "trn-get-fade"))
            {
                Invoke(deck.Set(deck.Id1, "fade", null));
                var f = new List<string>();
                ToolResult r = Invoke(deck.Get(deck.Id1));
                if (!Ok(r, f))
                {
                    Finish(run, "trn-get-fade", f);
                }
                else
                {
                    if (CountTr(r) != 1)
                    {
                        f.Add("get 应单元素");
                    }

                    if (!Eq(GetString(r, "slide_id"), deck.Id1))
                    {
                        f.Add("顶层 slide_id 应为 " + deck.Id1);
                    }

                    Dictionary<string, object> t = FirstTr(r);
                    if (t == null || !Eq(Get(t, "effect"), "fade") || !Eq(Get(t, "slide_id"), deck.Id1))
                    {
                        f.Add("get 应为 id1/fade");
                    }

                    Finish(run, "trn-get-fade", f);
                }
            }

            if (Want(filters, "trn-set-push"))
            {
                Finish(run, "trn-set-push", CheckSet(Invoke(deck.Set(deck.Id1, "push", "left")), "push", "left"));
            }

            if (Want(filters, "trn-set-wipe"))
            {
                Finish(run, "trn-set-wipe", CheckSet(Invoke(deck.Set(deck.Id1, "wipe", "up")), "wipe", "up"));
            }

            if (Want(filters, "trn-set-cover"))
            {
                Finish(run, "trn-set-cover", CheckSet(Invoke(deck.Set(deck.Id1, "cover", "right")), "cover", "right"));
            }

            if (Want(filters, "trn-set-uncover"))
            {
                Finish(run, "trn-set-uncover", CheckSet(Invoke(deck.Set(deck.Id1, "uncover", "down")), "uncover", "down"));
            }

            if (Want(filters, "trn-set-cut"))
            {
                Finish(run, "trn-set-cut", CheckSet(Invoke(deck.Set(deck.Id1, "cut", null)), "cut", null));
            }

            if (Want(filters, "trn-set-none"))
            {
                Invoke(deck.Set(deck.Id1, "fade", null));
                Finish(run, "trn-set-none", CheckSet(Invoke(deck.Set(deck.Id1, "none", null)), "none", null));
            }

            if (Want(filters, "trn-set-dir-down"))
            {
                Finish(run, "trn-set-dir-down", CheckSet(Invoke(deck.Set(deck.Id1, "push", "down")), "push", "down"));
            }

            if (Want(filters, "trn-list-count"))
            {
                var f = new List<string>();
                ToolResult r = Invoke(deck.List());
                if (!Ok(r, f))
                {
                    Finish(run, "trn-list-count", f);
                }
                else
                {
                    IList<Dictionary<string, object>> list = Trs(r);
                    if (list.Count != 2)
                    {
                        f.Add("长度应为 2 实际 " + list.Count);
                    }
                    else
                    {
                        if (GetInt(list[0], "index") != 1 || GetInt(list[1], "index") != 2)
                        {
                            f.Add("index 应为 1,2");
                        }

                        if (!Eq(Get(list[0], "slide_id"), deck.Id1) || !Eq(Get(list[1], "slide_id"), deck.Id2))
                        {
                            f.Add("slide_id 未对齐 COM");
                        }
                    }

                    Finish(run, "trn-list-count", f);
                }
            }

            if (Want(filters, "trn-list-after-two"))
            {
                Invoke(deck.Set(deck.Id1, "fade", null));
                Invoke(deck.Set(deck.Id2, "push", "left"));
                var f = new List<string>();
                ToolResult r = Invoke(deck.List());
                if (!Ok(r, f))
                {
                    Finish(run, "trn-list-after-two", f);
                }
                else
                {
                    IList<Dictionary<string, object>> list = Trs(r);
                    Dictionary<string, object> a = FindTr(list, deck.Id1);
                    Dictionary<string, object> b = FindTr(list, deck.Id2);
                    if (a == null || !Eq(Get(a, "effect"), "fade"))
                    {
                        f.Add("第 1 页应为 fade");
                    }

                    if (b == null || !Eq(Get(b, "effect"), "push"))
                    {
                        f.Add("第 2 页应为 push");
                    }

                    Finish(run, "trn-list-after-two", f);
                }
            }

            if (Want(filters, "trn-other-untouched"))
            {
                Invoke(deck.Clear(deck.Id2));
                string before = EffectOf(Invoke(deck.Get(deck.Id2)));
                Invoke(deck.Set(deck.Id1, "fade", null));
                var f = new List<string>();
                ToolResult r = Invoke(deck.Get(deck.Id2));
                if (!Ok(r, f))
                {
                    Finish(run, "trn-other-untouched", f);
                }
                else if (!Eq(EffectOf(r), before) && !Eq(EffectOf(r), "none"))
                {
                    f.Add("第 2 页被改了 实际 " + EffectOf(r));
                    Finish(run, "trn-other-untouched", f);
                }
                else
                {
                    Finish(run, "trn-other-untouched", f);
                }
            }

            if (Want(filters, "trn-set-twice"))
            {
                Invoke(deck.Set(deck.Id1, "fade", null));
                Invoke(deck.Set(deck.Id1, "push", "left"));
                var f = new List<string>();
                ToolResult r = Invoke(deck.Get(deck.Id1));
                if (!Ok(r, f))
                {
                    Finish(run, "trn-set-twice", f);
                }
                else if (!Eq(EffectOf(r), "push"))
                {
                    f.Add("应为 push 不是 fade 实际 " + EffectOf(r));
                    Finish(run, "trn-set-twice", f);
                }
                else
                {
                    Finish(run, "trn-set-twice", f);
                }
            }

            if (Want(filters, "trn-get-after-clear"))
            {
                Invoke(deck.Set(deck.Id1, "fade", null));
                Invoke(deck.Clear(deck.Id1));
                Finish(run, "trn-get-after-clear", CheckSet(Invoke(deck.Get(deck.Id1)), "none", null));
            }

            if (Want(filters, "trn-clear-none"))
            {
                Invoke(deck.Set(deck.Id1, "fade", null));
                Finish(run, "trn-clear-none", CheckSet(Invoke(deck.Clear(deck.Id1)), "none", null));
            }

            if (Want(filters, "trn-clear-keeps-advance"))
            {
                var set = deck.Set(deck.Id1, "fade", null);
                set["advance_on_click"] = false;
                Invoke(set);
                var f = CheckSet(Invoke(deck.Clear(deck.Id1)), "none", null);
                if (f.Count == 0)
                {
                    Dictionary<string, object> t = FirstTr(Invoke(deck.Get(deck.Id1)));
                    if (t == null || GetBool(t, "advance_on_click"))
                    {
                        f.Add("clear 后 advance_on_click 应仍 false");
                    }
                }

                Finish(run, "trn-clear-keeps-advance", f);
            }

            if (Want(filters, "trn-set-advance"))
            {
                var a = deck.Set(deck.Id1, "fade", null);
                a["advance_after_ms"] = 2000;
                var f = new List<string>();
                ToolResult r = Invoke(a);
                if (!Ok(r, f))
                {
                    Finish(run, "trn-set-advance", f);
                }
                else
                {
                    int ms = FirstTr(r) == null ? 0 : GetInt(FirstTr(r), "advance_after_ms");
                    if (!Near(ms, 2000, 300))
                    {
                        f.Add("advance_after_ms 应接近 2000 实际 " + ms);
                    }

                    Finish(run, "trn-set-advance", f);
                }
            }

            if (Want(filters, "trn-set-advance-click"))
            {
                var a = deck.Set(deck.Id1, "fade", null);
                a["advance_on_click"] = true;
                var f = new List<string>();
                ToolResult r = Invoke(a);
                if (!Ok(r, f))
                {
                    Finish(run, "trn-set-advance-click", f);
                }
                else
                {
                    Dictionary<string, object> t = FirstTr(r);
                    if (t == null || !GetBool(t, "advance_on_click"))
                    {
                        f.Add("advance_on_click 应为 true");
                    }

                    Finish(run, "trn-set-advance-click", f);
                }
            }

            if (Want(filters, "trn-advance-omit-keeps"))
            {
                var first = deck.Set(deck.Id1, "fade", null);
                first["advance_on_click"] = false;
                Invoke(first);
                var f = new List<string>();
                ToolResult r = Invoke(deck.Set(deck.Id1, "fade", null));
                if (!Ok(r, f))
                {
                    Finish(run, "trn-advance-omit-keeps", f);
                }
                else
                {
                    Dictionary<string, object> t = FirstTr(r);
                    if (t == null || GetBool(t, "advance_on_click"))
                    {
                        f.Add("省略 advance 后 click 应仍 false");
                    }

                    Finish(run, "trn-advance-omit-keeps", f);
                }
            }

            if (Want(filters, "trn-set-duration"))
            {
                var a = deck.Set(deck.Id1, "fade", null);
                a["duration_ms"] = 500;
                var f = new List<string>();
                ToolResult r = Invoke(a);
                if (!Ok(r, f))
                {
                    Finish(run, "trn-set-duration", f);
                }
                else
                {
                    int ms = FirstTr(r) == null ? 0 : GetInt(FirstTr(r), "duration_ms");
                    if (!Near(ms, 500, 250))
                    {
                        f.Add("duration_ms 应接近 500 实际 " + ms);
                    }

                    Finish(run, "trn-set-duration", f);
                }
            }

            if (Want(filters, "trn-hidden-ok"))
            {
                if (!deck.TryHide(2, true, out string hideErr))
                {
                    run.CaseFail("trn-hidden-ok", hideErr ?? "隐藏失败");
                }
                else
                {
                    var f = CheckSet(Invoke(deck.Set(deck.Id2, "fade", null)), "fade", null);
                    if (f.Count == 0 && !deck.IsHidden(2))
                    {
                        f.Add("Hidden 应变真却丢了");
                    }

                    Finish(run, "trn-hidden-ok", f);
                    deck.TryHide(2, false, out _);
                }
            }

            if (Want(filters, "trn-hidden-list"))
            {
                if (!deck.TryHide(2, true, out string hideErr))
                {
                    run.CaseFail("trn-hidden-list", hideErr ?? "隐藏失败");
                }
                else
                {
                    var f = new List<string>();
                    ToolResult r = Invoke(deck.List());
                    if (!Ok(r, f))
                    {
                        Finish(run, "trn-hidden-list", f);
                    }
                    else if (FindTr(Trs(r), deck.Id2) == null)
                    {
                        f.Add("list 丢了隐藏页");
                        Finish(run, "trn-hidden-list", f);
                    }
                    else
                    {
                        Finish(run, "trn-hidden-list", f);
                    }

                    deck.TryHide(2, false, out _);
                }
            }

            if (Want(filters, "trn-missing-slide"))
            {
                ExpectFail(run, "trn-missing-slide", deck.Get("999999"), "不存在");
            }

            if (Want(filters, "trn-sid-prefix"))
            {
                ExpectFail(run, "trn-sid-prefix", deck.Get("sid" + deck.Id1), null);
            }

            if (Want(filters, "trn-duration-neg"))
            {
                Invoke(deck.Set(deck.Id1, "fade", null));
                string before = EffectOf(Invoke(deck.Get(deck.Id1)));
                var a = deck.Set(deck.Id1, "push", "left");
                a["duration_ms"] = -1;
                var f = new List<string>();
                ToolResult r = Invoke(a);
                if (r == null || r.Success)
                {
                    f.Add("应失败");
                }
                else if (!ContainsAny(r.Error, "不能为负"))
                {
                    f.Add("文案应含不能为负 实际 " + (r.Error ?? ""));
                }

                if (!Eq(EffectOf(Invoke(deck.Get(deck.Id1))), before))
                {
                    f.Add("效果被改了");
                }

                Finish(run, "trn-duration-neg", f);
            }

            if (Want(filters, "trn-advance-neg"))
            {
                var a = deck.Set(deck.Id1, "fade", null);
                a["advance_after_ms"] = -1;
                ExpectFail(run, "trn-advance-neg", a, "不能为负");
            }

            if (Want(filters, "trn-echo-kind"))
            {
                var f = new List<string>();
                ToolResult r = Invoke(deck.List());
                if (!Ok(r, f))
                {
                    Finish(run, "trn-echo-kind", f);
                }
                else
                {
                    string expect = deck.Wpp ? "wpp" : "ppt";
                    if (!Eq(GetString(r, "kind"), expect))
                    {
                        f.Add("kind 应为 " + expect + " 实际 " + GetString(r, "kind"));
                    }

                    Finish(run, "trn-echo-kind", f);
                }
            }

            if (Want(filters, "trn-echo-slide"))
            {
                var f = new List<string>();
                ToolResult get = Invoke(deck.Get(deck.Id1));
                ToolResult list = Invoke(deck.List());
                if (!Ok(get, f) || !Ok(list, f))
                {
                    Finish(run, "trn-echo-slide", f);
                }
                else
                {
                    if (!Eq(GetString(get, "slide_id"), deck.Id1))
                    {
                        f.Add("get 顶层 slide_id 应为该页");
                    }

                    if (HasTopSlide(list))
                    {
                        f.Add("list 仍无顶层 id 却有了");
                    }

                    Finish(run, "trn-echo-slide", f);
                }
            }

            if (Want(filters, "trn-echo-display"))
            {
                Invoke(deck.Set(deck.Id1, "fade", null));
                var f = new List<string>();
                ToolResult r = Invoke(deck.Get(deck.Id1));
                if (!Ok(r, f))
                {
                    Finish(run, "trn-echo-display", f);
                }
                else
                {
                    string d = GetString(r, "display_contents");
                    if (d.IndexOf("slide_id", StringComparison.OrdinalIgnoreCase) < 0
                        || d.IndexOf("effect", StringComparison.OrdinalIgnoreCase) < 0)
                    {
                        f.Add("display_contents 应含 slide_id 与 effect");
                    }

                    Finish(run, "trn-echo-display", f);
                }
            }

            if (Want(filters, "trn-action-upper"))
            {
                var a = deck.Set(deck.Id1, "fade", null);
                a["action"] = "SET";
                ToolResult once = Invoke(a);
                var f = CheckSet(once, "fade", null);
                if (f.Count == 0 && !Eq(GetString(once, "action"), "set"))
                {
                    f.Add("回包 action 应为 set 实际 " + GetString(once, "action"));
                }

                Finish(run, "trn-action-upper", f);
            }

            if (Want(filters, "trn-action-spaces"))
            {
                var a = deck.Set(deck.Id1, "fade", null);
                a["action"] = " set ";
                ToolResult r = Invoke(a);
                var f = CheckSet(r, "fade", null);
                if (f.Count == 0 && !Eq(GetString(r, "action"), "set"))
                {
                    f.Add("回包 action 应为 set 实际 " + GetString(r, "action"));
                }

                Finish(run, "trn-action-spaces", f);
            }

            if (Want(filters, "trn-three-list"))
            {
                if (!deck.EnsureCount(3, out string threeErr))
                {
                    Finish(run, "trn-three-list", new List<string> { threeErr ?? "第 3 页失败" });
                }
                else
                {
                    var f = new List<string>();
                    ToolResult r = Invoke(deck.List());
                    if (!Ok(r, f))
                    {
                        Finish(run, "trn-three-list", f);
                    }
                    else
                    {
                        IList<Dictionary<string, object>> list = Trs(r);
                        if (list.Count != 3)
                        {
                            f.Add("长度应为 3 实际 " + list.Count);
                        }
                        else if (GetInt(list[0], "index") != 1
                            || GetInt(list[1], "index") != 2
                            || GetInt(list[2], "index") != 3)
                        {
                            f.Add("index 应为 1..3");
                        }

                        Finish(run, "trn-three-list", f);
                    }
                }
            }

            if (Want(filters, "trn-clear-not-anim"))
            {
                if (!deck.TryBox(1, out string box, out string boxErr))
                {
                    Finish(run, "trn-clear-not-anim", new List<string> { boxErr ?? "铺框失败" });
                }
                else
                {
                    var ani = deck.AniArgs();
                    ani["action"] = "add";
                    ani["ShapeId"] = box;
                    ani["category"] = "entrance";
                    ani["effect"] = "fade";
                    ani["slide_id"] = deck.Id1;
                    ToolResult added = InvokeAni(ani);
                    var f = new List<string>();
                    if (added == null || !added.Success)
                    {
                        f.Add(added == null ? "动画 add 无回包" : (added.Error ?? "动画 add 失败"));
                    }

                    Invoke(deck.Set(deck.Id1, "fade", null));
                    ToolResult cleared = Invoke(deck.Clear(deck.Id1));
                    if (!Ok(cleared, f))
                    {
                        Finish(run, "trn-clear-not-anim", f);
                    }
                    else if (!Eq(EffectOf(cleared), "none"))
                    {
                        f.Add("过渡应为 none");
                        Finish(run, "trn-clear-not-anim", f);
                    }
                    else
                    {
                        var list = deck.AniArgs();
                        list["action"] = "list";
                        list["slide_id"] = deck.Id1;
                        ToolResult still = InvokeAni(list);
                        if (still == null || !still.Success)
                        {
                            f.Add("动画 list 失败");
                        }
                        else if (!HasAni(still, box, "fade"))
                        {
                            f.Add("页内动画被过渡 clear 清掉了");
                        }

                        Finish(run, "trn-clear-not-anim", f);
                    }
                }
            }
        }

        private static void FailNeeded(TestRun run, IList<string> filters, string err)
        {
            for (int i = 0; i < ComIds.Length; i++)
            {
                if (Want(filters, ComIds[i]) && !Scored.Contains(ComIds[i]))
                {
                    Scored.Add(ComIds[i]);
                    run.CaseFail(ComIds[i], err ?? "铺页失败");
                }
            }
        }

        private static Dictionary<string, object> Args(string channelId)
        {
            return new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase)
            {
                ["channel_id"] = channelId
            };
        }

        private static ToolResult Invoke(Dictionary<string, object> args)
        {
            var tools = new Dictionary<string, Func<Dictionary<string, object>, Task<ToolResult>>>(
                StringComparer.OrdinalIgnoreCase);
            F_PptTransitionTool.Register(tools, null);
            if (!tools.TryGetValue("F_ppt_transition", out Func<Dictionary<string, object>, Task<ToolResult>> fn))
            {
                return null;
            }

            return fn(args ?? new Dictionary<string, object>()).GetAwaiter().GetResult();
        }

        private static ToolResult InvokeAni(Dictionary<string, object> args)
        {
            var tools = new Dictionary<string, Func<Dictionary<string, object>, Task<ToolResult>>>(
                StringComparer.OrdinalIgnoreCase);
            F_PptAnimationTool.Register(tools, null);
            if (!tools.TryGetValue("F_ppt_animation", out Func<Dictionary<string, object>, Task<ToolResult>> fn))
            {
                return null;
            }

            return fn(args ?? new Dictionary<string, object>()).GetAwaiter().GetResult();
        }

        private static List<string> CheckSet(ToolResult r, string effect, string dir)
        {
            var f = new List<string>();
            if (!Ok(r, f))
            {
                return f;
            }

            Dictionary<string, object> t = FirstTr(r);
            if (t == null)
            {
                f.Add("无 transitions[0]");
                return f;
            }

            if (!Eq(Get(t, "effect"), effect))
            {
                f.Add("effect 应为 " + effect + " 实际 " + Get(t, "effect"));
            }

            if (!string.IsNullOrEmpty(dir) && !Eq(Get(t, "dir"), dir))
            {
                f.Add("dir 应为 " + dir + " 实际 " + Get(t, "dir"));
            }

            return f;
        }

        private static bool Ok(ToolResult r, List<string> fails)
        {
            if (r == null || !r.Success)
            {
                fails.Add(r == null ? "无回包" : (r.Error ?? "失败"));
                return false;
            }

            return true;
        }

        private static bool ContainsAny(string text, string pattern)
        {
            if (string.IsNullOrEmpty(pattern))
            {
                return true;
            }

            string t = text ?? "";
            string[] parts = pattern.Split('|');
            for (int i = 0; i < parts.Length; i++)
            {
                if (parts[i].Length > 0 && t.IndexOf(parts[i], StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    return true;
                }
            }

            return false;
        }

        private static bool Near(int actual, int expect, int tol)
        {
            return Math.Abs(actual - expect) <= tol;
        }

        private static bool Eq(string a, string b)
        {
            return string.Equals(a ?? "", b ?? "", StringComparison.OrdinalIgnoreCase);
        }

        private static IList<Dictionary<string, object>> Trs(ToolResult r)
        {
            var list = new List<Dictionary<string, object>>();
            var d = r != null ? r.Data as Dictionary<string, object> : null;
            if (d == null || !d.TryGetValue("transitions", out object raw) || !(raw is IEnumerable en))
            {
                return list;
            }

            foreach (object item in en)
            {
                if (item is Dictionary<string, object> one)
                {
                    list.Add(one);
                }
            }

            return list;
        }

        private static int CountTr(ToolResult r)
        {
            return Trs(r).Count;
        }

        private static Dictionary<string, object> FirstTr(ToolResult r)
        {
            IList<Dictionary<string, object>> list = Trs(r);
            return list.Count > 0 ? list[0] : null;
        }

        private static Dictionary<string, object> FindTr(IList<Dictionary<string, object>> list, string slideId)
        {
            if (list == null)
            {
                return null;
            }

            for (int i = 0; i < list.Count; i++)
            {
                if (Eq(Get(list[i], "slide_id"), slideId))
                {
                    return list[i];
                }
            }

            return null;
        }

        private static string EffectOf(ToolResult r)
        {
            Dictionary<string, object> t = FirstTr(r);
            return t == null ? "" : Get(t, "effect");
        }

        private static bool HasTopSlide(ToolResult r)
        {
            var d = r != null ? r.Data as Dictionary<string, object> : null;
            if (d == null)
            {
                return false;
            }

            foreach (KeyValuePair<string, object> kv in d)
            {
                if (string.Equals(kv.Key, "slide_id", StringComparison.OrdinalIgnoreCase)
                    && kv.Value != null
                    && !string.IsNullOrWhiteSpace(Convert.ToString(kv.Value)))
                {
                    return true;
                }
            }

            return false;
        }

        private static bool HasAni(ToolResult r, string shapeId, string effect)
        {
            var d = r != null ? r.Data as Dictionary<string, object> : null;
            if (d == null || !d.TryGetValue("effects", out object raw) || !(raw is IEnumerable en))
            {
                return false;
            }

            foreach (object item in en)
            {
                var one = item as Dictionary<string, object>;
                if (one == null)
                {
                    continue;
                }

                if (Eq(Get(one, "ShapeId"), shapeId) && Eq(Get(one, "effect"), effect))
                {
                    return true;
                }
            }

            return false;
        }

        private static string Get(Dictionary<string, object> d, string key)
        {
            if (d == null)
            {
                return "";
            }

            if (d.TryGetValue(key, out object v) && v != null)
            {
                return Convert.ToString(v) ?? "";
            }

            foreach (KeyValuePair<string, object> kv in d)
            {
                if (string.Equals(kv.Key, key, StringComparison.OrdinalIgnoreCase) && kv.Value != null)
                {
                    return Convert.ToString(kv.Value) ?? "";
                }
            }

            return "";
        }

        private static int GetInt(Dictionary<string, object> d, string key)
        {
            string s = Get(d, key);
            int n;
            return int.TryParse(s, NumberStyles.Integer, CultureInfo.InvariantCulture, out n) ? n : 0;
        }

        private static bool GetBool(Dictionary<string, object> d, string key)
        {
            string s = Get(d, key);
            return string.Equals(s, "true", StringComparison.OrdinalIgnoreCase) || s == "1";
        }

        private static string GetString(ToolResult r, string key)
        {
            var d = r != null ? r.Data as Dictionary<string, object> : null;
            if (d == null)
            {
                return "";
            }

            foreach (KeyValuePair<string, object> kv in d)
            {
                if (string.Equals(kv.Key, key, StringComparison.OrdinalIgnoreCase) && kv.Value != null)
                {
                    return Convert.ToString(kv.Value) ?? "";
                }
            }

            return "";
        }

        private sealed class KindStub : IOperationChannel
        {
            public KindStub(string id, ChannelKind kind)
            {
                ChannelId = id;
                Kind = kind;
            }

            public string ChannelId { get; }

            public ChannelKind Kind { get; }
        }

        private sealed class Deck
        {
            public Deck(IOperationChannel channel, object pres, bool wpp)
            {
                Channel = channel;
                Pres = pres;
                Wpp = wpp;
                Refresh();
                EnsureFirstSlide();
            }

            public IOperationChannel Channel { get; }

            public object Pres { get; set; }

            public bool Wpp { get; }

            public string Id1
            {
                get { return IdAt(1); }
            }

            public string Id2
            {
                get { return IdAt(2); }
            }

            public Dictionary<string, object> Args()
            {
                return new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase)
                {
                    ["channel_id"] = Channel.ChannelId
                };
            }

            public Dictionary<string, object> AniArgs()
            {
                return Args();
            }

            public Dictionary<string, object> List()
            {
                var a = Args();
                a["action"] = "list";
                return a;
            }

            public Dictionary<string, object> Get(string slideId)
            {
                var a = Args();
                a["action"] = "get";
                a["slide_id"] = slideId;
                return a;
            }

            public Dictionary<string, object> Set(string slideId, string effect, string dir)
            {
                var a = Args();
                a["action"] = "set";
                a["slide_id"] = slideId;
                if (effect != null)
                {
                    a["effect"] = effect;
                }

                if (!string.IsNullOrEmpty(dir))
                {
                    a["dir"] = dir;
                }

                return a;
            }

            public Dictionary<string, object> Clear(string slideId)
            {
                var a = Args();
                a["action"] = "clear";
                a["slide_id"] = slideId;
                return a;
            }

            public void Refresh()
            {
                if (Wpp)
                {
                    var w = Channel as WppChannel;
                    if (w != null && w.TryGetLivePresentation(out object live) && live != null)
                    {
                        Pres = live;
                    }

                    return;
                }

                var p = Channel as PptChannel;
                if (p != null && p.TryGetLivePresentation(out PowerPoint.Presentation livePpt) && livePpt != null)
                {
                    Pres = livePpt;
                }
            }

            public int Count()
            {
                Refresh();
                if (Wpp)
                {
                    object slides = WppCom.GetProperty(Pres, "Slides");
                    return Convert.ToInt32(WppCom.GetProperty(slides, "Count") ?? 0);
                }

                return ((PowerPoint.Presentation)Pres).Slides.Count;
            }

            public string IdAt(int index1)
            {
                Refresh();
                if (Wpp)
                {
                    object slides = WppCom.GetProperty(Pres, "Slides");
                    object slide = WppCom.GetIndexed(slides, index1);
                    return Convert.ToString(
                        WppCom.GetProperty(slide, "SlideID") ?? WppCom.GetProperty(slide, "SlideId")) ?? "";
                }

                return ((PowerPoint.Presentation)Pres).Slides[index1].SlideID.ToString(CultureInfo.InvariantCulture);
            }

            public bool EnsureCount(int n, out string error)
            {
                error = null;
                EnsureFirstSlide();
                while (Count() < n)
                {
                    if (!TryAddBlank(out error))
                    {
                        return false;
                    }
                }

                return Count() >= n;
            }

            public bool TryHide(int index1, bool hidden, out string error)
            {
                error = null;
                Refresh();
                try
                {
                    if (Wpp)
                    {
                        object slides = WppCom.GetProperty(Pres, "Slides");
                        object slide = WppCom.GetIndexed(slides, index1);
                        object tr = WppCom.GetProperty(slide, "SlideShowTransition");
                        WppCom.TrySetProperty(tr, "Hidden", hidden ? -1 : 0);
                        return true;
                    }

                    ((PowerPoint.Presentation)Pres).Slides[index1].SlideShowTransition.Hidden = hidden
                        ? Office.MsoTriState.msoTrue
                        : Office.MsoTriState.msoFalse;
                    return true;
                }
                catch (Exception ex)
                {
                    error = ex.Message;
                    return false;
                }
            }

            public bool IsHidden(int index1)
            {
                Refresh();
                try
                {
                    if (Wpp)
                    {
                        object slides = WppCom.GetProperty(Pres, "Slides");
                        object slide = WppCom.GetIndexed(slides, index1);
                        object tr = WppCom.GetProperty(slide, "SlideShowTransition");
                        object raw = WppCom.GetProperty(tr, "Hidden");
                        if (raw == null)
                        {
                            return false;
                        }

                        if (raw is bool b)
                        {
                            return b;
                        }

                        int n = Convert.ToInt32(raw);
                        return n == -1 || n == 1;
                    }

                    return ((PowerPoint.Presentation)Pres).Slides[index1].SlideShowTransition.Hidden
                        == Office.MsoTriState.msoTrue;
                }
                catch
                {
                    return false;
                }
            }

            public bool TryBox(int index1, out string id, out string error)
            {
                id = null;
                Refresh();
                string slideId = IdAt(index1);
                if (!ContentHtml.TryParse(
                        ContentHtml.TextboxCreate("过渡页框", 8, 8, 28, 12),
                        slideId,
                        out PptHtmlApplyPlan plan,
                        out error))
                {
                    return false;
                }

                plan.AllowCreate = true;
                PptHtmlApplyResult result;
                bool ok = Wpp
                    ? PptHtmlWppApplier.TryApply(Pres, plan, "trn-wpp", out result, out error)
                    : PptHtmlPowerPointApplier.TryApply(
                        (PowerPoint.Presentation)Pres, plan, "trn-ppt", out result, out error);
                if (!ok || result == null || result.CreatedShapes == null || result.CreatedShapes.Count < 1)
                {
                    error = error ?? "无新建框";
                    return false;
                }

                if (result.CreatedShapes[0].TryGetValue("shape_id", out object v) && v != null)
                {
                    id = v.ToString();
                    return !string.IsNullOrEmpty(id);
                }

                error = "无 shape_id";
                return false;
            }

            private void EnsureFirstSlide()
            {
                Refresh();
                if (Count() > 0)
                {
                    return;
                }

                TryAddBlank(out _);
            }

            private bool TryAddBlank(out string error)
            {
                error = null;
                Refresh();
                try
                {
                    if (Wpp)
                    {
                        object slides = WppCom.GetProperty(Pres, "Slides");
                        int n = Convert.ToInt32(WppCom.GetProperty(slides, "Count") ?? 0);
                        try
                        {
                            WppCom.Invoke(slides, "Add", n + 1, 12);
                        }
                        catch
                        {
                            WppCom.Invoke(slides, "Add", n + 1, 1);
                        }

                        return true;
                    }

                    PowerPoint.Presentation ppt = (PowerPoint.Presentation)Pres;
                    ppt.Slides.Add(ppt.Slides.Count + 1, PowerPoint.PpSlideLayout.ppLayoutBlank);
                    return true;
                }
                catch (Exception ex)
                {
                    error = ex.Message;
                    return false;
                }
            }
        }
    }
}
