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
    /// <summary>F_ppt_animation。前缀 ani-。</summary>
    internal static class AnimationRunner
    {
        private static readonly HashSet<string> Scored = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        private static readonly string[] ComIds =
        {
            "ani-neg-list-no-slide", "ani-neg-replace-no-slide", "ani-neg-clear-no-slide",
            "ani-neg-add-no-shape", "ani-neg-bad-category", "ani-neg-bad-effect",
            "ani-neg-fly-no-dir", "ani-neg-fly-out-entrance", "ani-neg-spin-entrance",
            "ani-neg-bad-trigger", "ani-neg-bad-dir", "ani-neg-short-id",
            "ani-list-empty", "ani-add-fade", "ani-list-after-add", "ani-add-appear",
            "ani-add-fly-in", "ani-add-fly-out", "ani-add-wipe", "ani-add-float-up",
            "ani-add-spin", "ani-add-grow", "ani-add-exit", "ani-add-after-prev",
            "ani-add-with-prev", "ani-add-duration", "ani-add-delay", "ani-list-filter",
            "ani-list-other-page", "ani-mismatch-slide", "ani-missing-shape",
            "ani-appear-emphasis", "ani-replace-two", "ani-replace-empty",
            "ani-replace-wipes-old", "ani-replace-half-fail", "ani-two-on-same",
            "ani-clear-page", "ani-clear-one", "ani-clear-then-add", "ani-text-unchanged",
            "ani-neg-duration", "ani-neg-delay", "ani-action-upper", "ani-action-spaces",
            "ani-shape-id-alias", "ani-echo-kind", "ani-echo-slide", "ani-echo-display",
            "ani-add-same-slide"
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
                if (f.IndexOf("ani-", StringComparison.OrdinalIgnoreCase) >= 0
                    || string.Equals(f, "ani", StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
            }

            return false;
        }

        public static void Run(TestRun run, IList<string> nameFilters, bool useWpp)
        {
            Scored.Clear();
            Console.WriteLine(useWpp ? "页内动画 · WPP" : "页内动画 · PPT");
            RunParse(run, nameFilters);
            if (NeedCom(nameFilters))
            {
                RunCom(run, nameFilters, useWpp);
            }

            Console.WriteLine("页内动画 通过 " + run.CasesPassed
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
                if (string.Equals(f, "ani", StringComparison.OrdinalIgnoreCase)
                    || string.Equals(f, "ani-", StringComparison.OrdinalIgnoreCase))
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

        private static void RunParse(TestRun run, IList<string> filters)
        {
            Console.WriteLine("--- A 渠道 / 互斥（stub）---");
            ChannelRegistry.Register(new KindStub("ani-stub-ppt", ChannelKind.Ppt), false);
            ChannelRegistry.Register(new KindStub("ani-stub-word", ChannelKind.Word), false);
            ChannelRegistry.Register(new KindStub("ani-stub-excel", ChannelKind.Excel), false);

            if (Want(filters, "ani-neg-no-channel"))
            {
                ExpectFail(run, "ani-neg-no-channel", new Dictionary<string, object>(), null);
            }

            if (Want(filters, "ani-neg-no-action"))
            {
                ExpectFail(run, "ani-neg-no-action", Args("ani-stub-ppt"), "list|add|replace_all|clear");
            }

            if (Want(filters, "ani-neg-bad-action"))
            {
                var a = Args("ani-stub-ppt");
                a["action"] = "set";
                ExpectFail(run, "ani-neg-bad-action", a, null);
            }

            if (Want(filters, "ani-neg-word"))
            {
                var a = Args("ani-stub-word");
                a["action"] = "list";
                a["slide_id"] = "1";
                ExpectFail(run, "ani-neg-word", a, "渠道");
            }

            if (Want(filters, "ani-neg-excel"))
            {
                var a = Args("ani-stub-excel");
                a["action"] = "list";
                a["slide_id"] = "1";
                ExpectFail(run, "ani-neg-excel", a, "渠道");
            }

            if (Want(filters, "ani-neg-add-two"))
            {
                var a = Args("ani-stub-ppt");
                a["action"] = "add";
                a["effects"] = new List<Dictionary<string, object>>
                {
                    Spec("sid1-s2", "entrance", "fade", null, null),
                    Spec("sid1-s3", "entrance", "fade", null, null)
                };
                ExpectFail(run, "ani-neg-add-two", a, "一条");
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
            Console.WriteLine(useWpp ? "--- A Host / B / C COM（WPP）---" : "--- A Host / B / C COM ---");
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
            if (!deck.Seed(out string err))
            {
                FailNeeded(run, filters, err ?? "铺页失败");
                return;
            }

            try
            {
                RunHostNeg(run, filters, deck);
                RunWrite(run, filters, deck);
                RunReplace(run, filters, deck);
            }
            catch (Exception ex)
            {
                FailNeeded(run, filters, "COM 中断: " + ex.Message);
            }
        }

        private static void RunHostNeg(TestRun run, IList<string> filters, Deck deck)
        {
            if (Want(filters, "ani-neg-list-no-slide"))
            {
                var a = deck.Args();
                a["action"] = "list";
                ExpectFail(run, "ani-neg-list-no-slide", a, "slide_id");
            }

            if (Want(filters, "ani-neg-replace-no-slide"))
            {
                var a = deck.Args();
                a["action"] = "replace_all";
                a["effects"] = new List<Dictionary<string, object>>();
                ExpectFail(run, "ani-neg-replace-no-slide", a, "slide_id");
            }

            if (Want(filters, "ani-neg-clear-no-slide"))
            {
                var a = deck.Args();
                a["action"] = "clear";
                ExpectFail(run, "ani-neg-clear-no-slide", a, "slide_id");
            }

            if (Want(filters, "ani-neg-add-no-shape"))
            {
                var a = deck.Args();
                a["action"] = "add";
                a["category"] = "entrance";
                a["effect"] = "fade";
                ExpectFail(run, "ani-neg-add-no-shape", a, "ShapeId");
            }

            if (Want(filters, "ani-neg-bad-category"))
            {
                ExpectFail(run, "ani-neg-bad-category",
                    deck.AddFlat(deck.A, "move", "fade", null, null, null, null, null),
                    "entrance|emphasis|exit");
            }

            if (Want(filters, "ani-neg-bad-effect"))
            {
                ExpectFail(run, "ani-neg-bad-effect",
                    deck.AddFlat(deck.A, "entrance", "wheel", null, null, null, null, null),
                    "不支持的 effect");
            }

            if (Want(filters, "ani-neg-fly-no-dir"))
            {
                ExpectFail(run, "ani-neg-fly-no-dir",
                    deck.AddFlat(deck.A, "entrance", "fly_in", null, null, null, null, null),
                    "dir");
            }

            if (Want(filters, "ani-neg-fly-out-entrance"))
            {
                ExpectFail(run, "ani-neg-fly-out-entrance",
                    deck.AddFlat(deck.A, "entrance", "fly_out", "left", null, null, null, null),
                    "仅用于 exit");
            }

            if (Want(filters, "ani-neg-spin-entrance"))
            {
                ExpectFail(run, "ani-neg-spin-entrance",
                    deck.AddFlat(deck.A, "entrance", "spin", null, null, null, null, null),
                    "仅用于 emphasis");
            }

            if (Want(filters, "ani-neg-bad-trigger"))
            {
                ExpectFail(run, "ani-neg-bad-trigger",
                    deck.AddFlat(deck.A, "entrance", "fade", null, "on_shape", null, null, null),
                    "on_click|with_previous|after_previous");
            }

            if (Want(filters, "ani-neg-bad-dir"))
            {
                ExpectFail(run, "ani-neg-bad-dir",
                    deck.AddFlat(deck.A, "entrance", "fly_in", "diagonal", null, null, null, null),
                    "left|right|top|bottom");
            }

            if (Want(filters, "ani-neg-short-id"))
            {
                ExpectFail(run, "ani-neg-short-id",
                    deck.AddFlat("s2", "entrance", "fade", null, null, null, null, null),
                    "非法 ShapeId");
            }
        }

        private static void RunWrite(TestRun run, IList<string> filters, Deck deck)
        {
            if (Want(filters, "ani-list-empty"))
            {
                deck.ClearPage();
                var f = new List<string>();
                ToolResult r = Invoke(deck.List(null));
                if (!Ok(r, f) || CountEffects(r) != 0)
                {
                    if (Ok(r, new List<string>()) && CountEffects(r) != 0)
                    {
                        f.Add("effects 应空，实际 " + CountEffects(r));
                    }
                }

                Finish(run, "ani-list-empty", f);
            }

            if (Want(filters, "ani-add-fade"))
            {
                deck.ClearPage();
                Finish(run, "ani-add-fade", CheckAdd(
                    Invoke(deck.AddFlat(deck.A, "entrance", "fade", null, null, null, null, null)),
                    deck.A,
                    "fade",
                    "entrance",
                    null,
                    1));
                if (AbortIfDead(run, filters, deck))
                {
                    return;
                }
            }

            if (Want(filters, "ani-list-after-add"))
            {
                deck.ClearPage();
                Invoke(deck.AddFlat(deck.A, "entrance", "fade", null, null, null, null, null));
                var f = new List<string>();
                ToolResult r = Invoke(deck.List(null));
                if (!Ok(r, f))
                {
                    Finish(run, "ani-list-after-add", f);
                }
                else
                {
                    Dictionary<string, object> e = FindEffect(r, deck.A);
                    if (e == null)
                    {
                        f.Add("list 看不到刚加的 fade");
                    }
                    else if (!Eq(Get(e, "trigger"), "on_click"))
                    {
                        f.Add("默认 trigger 应为 on_click 实际 " + Get(e, "trigger"));
                    }

                    Finish(run, "ani-list-after-add", f);
                }

                if (AbortIfDead(run, filters, deck))
                {
                    return;
                }
            }

            if (Want(filters, "ani-add-appear"))
            {
                deck.ClearPage();
                Finish(run, "ani-add-appear", CheckAdd(
                    Invoke(deck.AddFlat(deck.A, "entrance", "appear", null, null, null, null, null)),
                    deck.A,
                    "appear",
                    "entrance",
                    null,
                    1));
            }

            if (Want(filters, "ani-add-fly-in"))
            {
                deck.ClearPage();
                Finish(run, "ani-add-fly-in", CheckAdd(
                    Invoke(deck.AddFlat(deck.A, "entrance", "fly_in", "left", null, null, null, null)),
                    deck.A,
                    "fly_in",
                    "entrance",
                    "left",
                    1));
            }

            if (Want(filters, "ani-add-fly-out"))
            {
                deck.ClearPage();
                Finish(run, "ani-add-fly-out", CheckAdd(
                    Invoke(deck.AddFlat(deck.A, "exit", "fly_out", "right", null, null, null, null)),
                    deck.A,
                    "fly_out",
                    "exit",
                    "right",
                    1));
            }

            if (Want(filters, "ani-add-wipe"))
            {
                deck.ClearPage();
                Finish(run, "ani-add-wipe", CheckAdd(
                    Invoke(deck.AddFlat(deck.A, "entrance", "wipe", "right", null, null, null, null)),
                    deck.A,
                    "wipe",
                    "entrance",
                    null,
                    1));
            }

            if (Want(filters, "ani-add-float-up"))
            {
                deck.ClearPage();
                Finish(run, "ani-add-float-up", CheckAdd(
                    Invoke(deck.AddFlat(deck.A, "entrance", "float_up", null, null, null, null, null)),
                    deck.A,
                    "float_up",
                    "entrance",
                    null,
                    1));
            }

            if (Want(filters, "ani-add-spin"))
            {
                deck.ClearPage();
                Finish(run, "ani-add-spin", CheckAdd(
                    Invoke(deck.AddFlat(deck.A, "emphasis", "spin", null, null, null, null, null)),
                    deck.A,
                    "spin",
                    "emphasis",
                    null,
                    1));
            }

            if (Want(filters, "ani-add-grow"))
            {
                deck.ClearPage();
                Finish(run, "ani-add-grow", CheckAdd(
                    Invoke(deck.AddFlat(deck.A, "emphasis", "grow_shrink", null, null, null, null, null)),
                    deck.A,
                    "grow_shrink",
                    "emphasis",
                    null,
                    1));
            }

            if (Want(filters, "ani-add-exit"))
            {
                deck.ClearPage();
                Finish(run, "ani-add-exit", CheckAdd(
                    Invoke(deck.AddFlat(deck.A, "exit", "appear", null, null, null, null, null)),
                    deck.A,
                    "appear",
                    "exit",
                    null,
                    1));
            }

            if (Want(filters, "ani-add-after-prev"))
            {
                deck.ClearPage();
                var f = CheckAdd(
                    Invoke(deck.AddFlat(deck.A, "entrance", "fade", null, "after_previous", null, null, null)),
                    deck.A,
                    "fade",
                    "entrance",
                    null,
                    1);
                if (f.Count == 0)
                {
                    Dictionary<string, object> e = FindEffect(
                        Invoke(deck.List(null)),
                        deck.A);
                    if (e == null || !Eq(Get(e, "trigger"), "after_previous"))
                    {
                        f.Add("trigger 应为 after_previous");
                    }
                }

                Finish(run, "ani-add-after-prev", f);
            }

            if (Want(filters, "ani-add-with-prev"))
            {
                deck.ClearPage();
                var f = CheckAdd(
                    Invoke(deck.AddFlat(deck.A, "entrance", "fade", null, "with_previous", null, null, null)),
                    deck.A,
                    "fade",
                    "entrance",
                    null,
                    1);
                if (f.Count == 0)
                {
                    Dictionary<string, object> e = FindEffect(
                        Invoke(deck.List(null)),
                        deck.A);
                    if (e == null || !Eq(Get(e, "trigger"), "with_previous"))
                    {
                        f.Add("trigger 应为 with_previous");
                    }
                }

                Finish(run, "ani-add-with-prev", f);
            }

            if (Want(filters, "ani-add-duration"))
            {
                deck.ClearPage();
                var f = CheckAdd(
                    Invoke(deck.AddFlat(deck.A, "entrance", "fade", null, null, 800, null, null)),
                    deck.A,
                    "fade",
                    "entrance",
                    null,
                    1);
                if (f.Count == 0)
                {
                    Dictionary<string, object> e = FindEffect(
                        Invoke(deck.List(null)),
                        deck.A);
                    int ms = e == null ? 0 : GetInt(e, "duration_ms");
                    if (!Near(ms, 800, 250))
                    {
                        f.Add("duration_ms 应接近 800 实际 " + ms);
                    }
                }

                Finish(run, "ani-add-duration", f);
            }

            if (Want(filters, "ani-add-delay"))
            {
                deck.ClearPage();
                var f = CheckAdd(
                    Invoke(deck.AddFlat(deck.A, "entrance", "fade", null, null, null, 200, null)),
                    deck.A,
                    "fade",
                    "entrance",
                    null,
                    1);
                if (f.Count == 0)
                {
                    Dictionary<string, object> e = FindEffect(
                        Invoke(deck.List(null)),
                        deck.A);
                    int ms = e == null ? 0 : GetInt(e, "delay_ms");
                    if (!Near(ms, 200, 150))
                    {
                        f.Add("delay_ms 应接近 200 实际 " + ms);
                    }
                }

                Finish(run, "ani-add-delay", f);
            }

            if (Want(filters, "ani-list-filter"))
            {
                deck.ClearPage();
                Invoke(deck.AddFlat(deck.A, "entrance", "fade", null, null, null, null, null));
                Invoke(deck.AddFlat(deck.B, "entrance", "appear", null, null, null, null, null));
                var f = new List<string>();
                ToolResult r = Invoke(deck.List(deck.A));
                if (!Ok(r, f))
                {
                    Finish(run, "ani-list-filter", f);
                }
                else
                {
                    if (FindEffect(r, deck.A) == null)
                    {
                        f.Add("过滤后应含 a");
                    }

                    if (FindEffect(r, deck.B) != null)
                    {
                        f.Add("过滤后不应含 b");
                    }

                    Finish(run, "ani-list-filter", f);
                }
            }

            if (Want(filters, "ani-list-other-page"))
            {
                deck.ClearPage();
                Invoke(deck.AddFlat(deck.A, "entrance", "fade", null, null, null, null, null));
                var f = new List<string>();
                ToolResult r = Invoke(deck.ListOn(deck.SlideId2, null));
                if (!Ok(r, f))
                {
                    Finish(run, "ani-list-other-page", f);
                }
                else if (CountEffects(r) != 0)
                {
                    f.Add("第 2 页应空，实际 " + CountEffects(r));
                    Finish(run, "ani-list-other-page", f);
                }
                else
                {
                    Finish(run, "ani-list-other-page", f);
                }
            }

            if (Want(filters, "ani-mismatch-slide"))
            {
                deck.ClearPage();
                int before = CountEffects(Invoke(deck.List(null)));
                var a = deck.AddFlat(deck.A, "entrance", "fade", null, null, null, null, null);
                a["slide_id"] = deck.SlideId2;
                var f = new List<string>();
                ToolResult r = Invoke(a);
                if (r == null || r.Success)
                {
                    f.Add("应失败");
                }
                else if (!ContainsAny(r.Error, "不一致"))
                {
                    f.Add("文案应含不一致 实际 " + (r.Error ?? ""));
                }

                if (CountEffects(Invoke(deck.List(null))) != before)
                {
                    f.Add("时间线被改了");
                }

                Finish(run, "ani-mismatch-slide", f);
            }

            if (Want(filters, "ani-missing-shape"))
            {
                deck.ClearPage();
                int before = CountEffects(Invoke(deck.List(null)));
                string bad = "sid" + deck.SlideId + "-s99999";
                var f = new List<string>();
                ToolResult r = Invoke(deck.AddFlat(bad, "entrance", "fade", null, null, null, null, null));
                if (r == null || r.Success)
                {
                    f.Add("应失败");
                }
                else if (!ContainsAny(r.Error, "不存在"))
                {
                    f.Add("文案应含不存在 实际 " + (r.Error ?? ""));
                }

                if (CountEffects(Invoke(deck.List(null))) != before)
                {
                    f.Add("时间线被改了");
                }

                Finish(run, "ani-missing-shape", f);
            }

            if (Want(filters, "ani-appear-emphasis"))
            {
                deck.ClearPage();
                ExpectFail(
                    run,
                    "ani-appear-emphasis",
                    deck.AddFlat(deck.A, "emphasis", "appear", null, null, null, null, null),
                    "不适用于 emphasis");
            }
        }

        private static bool AbortIfDead(TestRun run, IList<string> filters, Deck deck)
        {
            if (deck == null || deck.IsLive())
            {
                return false;
            }

            FailNeeded(run, filters, "WPP 在 F_ppt_animation add 后渠道关闭（产品）");
            return true;
        }

        private static void RunReplace(TestRun run, IList<string> filters, Deck deck)
        {
            if (AbortIfDead(run, filters, deck))
            {
                return;
            }
            if (Want(filters, "ani-replace-two"))
            {
                deck.ClearPage();
                var effects = new List<Dictionary<string, object>>
                {
                    Spec(deck.A, "entrance", "fade", null, null),
                    Spec(deck.B, "entrance", "fly_in", "left", null)
                };
                var f = new List<string>();
                ToolResult r = Invoke(deck.Replace(effects));
                if (!Ok(r, f))
                {
                    Finish(run, "ani-replace-two", f);
                }
                else
                {
                    IList<Dictionary<string, object>> list = EffectsOf(r);
                    if (list.Count < 2)
                    {
                        f.Add("应有 2 条 实际 " + list.Count);
                    }
                    else if (!Eq(Get(list[0], "ShapeId"), deck.A) || !Eq(Get(list[1], "ShapeId"), deck.B))
                    {
                        f.Add("序应为 a 再 b");
                    }

                    Finish(run, "ani-replace-two", f);
                }
            }

            if (Want(filters, "ani-replace-empty"))
            {
                deck.ClearPage();
                Invoke(deck.AddFlat(deck.A, "entrance", "fade", null, null, null, null, null));
                var f = new List<string>();
                ToolResult r = Invoke(deck.Replace(new List<Dictionary<string, object>>()));
                if (!Ok(r, f))
                {
                    Finish(run, "ani-replace-empty", f);
                }
                else if (CountEffects(r) != 0)
                {
                    f.Add("应空 实际 " + CountEffects(r));
                    Finish(run, "ani-replace-empty", f);
                }
                else
                {
                    Finish(run, "ani-replace-empty", f);
                }
            }

            if (Want(filters, "ani-replace-wipes-old"))
            {
                deck.ClearPage();
                Invoke(deck.AddFlat(deck.A, "entrance", "fade", null, null, null, null, null));
                var only = new List<Dictionary<string, object>>
                {
                    Spec(deck.A, "entrance", "wipe", "right", null)
                };
                var f = new List<string>();
                ToolResult r = Invoke(deck.Replace(only));
                if (!Ok(r, f))
                {
                    Finish(run, "ani-replace-wipes-old", f);
                }
                else
                {
                    if (HasEffectName(r, "fade"))
                    {
                        f.Add("旧 fade 还在");
                    }

                    if (!HasEffectName(r, "wipe"))
                    {
                        f.Add("应只有 wipe");
                    }

                    Finish(run, "ani-replace-wipes-old", f);
                }
            }

            if (Want(filters, "ani-replace-half-fail"))
            {
                deck.ClearPage();
                Invoke(deck.AddFlat(deck.A, "entrance", "fade", null, null, null, null, null));
                int before = CountEffects(Invoke(deck.List(null)));
                string bad = "sid" + deck.SlideId + "-s99999";
                var half = new List<Dictionary<string, object>>
                {
                    Spec(deck.A, "entrance", "fade", null, null),
                    Spec(bad, "entrance", "appear", null, null)
                };
                var f = new List<string>();
                ToolResult r = Invoke(deck.Replace(half));
                if (r == null || r.Success)
                {
                    f.Add("整次应失败");
                }

                if (CountEffects(Invoke(deck.List(null))) != before)
                {
                    f.Add("旧时间线被动了");
                }

                Finish(run, "ani-replace-half-fail", f);
            }

            if (Want(filters, "ani-two-on-same"))
            {
                deck.ClearPage();
                var f = new List<string>();
                ToolResult first = Invoke(deck.AddFlat(deck.A, "entrance", "fade", null, null, null, null, null));
                ToolResult second = Invoke(deck.AddFlat(deck.A, "emphasis", "spin", null, null, null, null, null));
                if (!Ok(first, f) || !Ok(second, f))
                {
                    Finish(run, "ani-two-on-same", f);
                }
                else
                {
                    ToolResult list = Invoke(deck.List(null));
                    if (CountFor(list, deck.A) < 2)
                    {
                        f.Add("a 上应至少 2 条 实际 " + CountFor(list, deck.A));
                    }

                    if (!HasEffectName(list, "fade") || !HasEffectName(list, "spin"))
                    {
                        f.Add("a 上应同时有 fade 与 spin");
                    }

                    Finish(run, "ani-two-on-same", f);
                }
            }

            if (Want(filters, "ani-clear-page"))
            {
                deck.ClearPage();
                Invoke(deck.AddFlat(deck.A, "entrance", "fade", null, null, null, null, null));
                Invoke(deck.AddFlat(deck.B, "entrance", "appear", null, null, null, null, null));
                var f = new List<string>();
                ToolResult r = Invoke(deck.Clear(null));
                if (!Ok(r, f))
                {
                    Finish(run, "ani-clear-page", f);
                }
                else if (CountEffects(r) != 0)
                {
                    f.Add("清整页后应空");
                    Finish(run, "ani-clear-page", f);
                }
                else
                {
                    Finish(run, "ani-clear-page", f);
                }
            }

            if (Want(filters, "ani-clear-one"))
            {
                deck.ClearPage();
                Invoke(deck.AddFlat(deck.A, "entrance", "fade", null, null, null, null, null));
                Invoke(deck.AddFlat(deck.B, "entrance", "appear", null, null, null, null, null));
                var f = new List<string>();
                ToolResult r = Invoke(deck.Clear(deck.A));
                if (!Ok(r, f))
                {
                    Finish(run, "ani-clear-one", f);
                }
                else
                {
                    ToolResult list = Invoke(deck.List(null));
                    if (FindEffect(list, deck.A) != null)
                    {
                        f.Add("a 应被清");
                    }

                    if (FindEffect(list, deck.B) == null)
                    {
                        f.Add("b 应仍在");
                    }

                    Finish(run, "ani-clear-one", f);
                }
            }

            if (Want(filters, "ani-clear-then-add"))
            {
                deck.ClearPage();
                Invoke(deck.AddFlat(deck.A, "entrance", "fade", null, null, null, null, null));
                Invoke(deck.Clear(null));
                var f = CheckAdd(
                    Invoke(deck.AddFlat(deck.A, "entrance", "fade", null, null, null, null, null)),
                    deck.A,
                    "fade",
                    "entrance",
                    null,
                    1);
                if (f.Count == 0 && CountEffects(Invoke(deck.List(null))) != 1)
                {
                    f.Add("应恰好 1 条");
                }

                Finish(run, "ani-clear-then-add", f);
            }

            if (Want(filters, "ani-text-unchanged"))
            {
                if (!deck.SeedText("动画源", out string box, out string seedErr))
                {
                    run.CaseFail("ani-text-unchanged", seedErr ?? "铺框失败");
                }
                else
                {
                    Invoke(deck.AddFlat(box, "entrance", "fade", null, null, null, null, null));
                    Invoke(deck.Clear(null));
                    var f = new List<string>();
                    if (!deck.PageHasText("动画源"))
                    {
                        f.Add("文案丢了");
                    }

                    Finish(run, "ani-text-unchanged", f);
                    if (!deck.Seed(out string again))
                    {
                        FailNeeded(run, filters, again ?? "重铺失败",
                            "ani-neg-duration", "ani-neg-delay", "ani-action-upper",
                            "ani-action-spaces", "ani-shape-id-alias", "ani-echo-kind",
                            "ani-echo-slide", "ani-echo-display", "ani-add-same-slide");
                        return;
                    }
                }
            }

            if (Want(filters, "ani-neg-duration"))
            {
                deck.ClearPage();
                int before = CountEffects(Invoke(deck.List(null)));
                var f = new List<string>();
                ToolResult r = Invoke(deck.AddFlat(deck.A, "entrance", "fade", null, null, -1, null, null));
                if (r == null || r.Success)
                {
                    f.Add("应失败");
                }
                else if (!ContainsAny(r.Error, "负|duration"))
                {
                    f.Add("文案应含负/duration 实际 " + (r.Error ?? ""));
                }

                if (CountEffects(Invoke(deck.List(null))) != before)
                {
                    f.Add("未写入却改了时间线");
                }

                Finish(run, "ani-neg-duration", f);
            }

            if (Want(filters, "ani-neg-delay"))
            {
                deck.ClearPage();
                int before = CountEffects(Invoke(deck.List(null)));
                var f = new List<string>();
                ToolResult r = Invoke(deck.AddFlat(deck.A, "entrance", "fade", null, null, null, -1, null));
                if (r == null || r.Success)
                {
                    f.Add("应失败");
                }

                if (CountEffects(Invoke(deck.List(null))) != before)
                {
                    f.Add("未写入却改了时间线");
                }

                Finish(run, "ani-neg-delay", f);
            }

            if (Want(filters, "ani-action-upper"))
            {
                deck.ClearPage();
                var a = deck.AddFlat(deck.A, "entrance", "fade", null, null, null, null, null);
                a["action"] = "ADD";
                ToolResult add = Invoke(a);
                var f = CheckAdd(add, deck.A, "fade", "entrance", null, 1);
                if (f.Count == 0 && !Eq(GetString(add, "action"), "add"))
                {
                    f.Add("回包 action 应为 add 实际 " + GetString(add, "action"));
                }

                Finish(run, "ani-action-upper", f);
            }

            if (Want(filters, "ani-action-spaces"))
            {
                deck.ClearPage();
                var a = deck.AddFlat(deck.A, "entrance", "fade", null, null, null, null, null);
                a["action"] = " add ";
                ToolResult r = Invoke(a);
                var f = CheckAdd(r, deck.A, "fade", "entrance", null, 1);
                if (f.Count == 0 && !Eq(GetString(r, "action"), "add"))
                {
                    f.Add("回包 action 应为 add 实际 " + GetString(r, "action"));
                }

                Finish(run, "ani-action-spaces", f);
            }

            if (Want(filters, "ani-shape-id-alias"))
            {
                deck.ClearPage();
                var a = deck.Args();
                a["action"] = "add";
                a["shape_id"] = deck.A;
                a["category"] = "entrance";
                a["effect"] = "fade";
                a["slide_id"] = deck.SlideId;
                Finish(run, "ani-shape-id-alias", CheckAdd(Invoke(a), deck.A, "fade", "entrance", null, 1));
            }

            if (Want(filters, "ani-echo-kind"))
            {
                deck.ClearPage();
                ToolResult r = Invoke(deck.List(null));
                var f = new List<string>();
                if (!Ok(r, f))
                {
                    Finish(run, "ani-echo-kind", f);
                }
                else
                {
                    string kind = GetString(r, "kind");
                    string expect = deck.Wpp ? "wpp" : "ppt";
                    if (!Eq(kind, expect))
                    {
                        f.Add("kind 应为 " + expect + " 实际 " + kind);
                    }

                    Finish(run, "ani-echo-kind", f);
                }
            }

            if (Want(filters, "ani-echo-slide"))
            {
                deck.ClearPage();
                ToolResult r = Invoke(deck.AddFlat(deck.A, "entrance", "fade", null, null, null, null, null));
                var f = new List<string>();
                if (!Ok(r, f))
                {
                    Finish(run, "ani-echo-slide", f);
                }
                else if (!Eq(GetString(r, "slide_id"), deck.SlideId))
                {
                    f.Add("slide_id 应为 " + deck.SlideId + " 实际 " + GetString(r, "slide_id"));
                    Finish(run, "ani-echo-slide", f);
                }
                else
                {
                    Finish(run, "ani-echo-slide", f);
                }
            }

            if (Want(filters, "ani-echo-display"))
            {
                deck.ClearPage();
                ToolResult r = Invoke(deck.AddFlat(deck.A, "entrance", "fade", null, null, null, null, null));
                var f = new List<string>();
                if (!Ok(r, f))
                {
                    Finish(run, "ani-echo-display", f);
                }
                else
                {
                    string d = GetString(r, "display_contents");
                    if (d.IndexOf("action=", StringComparison.OrdinalIgnoreCase) < 0
                        || d.IndexOf("effects=", StringComparison.OrdinalIgnoreCase) < 0)
                    {
                        f.Add("display_contents 应含 action= 与 effects=");
                    }

                    Finish(run, "ani-echo-display", f);
                }
            }

            if (Want(filters, "ani-add-same-slide"))
            {
                deck.ClearPage();
                var a = deck.AddFlat(deck.A, "entrance", "fade", null, null, null, null, null);
                a["slide_id"] = deck.SlideId;
                Finish(run, "ani-add-same-slide", CheckAdd(Invoke(a), deck.A, "fade", "entrance", null, 1));
            }
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

        private static void FailNeeded(TestRun run, IList<string> filters, string err, params string[] only)
        {
            string[] ids = only != null && only.Length > 0 ? only : ComIds;
            for (int i = 0; i < ids.Length; i++)
            {
                if (Want(filters, ids[i]) && !Scored.Contains(ids[i]))
                {
                    Scored.Add(ids[i]);
                    run.CaseFail(ids[i], err ?? "铺页失败");
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

        private static Dictionary<string, object> Spec(
            string shapeId,
            string category,
            string effect,
            string dir,
            string trigger)
        {
            var d = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase)
            {
                ["ShapeId"] = shapeId,
                ["category"] = category,
                ["effect"] = effect
            };
            if (!string.IsNullOrEmpty(dir))
            {
                d["dir"] = dir;
            }

            if (!string.IsNullOrEmpty(trigger))
            {
                d["trigger"] = trigger;
            }

            return d;
        }

        private static ToolResult Invoke(Dictionary<string, object> args)
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

        private static List<string> CheckAdd(
            ToolResult r,
            string shapeId,
            string effect,
            string category,
            string dir,
            int added)
        {
            var f = new List<string>();
            if (!Ok(r, f))
            {
                return f;
            }

            if (GetIntData(r, "added_count") != added)
            {
                f.Add("added_count=" + GetIntData(r, "added_count") + " 期望 " + added);
            }

            Dictionary<string, object> e = FindEffect(r, shapeId);
            if (e == null)
            {
                f.Add("回包无 ShapeId=" + shapeId + " 实际 " + DumpEffects(r));
                return f;
            }

            if (!Eq(Get(e, "effect"), effect))
            {
                f.Add("effect 应为 " + effect + " 实际 " + Get(e, "effect"));
            }

            if (!Eq(Get(e, "category"), category))
            {
                f.Add("category 应为 " + category + " 实际 " + Get(e, "category"));
            }

            if (!string.IsNullOrEmpty(dir) && !Eq(Get(e, "dir"), dir))
            {
                f.Add("dir 应为 " + dir + " 实际 " + Get(e, "dir"));
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

        private static IList<Dictionary<string, object>> EffectsOf(ToolResult r)
        {
            var list = new List<Dictionary<string, object>>();
            var d = r != null ? r.Data as Dictionary<string, object> : null;
            if (d == null || !d.TryGetValue("effects", out object raw) || !(raw is IEnumerable en))
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

        private static int CountEffects(ToolResult r)
        {
            return EffectsOf(r).Count;
        }

        private static int CountFor(ToolResult r, string shapeId)
        {
            int n = 0;
            IList<Dictionary<string, object>> list = EffectsOf(r);
            for (int i = 0; i < list.Count; i++)
            {
                if (Eq(Get(list[i], "ShapeId"), shapeId))
                {
                    n++;
                }
            }

            return n;
        }

        private static Dictionary<string, object> FindEffect(ToolResult r, string shapeId)
        {
            IList<Dictionary<string, object>> list = EffectsOf(r);
            for (int i = 0; i < list.Count; i++)
            {
                if (Eq(Get(list[i], "ShapeId"), shapeId))
                {
                    return list[i];
                }
            }

            return null;
        }

        private static string DumpEffects(ToolResult r)
        {
            IList<Dictionary<string, object>> list = EffectsOf(r);
            if (list.Count == 0)
            {
                return "(空, added=" + GetIntData(r, "added_count")
                    + " warn=" + GetString(r, "warnings") + ")";
            }

            var sb = new System.Text.StringBuilder();
            for (int i = 0; i < list.Count; i++)
            {
                if (i > 0)
                {
                    sb.Append(", ");
                }

                sb.Append(Get(list[i], "ShapeId"))
                    .Append("/")
                    .Append(Get(list[i], "effect"))
                    .Append("/")
                    .Append(Get(list[i], "category"));
            }

            return sb.ToString();
        }

        private static bool HasEffectName(ToolResult r, string effect)
        {
            IList<Dictionary<string, object>> list = EffectsOf(r);
            for (int i = 0; i < list.Count; i++)
            {
                if (Eq(Get(list[i], "effect"), effect))
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

        private static string GetString(ToolResult r, string key)
        {
            var d = r != null ? r.Data as Dictionary<string, object> : null;
            if (d == null || !d.TryGetValue(key, out object v) || v == null)
            {
                return "";
            }

            return Convert.ToString(v) ?? "";
        }

        private static int GetIntData(ToolResult r, string key)
        {
            var d = r != null ? r.Data as Dictionary<string, object> : null;
            if (d == null || !d.TryGetValue(key, out object v) || v == null)
            {
                return 0;
            }

            try
            {
                return Convert.ToInt32(v, CultureInfo.InvariantCulture);
            }
            catch
            {
                return 0;
            }
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
                SlideId = ReadSlideId(1);
            }

            public IOperationChannel Channel { get; }

            public object Pres { get; set; }

            public bool Wpp { get; }

            public string SlideId { get; private set; }

            public string SlideId2 { get; private set; }

            public string A { get; private set; }

            public string B { get; private set; }

            public string C { get; private set; }

            public Dictionary<string, object> Args()
            {
                return new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase)
                {
                    ["channel_id"] = Channel.ChannelId
                };
            }

            public Dictionary<string, object> List(string shapeId)
            {
                return ListOn(SlideId, shapeId);
            }

            public Dictionary<string, object> ListOn(string slideId, string shapeId)
            {
                var a = Args();
                a["action"] = "list";
                a["slide_id"] = slideId;
                if (!string.IsNullOrEmpty(shapeId))
                {
                    a["ShapeId"] = shapeId;
                }

                return a;
            }

            public Dictionary<string, object> AddFlat(
                string shapeId,
                string category,
                string effect,
                string dir,
                string trigger,
                int? durationMs,
                int? delayMs,
                string slideId)
            {
                var a = Args();
                a["action"] = "add";
                a["ShapeId"] = shapeId;
                a["category"] = category;
                a["effect"] = effect;
                if (!string.IsNullOrEmpty(dir))
                {
                    a["dir"] = dir;
                }

                if (!string.IsNullOrEmpty(trigger))
                {
                    a["trigger"] = trigger;
                }

                if (durationMs.HasValue)
                {
                    a["duration_ms"] = durationMs.Value;
                }

                if (delayMs.HasValue)
                {
                    a["delay_ms"] = delayMs.Value;
                }

                if (!string.IsNullOrEmpty(slideId))
                {
                    a["slide_id"] = slideId;
                }

                return a;
            }

            public Dictionary<string, object> Replace(List<Dictionary<string, object>> effects)
            {
                var a = Args();
                a["action"] = "replace_all";
                a["slide_id"] = SlideId;
                a["effects"] = effects ?? new List<Dictionary<string, object>>();
                return a;
            }

            public Dictionary<string, object> Clear(string shapeId)
            {
                var a = Args();
                a["action"] = "clear";
                a["slide_id"] = SlideId;
                if (!string.IsNullOrEmpty(shapeId))
                {
                    a["ShapeId"] = shapeId;
                }

                return a;
            }

            public void ClearPage()
            {
                Invoke(Clear(null));
                if (!string.IsNullOrEmpty(SlideId2))
                {
                    var a = Args();
                    a["action"] = "clear";
                    a["slide_id"] = SlideId2;
                    Invoke(a);
                }
            }

            public bool Seed(out string error)
            {
                error = null;
                EnsureFirstSlide();
                SlideId = ReadSlideId(1);
                Wipe(SlideId);
                if (!TryApplyOn(SlideId, ContentHtml.TwoTextboxes("框甲", "框乙"), out List<string> ids, out error)
                    || ids.Count < 2)
                {
                    error = error ?? "两框不足";
                    return false;
                }

                A = ids[0];
                B = ids[1];
                if (!EnsureSlide2(out error))
                {
                    return false;
                }

                Wipe(SlideId2);
                if (!TryApplyOn(SlideId2, ContentHtml.TextboxCreate("第二页", 10, 10, 30, 12), out List<string> page2, out error)
                    || page2.Count < 1)
                {
                    error = error ?? "第2页无框";
                    return false;
                }

                C = page2[0];
                return true;
            }

            public bool SeedText(string text, out string box, out string error)
            {
                box = null;
                Wipe(SlideId);
                if (!TryApplyOn(SlideId, ContentHtml.TextboxCreate(text, 8, 8, 30, 12), out List<string> ids, out error)
                    || ids.Count < 1)
                {
                    error = error ?? "铺文失败";
                    return false;
                }

                box = ids[0];
                return true;
            }

            public bool IsLive()
            {
                try
                {
                    Refresh();
                    if (Wpp)
                    {
                        var w = Channel as WppChannel;
                        return w != null && w.TryGetLivePresentation(out _);
                    }

                    var p = Channel as PptChannel;
                    return p != null && p.TryGetLivePresentation(out PowerPoint.Presentation live) && live != null;
                }
                catch
                {
                    return false;
                }
            }

            public bool PageHasText(string text)
            {
                if (!TryRead(SlideId, null, out PptHtmlReadResult tree, out _) || tree == null)
                {
                    return false;
                }

                return TreeHasText(tree.Shapes, text);
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

            private void EnsureFirstSlide()
            {
                Refresh();
                if (Wpp)
                {
                    object slides = WppCom.GetProperty(Pres, "Slides");
                    int n = Convert.ToInt32(WppCom.GetProperty(slides, "Count") ?? 0);
                    if (n > 0)
                    {
                        return;
                    }

                    try
                    {
                        WppCom.Invoke(slides, "Add", 1, 12);
                    }
                    catch
                    {
                        WppCom.Invoke(slides, "Add", 1, 1);
                    }

                    return;
                }

                PowerPoint.Presentation ppt = (PowerPoint.Presentation)Pres;
                if (ppt.Slides.Count == 0)
                {
                    ppt.Slides.Add(1, PowerPoint.PpSlideLayout.ppLayoutBlank);
                }
            }

            private bool EnsureSlide2(out string error)
            {
                error = null;
                if (!string.IsNullOrEmpty(SlideId2))
                {
                    return true;
                }

                Refresh();
                try
                {
                    if (Wpp)
                    {
                        object slides = WppCom.GetProperty(Pres, "Slides");
                        int n = Convert.ToInt32(WppCom.GetProperty(slides, "Count") ?? 0);
                        object created;
                        try
                        {
                            created = WppCom.Invoke(slides, "Add", n + 1, 12);
                        }
                        catch
                        {
                            created = WppCom.Invoke(slides, "Add", n + 1, 1);
                        }

                        SlideId2 = Convert.ToString(
                            WppCom.GetProperty(created, "SlideID") ?? WppCom.GetProperty(created, "SlideId")) ?? "";
                        return !string.IsNullOrEmpty(SlideId2);
                    }

                    PowerPoint.Presentation ppt = (PowerPoint.Presentation)Pres;
                    PowerPoint.Slide s = ppt.Slides.Add(ppt.Slides.Count + 1, PowerPoint.PpSlideLayout.ppLayoutBlank);
                    SlideId2 = s.SlideID.ToString(CultureInfo.InvariantCulture);
                    return true;
                }
                catch (Exception ex)
                {
                    error = ex.Message;
                    return false;
                }
            }

            private string ReadSlideId(int index1)
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

            private bool TryApplyOn(string slideId, string html, out List<string> ids, out string error)
            {
                ids = new List<string>();
                Refresh();
                if (!ContentHtml.TryParse(html, slideId, out PptHtmlApplyPlan plan, out error))
                {
                    return false;
                }

                plan.AllowCreate = true;
                ContentHtml.ResolvePicturePaths(plan);
                PptHtmlApplyResult result;
                bool ok = Wpp
                    ? PptHtmlWppApplier.TryApply(Pres, plan, "ani-wpp", out result, out error)
                    : PptHtmlPowerPointApplier.TryApply(
                        (PowerPoint.Presentation)Pres, plan, "ani-ppt", out result, out error);
                if (!ok)
                {
                    return false;
                }

                if (result != null && result.CreatedShapes != null)
                {
                    for (int i = 0; i < result.CreatedShapes.Count; i++)
                    {
                        Dictionary<string, object> d = result.CreatedShapes[i];
                        if (d != null && d.TryGetValue("shape_id", out object v) && v != null)
                        {
                            ids.Add(v.ToString());
                        }
                    }
                }

                return true;
            }

            private bool TryRead(string slideId, string shapeId, out PptHtmlReadResult tree, out string error)
            {
                Refresh();
                if (Wpp)
                {
                    return PptHtmlWppReader.TryRead(
                        Pres, slideId, "ani-wpp", "wpp", shapeId, true, false, out tree, out error);
                }

                return PptHtmlPowerPointReader.TryRead(
                    (PowerPoint.Presentation)Pres, slideId, "ani-ppt", "ppt", shapeId, true, false, out tree, out error);
            }

            private void Wipe(string slideId)
            {
                Refresh();
                try
                {
                    if (Wpp)
                    {
                        object slide = FindSlide(slideId);
                        object shapes = WppCom.GetProperty(slide, "Shapes");
                        int n = Convert.ToInt32(WppCom.GetProperty(shapes, "Count") ?? 0);
                        for (int i = n; i >= 1; i--)
                        {
                            try
                            {
                                WppCom.Invoke(WppCom.GetIndexed(shapes, i), "Delete");
                            }
                            catch
                            {
                            }
                        }

                        return;
                    }

                    PowerPoint.Slide slidePpt = FindSlidePpt(slideId);
                    for (int i = slidePpt.Shapes.Count; i >= 1; i--)
                    {
                        try
                        {
                            slidePpt.Shapes[i].Delete();
                        }
                        catch
                        {
                        }
                    }
                }
                catch
                {
                }
            }

            private object FindSlide(string slideId)
            {
                object slides = WppCom.GetProperty(Pres, "Slides");
                int n = Convert.ToInt32(WppCom.GetProperty(slides, "Count") ?? 0);
                for (int i = 1; i <= n; i++)
                {
                    object slide = WppCom.GetIndexed(slides, i);
                    string id = Convert.ToString(
                        WppCom.GetProperty(slide, "SlideID") ?? WppCom.GetProperty(slide, "SlideId")) ?? "";
                    if (string.Equals(id, slideId, StringComparison.Ordinal))
                    {
                        return slide;
                    }
                }

                return null;
            }

            private PowerPoint.Slide FindSlidePpt(string slideId)
            {
                PowerPoint.Presentation ppt = (PowerPoint.Presentation)Pres;
                int want;
                int.TryParse(slideId, NumberStyles.Integer, CultureInfo.InvariantCulture, out want);
                for (int i = 1; i <= ppt.Slides.Count; i++)
                {
                    if (ppt.Slides[i].SlideID == want)
                    {
                        return ppt.Slides[i];
                    }
                }

                return ppt.Slides[1];
            }

            private static bool TreeHasText(IList<PptHtmlShapeNode> nodes, string text)
            {
                if (nodes == null)
                {
                    return false;
                }

                for (int i = 0; i < nodes.Count; i++)
                {
                    PptHtmlShapeNode n = nodes[i];
                    if (n == null)
                    {
                        continue;
                    }

                    if (!string.IsNullOrEmpty(n.Text)
                        && n.Text.IndexOf(text, StringComparison.Ordinal) >= 0)
                    {
                        return true;
                    }

                    if (TreeHasText(n.Children, text))
                    {
                        return true;
                    }
                }

                return false;
            }
        }
    }
}
