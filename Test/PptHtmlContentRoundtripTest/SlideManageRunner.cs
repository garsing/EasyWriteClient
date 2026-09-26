using System;
using System.Collections.Generic;
using System.Globalization;
using System.Threading.Tasks;
using WordAddIn1;
using WordAddIn1.OpenFiles;
using WordAddIn1.PresentationHost;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PptHtmlContentRoundtripTest
{
    /// <summary>F_manage_ppt_slide。前缀 sld-。</summary>
    internal static partial class SlideManageRunner
    {
        public static bool Wants(IList<string> nameFilters)
        {
            if (nameFilters == null || nameFilters.Count == 0)
            {
                return false;
            }

            for (int i = 0; i < nameFilters.Count; i++)
            {
                string f = nameFilters[i] ?? "";
                if (f.IndexOf("sld-", StringComparison.OrdinalIgnoreCase) >= 0
                    || string.Equals(f, "sld", StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
            }

            return false;
        }

        public static void Run(TestRun run, IList<string> nameFilters, bool useWpp)
        {
            Console.WriteLine(useWpp ? "管页 · WPP" : "管页 · PPT");
            RunParse(run, nameFilters);
            RunExtraParse(run, nameFilters);
            if (NeedCom(nameFilters))
            {
                RunCom(run, nameFilters, useWpp);
            }

            Console.WriteLine("管页 通过 " + run.CasesPassed
                + "  失败 " + run.CasesFailed
                + "  跳过 " + run.CasesSkipped);
        }

        private static bool NeedCom(IList<string> filters)
        {
            if (IsBroad(filters))
            {
                return true;
            }

            string[] com =
            {
                "sld-neg-bad-action", "sld-add-", "sld-dup-", "sld-del-",
                "sld-move-", "sld-reorder-", "sld-e-"
            };
            for (int i = 0; i < com.Length; i++)
            {
                if (Want(filters, com[i]))
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
                if (string.Equals(f, "sld", StringComparison.OrdinalIgnoreCase)
                    || string.Equals(f, "sld-", StringComparison.OrdinalIgnoreCase))
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
            if (fails.Count == 0)
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
            Console.WriteLine("--- A 渠道 / 互斥 ---");
            ChannelRegistry.Register(new KindStub("sld-stub-ppt", ChannelKind.Ppt), false);
            ChannelRegistry.Register(new KindStub("sld-stub-word", ChannelKind.Word), false);

            if (Want(filters, "sld-neg-no-channel"))
            {
                ExpectFail(run, "sld-neg-no-channel", new Dictionary<string, object>(), null);
            }

            if (Want(filters, "sld-neg-no-action"))
            {
                ExpectFail(
                    run,
                    "sld-neg-no-action",
                    Args("sld-stub-ppt"),
                    "add|duplicate|delete|move|reorder");
            }

            if (Want(filters, "sld-neg-wrong-kind"))
            {
                var a = Args("sld-stub-word");
                a["action"] = "add";
                ExpectFail(run, "sld-neg-wrong-kind", a, "ppt");
            }

            Foreign(run, filters, "sld-neg-add-source-channel", "add", "source_channel_id", "x", "source_channel_id");
            Foreign(run, filters, "sld-neg-add-slide-ids", "add", "slide_ids", new[] { "1" }, "slide_ids");
            Foreign(run, filters, "sld-neg-move-source-channel", "move", "source_channel_id", "x", "source_channel_id");
            Foreign(run, filters, "sld-neg-move-slide-ids", "move", "slide_ids", new[] { "1" }, "slide_ids");
            Foreign(run, filters, "sld-neg-reorder-source-channel", "reorder", "source_channel_id", "x", "source_channel_id");
            Foreign(run, filters, "sld-neg-reorder-slide-ids", "reorder", "slide_ids", new[] { "1" }, "slide_ids");
            Foreign(run, filters, "sld-neg-delete-source-channel", "delete", "source_channel_id", "x", "source_channel_id");

            if (Want(filters, "sld-neg-dup-empty-source"))
            {
                var a = Args("sld-stub-ppt");
                a["action"] = "duplicate";
                a["slide_id"] = "1";
                a["source_channel_id"] = "";
                ExpectFail(run, "sld-neg-dup-empty-source", a, "source_channel_id");
            }
        }

        private static void Foreign(
            TestRun run,
            IList<string> filters,
            string id,
            string action,
            string extra,
            object extraVal,
            string contains)
        {
            if (!Want(filters, id))
            {
                return;
            }

            var a = Args("sld-stub-ppt");
            a["action"] = action;
            a[extra] = extraVal;
            ExpectFail(run, id, a, contains);
        }

        private static void ExpectFail(TestRun run, string id, Dictionary<string, object> args, string contains)
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
            else if (contains != null
                && (r.Error ?? "").IndexOf(contains, StringComparison.OrdinalIgnoreCase) < 0)
            {
                fails.Add("文案应含「" + contains + "」实际 " + (r.Error ?? ""));
            }

            Finish(run, id, fails);
        }

        private static void RunCom(TestRun run, IList<string> filters, bool useWpp)
        {
            Console.WriteLine(useWpp ? "--- B–H COM（WPP）---" : "--- B–H COM ---");
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
            var opened = new List<PowerPoint.Presentation>();
            try
            {
                app = new PowerPoint.Application();
                app.Visible = Office.MsoTriState.msoTrue;
                try
                {
                    app.DisplayAlerts = PowerPoint.PpAlertLevel.ppAlertsNone;
                }
                catch
                {
                }

                PowerPoint.Presentation pres = app.Presentations.Add(Office.MsoTriState.msoTrue);
                opened.Add(pres);
                PptChannel ch = ChannelRegistry.CreateOrGetPpt(pres, "", false);
                if (!ch.TryGetLivePresentation(out PowerPoint.Presentation live) || live == null)
                {
                    run.CaseFail("sld-", "渠道拿不到活稿");
                    return;
                }

                var deck = new Deck(ch, live, wpp: false);
                RunBlocks(run, filters, deck, () =>
                {
                    PowerPoint.Presentation p2 = app.Presentations.Add(Office.MsoTriState.msoTrue);
                    opened.Add(p2);
                    PptChannel c2 = ChannelRegistry.CreateOrGetPpt(p2, "", false);
                    if (!c2.TryGetLivePresentation(out PowerPoint.Presentation l2) || l2 == null)
                    {
                        throw new InvalidOperationException("第二份稿渠道失败");
                    }

                    return new Deck(c2, l2, wpp: false);
                });
            }
            catch (Exception ex)
            {
                run.CaseFail("sld-", "COM 失败: " + ex.Message);
            }
            finally
            {
                for (int i = 0; i < opened.Count; i++)
                {
                    try
                    {
                        opened[i].Saved = Office.MsoTriState.msoTrue;
                        opened[i].Close();
                    }
                    catch
                    {
                    }
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

        private static void RunComWpp(TestRun run, IList<string> filters)
        {
            object app = null;
            var opened = new List<object>();
            try
            {
                if (!WppCom.TryCreateApplication(out app, out string startErr))
                {
                    run.CaseFail("sld-", startErr);
                    return;
                }

                object presentations = WppCom.GetProperty(app, "Presentations");
                object pres = TryAddPres(presentations);
                if (pres == null)
                {
                    run.CaseFail("sld-", "WPP Add 失败");
                    return;
                }

                opened.Add(pres);
                WppChannel ch = ChannelRegistry.CreateOrGetWpp(pres, "", false);
                if (!ch.TryGetLivePresentation(out object live) || live == null)
                {
                    run.CaseFail("sld-", "渠道拿不到活稿");
                    return;
                }

                var deck = new Deck(ch, live, wpp: true);
                RunBlocks(run, filters, deck, () =>
                {
                    object p2 = TryAddPres(presentations);
                    if (p2 == null)
                    {
                        throw new InvalidOperationException("WPP 第二份 Add 失败");
                    }

                    opened.Add(p2);
                    WppChannel c2 = ChannelRegistry.CreateOrGetWpp(p2, "", false);
                    if (!c2.TryGetLivePresentation(out object l2) || l2 == null)
                    {
                        throw new InvalidOperationException("第二份稿渠道失败");
                    }

                    return new Deck(c2, l2, wpp: true);
                });
            }
            catch (Exception ex)
            {
                run.CaseFail("sld-", "COM 失败: " + ex.Message);
            }
            finally
            {
                for (int i = 0; i < opened.Count; i++)
                {
                    try
                    {
                        WppCom.TrySetProperty(opened[i], "Saved", true);
                        WppCom.Invoke(opened[i], "Close");
                    }
                    catch
                    {
                    }
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

        private static object TryAddPres(object presentations)
        {
            try
            {
                return WppCom.Invoke(presentations, "Add", true);
            }
            catch
            {
                try
                {
                    return WppCom.Invoke(presentations, "Add");
                }
                catch
                {
                    return null;
                }
            }
        }

        private static void RunBlocks(
            TestRun run,
            IList<string> filters,
            Deck deck,
            Func<Deck> openSecond)
        {
            if (Want(filters, "sld-neg-bad-action"))
            {
                var a = deck.Args();
                a["action"] = "group";
                ExpectFail(run, "sld-neg-bad-action", a, "add|duplicate|delete|move|reorder");
            }

            RunAdd(run, filters, deck);
            RunDup(run, filters, deck);
            RunDel(run, filters, deck);
            RunMove(run, filters, deck);
            RunReorder(run, filters, deck);
            RunCross(run, filters, deck, openSecond);
            RunHidden(run, filters, deck);
            RunExtra(run, filters, deck, openSecond);
        }

        private static void RunAdd(TestRun run, IList<string> filters, Deck deck)
        {
            if (!Want(filters, "sld-add-"))
            {
                return;
            }

            if (!deck.EnsureCount(1, out string err))
            {
                FailAll(run, filters, err, AddIds());
                return;
            }

            string id1 = deck.Ids()[0];

            if (Want(filters, "sld-add-to-0-fail"))
            {
                int before = deck.Count();
                var a = deck.Args();
                a["action"] = "add";
                a["to_index"] = 0;
                ExpectFailUnchanged(run, "sld-add-to-0-fail", a, deck, before, "越界");
            }

            if (Want(filters, "sld-add-to-over-fail"))
            {
                if (!deck.EnsureCount(1, out err))
                {
                    run.CaseFail("sld-add-to-over-fail", err);
                }
                else
                {
                    int before = deck.Count();
                    var a = deck.Args();
                    a["action"] = "add";
                    a["to_index"] = 3;
                    ExpectFailUnchanged(run, "sld-add-to-over-fail", a, deck, before, "越界");
                }
            }

            if (Want(filters, "sld-add-layout-missing-fail"))
            {
                if (!deck.EnsureCount(1, out err))
                {
                    run.CaseFail("sld-add-layout-missing-fail", err);
                }
                else
                {
                    int before = deck.Count();
                    var a = deck.Args();
                    a["action"] = "add";
                    a["layout"] = "__no_such_layout__";
                    ExpectFailUnchanged(run, "sld-add-layout-missing-fail", a, deck, before, "版式");
                }
            }

            if (Want(filters, "sld-add-tail")
                || Want(filters, "sld-add-returns-slides")
                || Want(filters, "sld-add-hidden-false"))
            {
                if (!deck.EnsureCount(1, out err))
                {
                    FailNeeded(run, filters, err, "sld-add-tail", "sld-add-returns-slides", "sld-add-hidden-false");
                }
                else
                {
                    id1 = deck.Ids()[0];
                    int before = deck.Count();
                    var a = deck.Args();
                    a["action"] = "add";
                    ToolResult r = Invoke(a);
                    if (Want(filters, "sld-add-tail"))
                    {
                        AssertAddTail(run, r, deck, before, id1);
                    }

                    if (Want(filters, "sld-add-returns-slides"))
                    {
                        AssertReturnsSlides(run, r);
                    }

                    if (Want(filters, "sld-add-hidden-false"))
                    {
                        AssertNewHiddenFalse(run, r);
                    }
                }
            }

            if (Want(filters, "sld-add-to-1"))
            {
                if (!deck.EnsureCount(1, out err))
                {
                    run.CaseFail("sld-add-to-1", err);
                }
                else
                {
                    id1 = deck.Ids()[0];
                    var a = deck.Args();
                    a["action"] = "add";
                    a["to_index"] = 1;
                    ToolResult r = Invoke(a);
                    var fails = new List<string>();
                    if (!Ok(r, fails) || !Align(r, deck, fails))
                    {
                        Finish(run, "sld-add-to-1", fails);
                    }
                    else
                    {
                        IList<SlideRow> rows = Rows(r);
                        if (rows.Count < 2 || rows[0].Index != 1
                            || string.Equals(rows[0].SlideId, id1, StringComparison.Ordinal))
                        {
                            fails.Add("新页应在 index=1");
                        }

                        if (IndexOf(rows, id1) != 2)
                        {
                            fails.Add("原首页应变 2");
                        }

                        Finish(run, "sld-add-to-1", fails);
                    }
                }
            }

            if (Want(filters, "sld-add-to-2"))
            {
                if (!deck.EnsureCount(2, out err))
                {
                    run.CaseFail("sld-add-to-2", err);
                }
                else
                {
                    string old2 = deck.Ids()[1];
                    var a = deck.Args();
                    a["action"] = "add";
                    a["to_index"] = 2;
                    ToolResult r = Invoke(a);
                    var fails = new List<string>();
                    if (!Ok(r, fails) || !Align(r, deck, fails))
                    {
                        Finish(run, "sld-add-to-2", fails);
                    }
                    else
                    {
                        IList<SlideRow> rows = Rows(r);
                        if (GetString(r, "slide_id").Length == 0 || IndexOf(rows, GetString(r, "slide_id")) != 2)
                        {
                            fails.Add("新页应在 index=2");
                        }

                        if (IndexOf(rows, old2) != 3)
                        {
                            fails.Add("原第 2 页应变 3");
                        }

                        Finish(run, "sld-add-to-2", fails);
                    }
                }
            }

            if (Want(filters, "sld-add-layout-ok"))
            {
                if (!deck.EnsureCount(1, out err))
                {
                    run.CaseFail("sld-add-layout-ok", err);
                }
                else
                {
                    var seed = deck.Args();
                    seed["action"] = "add";
                    ToolResult seedR = Invoke(seed);
                    string layout = "";
                    IList<SlideRow> seedRows = Rows(seedR);
                    if (seedRows.Count > 0)
                    {
                        layout = seedRows[seedRows.Count - 1].Layout ?? "";
                    }

                    var a = deck.Args();
                    a["action"] = "add";
                    if (!string.IsNullOrWhiteSpace(layout))
                    {
                        a["layout"] = layout;
                    }

                    ToolResult r = Invoke(a);
                    var fails = new List<string>();
                    if (!Ok(r, fails) || !Align(r, deck, fails))
                    {
                        Finish(run, "sld-add-layout-ok", fails);
                    }
                    else if (!string.IsNullOrWhiteSpace(layout))
                    {
                        string got = "";
                        IList<SlideRow> rows = Rows(r);
                        string nid = GetString(r, "slide_id");
                        for (int i = 0; i < rows.Count; i++)
                        {
                            if (string.Equals(rows[i].SlideId, nid, StringComparison.Ordinal))
                            {
                                got = rows[i].Layout ?? "";
                                break;
                            }
                        }

                        if (!string.Equals(got, layout, StringComparison.OrdinalIgnoreCase))
                        {
                            fails.Add("layout=" + got + " 期望 " + layout);
                        }

                        Finish(run, "sld-add-layout-ok", fails);
                    }
                    else
                    {
                        Finish(run, "sld-add-layout-ok", fails);
                    }
                }
            }
        }

        private static string[] AddIds()
        {
            return new[]
            {
                "sld-add-tail", "sld-add-to-1", "sld-add-to-2", "sld-add-to-0-fail",
                "sld-add-to-over-fail", "sld-add-layout-ok", "sld-add-layout-missing-fail",
                "sld-add-returns-slides", "sld-add-hidden-false"
            };
        }

        private static void RunDup(TestRun run, IList<string> filters, Deck deck)
        {
            if (!Want(filters, "sld-dup-tail")
                && !Want(filters, "sld-dup-to-1")
                && !Want(filters, "sld-dup-ids-")
                && !Want(filters, "sld-dup-missing")
                && !Want(filters, "sld-dup-empty")
                && !Want(filters, "sld-dup-keeps")
                && !Want(filters, "sld-dup-to-over")
                && !Want(filters, "sld-dup-cross"))
            {
                if (!Want(filters, "sld-dup-"))
                {
                    return;
                }
            }

            bool anySame = Want(filters, "sld-dup-tail")
                || Want(filters, "sld-dup-to-1")
                || Want(filters, "sld-dup-ids-two")
                || Want(filters, "sld-dup-ids-repeat")
                || Want(filters, "sld-dup-ids-wins")
                || Want(filters, "sld-dup-missing-fail")
                || Want(filters, "sld-dup-empty-ids-fail")
                || Want(filters, "sld-dup-keeps-content")
                || Want(filters, "sld-dup-to-over-fail");
            if (!anySame)
            {
                return;
            }

            if (!deck.EnsureCount(3, out string err))
            {
                FailNeeded(
                    run,
                    filters,
                    err,
                    "sld-dup-tail",
                    "sld-dup-to-1",
                    "sld-dup-ids-two",
                    "sld-dup-ids-repeat",
                    "sld-dup-ids-wins",
                    "sld-dup-missing-fail",
                    "sld-dup-empty-ids-fail",
                    "sld-dup-keeps-content",
                    "sld-dup-to-over-fail");
                return;
            }

            IList<string> abc = deck.Ids();
            string a = abc[0];
            string b = abc[1];
            string c = abc[2];

            if (Want(filters, "sld-dup-missing-fail"))
            {
                ExpectFailUnchanged(
                    run,
                    "sld-dup-missing-fail",
                    Dup(deck, "999999"),
                    deck,
                    3,
                    "不存在");
            }

            if (Want(filters, "sld-dup-empty-ids-fail"))
            {
                var e = deck.Args();
                e["action"] = "duplicate";
                e["slide_ids"] = new string[0];
                ExpectFailUnchanged(run, "sld-dup-empty-ids-fail", e, deck, deck.Count(), "不能为空");
            }

            if (Want(filters, "sld-dup-to-over-fail"))
            {
                if (!deck.EnsureCount(3, out err))
                {
                    run.CaseFail("sld-dup-to-over-fail", err);
                }
                else
                {
                    var o = Dup(deck, a);
                    o["to_index"] = 5;
                    ExpectFailUnchanged(run, "sld-dup-to-over-fail", o, deck, 3, "越界");
                }
            }

            if (Want(filters, "sld-dup-tail"))
            {
                if (!deck.EnsureCount(3, out err))
                {
                    run.CaseFail("sld-dup-tail", err);
                }
                else
                {
                    a = deck.Ids()[0];
                    var d = Dup(deck, a);
                    ToolResult r = Invoke(d);
                    var fails = new List<string>();
                    if (!Ok(r, fails) || !Align(r, deck, fails))
                    {
                        Finish(run, "sld-dup-tail", fails);
                    }
                    else
                    {
                        if (GetInt(r, "slide_count") != 4)
                        {
                            fails.Add("应为 4 页");
                        }

                        string nid = GetString(r, "slide_id");
                        if (string.Equals(nid, a, StringComparison.Ordinal) || IndexOf(Rows(r), nid) != 4)
                        {
                            fails.Add("新页应在末尾且 id≠源");
                        }

                        IList<CreatedRow> cr = Created(r);
                        if (cr.Count != 1 || !string.Equals(cr[0].Source, a, StringComparison.Ordinal))
                        {
                            fails.Add("created 应为 1 条源 a");
                        }

                        Finish(run, "sld-dup-tail", fails);
                    }
                }
            }

            if (Want(filters, "sld-dup-to-1"))
            {
                if (!deck.EnsureCount(3, out err))
                {
                    run.CaseFail("sld-dup-to-1", err);
                }
                else
                {
                    c = deck.Ids()[2];
                    var d = Dup(deck, c);
                    d["to_index"] = 1;
                    ToolResult r = Invoke(d);
                    var fails = new List<string>();
                    if (!Ok(r, fails) || !Align(r, deck, fails))
                    {
                        Finish(run, "sld-dup-to-1", fails);
                    }
                    else
                    {
                        if (IndexOf(Rows(r), GetString(r, "slide_id")) != 1)
                        {
                            fails.Add("新页应在 1");
                        }

                        if (IndexOf(Rows(r), c) < 1)
                        {
                            fails.Add("源 c 应仍在");
                        }

                        Finish(run, "sld-dup-to-1", fails);
                    }
                }
            }

            if (Want(filters, "sld-dup-ids-two"))
            {
                if (!deck.EnsureCount(3, out err))
                {
                    run.CaseFail("sld-dup-ids-two", err);
                }
                else
                {
                    abc = deck.Ids();
                    a = abc[0];
                    c = abc[2];
                    int before = deck.Count();
                    var d = deck.Args();
                    d["action"] = "duplicate";
                    d["slide_ids"] = new[] { a, c };
                    ToolResult r = Invoke(d);
                    var fails = new List<string>();
                    if (!Ok(r, fails) || !Align(r, deck, fails))
                    {
                        Finish(run, "sld-dup-ids-two", fails);
                    }
                    else
                    {
                        if (GetInt(r, "slide_count") != before + 2)
                        {
                            fails.Add("应 +2");
                        }

                        IList<CreatedRow> cr = Created(r);
                        if (cr.Count != 2
                            || !string.Equals(cr[0].Source, a, StringComparison.Ordinal)
                            || !string.Equals(cr[1].Source, c, StringComparison.Ordinal))
                        {
                            fails.Add("created 源序应为 a,c");
                        }

                        if (!string.Equals(GetString(r, "slide_id"), cr.Count == 2 ? cr[1].NewId : "", StringComparison.Ordinal))
                        {
                            fails.Add("焦点应为第二张新页");
                        }

                        Finish(run, "sld-dup-ids-two", fails);
                    }
                }
            }

            if (Want(filters, "sld-dup-ids-repeat"))
            {
                if (!deck.EnsureCount(3, out err))
                {
                    run.CaseFail("sld-dup-ids-repeat", err);
                }
                else
                {
                    a = deck.Ids()[0];
                    var d = deck.Args();
                    d["action"] = "duplicate";
                    d["slide_ids"] = new[] { a, a };
                    ToolResult r = Invoke(d);
                    var fails = new List<string>();
                    if (!Ok(r, fails) || !Align(r, deck, fails))
                    {
                        Finish(run, "sld-dup-ids-repeat", fails);
                    }
                    else
                    {
                        IList<CreatedRow> cr = Created(r);
                        if (cr.Count != 2 || string.Equals(cr[0].NewId, cr[1].NewId, StringComparison.Ordinal))
                        {
                            fails.Add("两张新 id 应不同");
                        }

                        Finish(run, "sld-dup-ids-repeat", fails);
                    }
                }
            }

            if (Want(filters, "sld-dup-ids-wins"))
            {
                if (!deck.EnsureCount(3, out err))
                {
                    run.CaseFail("sld-dup-ids-wins", err);
                }
                else
                {
                    abc = deck.Ids();
                    a = abc[0];
                    b = abc[1];
                    var d = deck.Args();
                    d["action"] = "duplicate";
                    d["slide_id"] = b;
                    d["slide_ids"] = new[] { a };
                    ToolResult r = Invoke(d);
                    var fails = new List<string>();
                    if (!Ok(r, fails) || !Align(r, deck, fails))
                    {
                        Finish(run, "sld-dup-ids-wins", fails);
                    }
                    else
                    {
                        IList<CreatedRow> cr = Created(r);
                        if (cr.Count != 1 || !string.Equals(cr[0].Source, a, StringComparison.Ordinal))
                        {
                            fails.Add("应以 slide_ids 为准（源 a）");
                        }

                        Finish(run, "sld-dup-ids-wins", fails);
                    }
                }
            }

            if (Want(filters, "sld-dup-keeps-content"))
            {
                if (!deck.EnsureCount(3, out err) || !deck.TryApplyText(deck.Ids()[0], "管页源", out err))
                {
                    run.CaseFail("sld-dup-keeps-content", err ?? "铺字失败");
                }
                else
                {
                    a = deck.Ids()[0];
                    ToolResult r = Invoke(Dup(deck, a));
                    var fails = new List<string>();
                    if (!Ok(r, fails))
                    {
                        Finish(run, "sld-dup-keeps-content", fails);
                    }
                    else
                    {
                        string nid = GetString(r, "slide_id");
                        if (!deck.PageHasText(nid, "管页源", out err))
                        {
                            fails.Add(err ?? "新页无「管页源」");
                        }

                        if (IndexOf(Rows(r), a) < 1)
                        {
                            fails.Add("源页应仍在");
                        }

                        Finish(run, "sld-dup-keeps-content", fails);
                    }
                }
            }
        }

        private static void RunDel(TestRun run, IList<string> filters, Deck deck)
        {
            if (!Want(filters, "sld-del-"))
            {
                return;
            }

            if (Want(filters, "sld-del-confirm-missing") || Want(filters, "sld-del-confirm-false")
                || Want(filters, "sld-del-one") || Want(filters, "sld-del-ids-dup-fail")
                || Want(filters, "sld-del-missing-fail"))
            {
                if (!deck.EnsureCount(3, out string err3))
                {
                    FailNeeded(
                        run,
                        filters,
                        err3,
                        "sld-del-one",
                        "sld-del-confirm-missing",
                        "sld-del-confirm-false",
                        "sld-del-ids-dup-fail",
                        "sld-del-missing-fail");
                }
                else
                {
                    IList<string> abc = deck.Ids();
                    string a = abc[0];
                    string b = abc[1];
                    if (Want(filters, "sld-del-confirm-missing"))
                    {
                        var m = deck.Args();
                        m["action"] = "delete";
                        m["slide_id"] = b;
                        ExpectFailUnchanged(run, "sld-del-confirm-missing", m, deck, 3, "confirm");
                    }

                    if (Want(filters, "sld-del-confirm-false"))
                    {
                        var m = deck.Args();
                        m["action"] = "delete";
                        m["slide_id"] = b;
                        m["confirm"] = false;
                        ExpectFailUnchanged(run, "sld-del-confirm-false", m, deck, 3, "confirm");
                    }

                    if (Want(filters, "sld-del-missing-fail"))
                    {
                        var m = Del(deck, "999999");
                        ExpectFailUnchanged(run, "sld-del-missing-fail", m, deck, 3, "不存在");
                    }

                    if (Want(filters, "sld-del-ids-dup-fail"))
                    {
                        var m = deck.Args();
                        m["action"] = "delete";
                        m["confirm"] = true;
                        m["slide_ids"] = new[] { a, a };
                        ExpectFailUnchanged(run, "sld-del-ids-dup-fail", m, deck, 3, "重复");
                    }

                    if (Want(filters, "sld-del-one"))
                    {
                        if (!deck.EnsureCount(3, out err3))
                        {
                            run.CaseFail("sld-del-one", err3);
                        }
                        else
                        {
                            b = deck.Ids()[1];
                            ToolResult r = Invoke(Del(deck, b));
                            var fails = new List<string>();
                            if (!Ok(r, fails) || !Align(r, deck, fails))
                            {
                                Finish(run, "sld-del-one", fails);
                            }
                            else
                            {
                                if (GetInt(r, "slide_count") != 2 || IndexOf(Rows(r), b) >= 0)
                                {
                                    fails.Add("应剩 2 页且无 b");
                                }

                                if (!DeletedHas(r, b) || !string.Equals(GetString(r, "slide_id"), b, StringComparison.Ordinal))
                                {
                                    fails.Add("deleted/焦点应为 b");
                                }

                                Finish(run, "sld-del-one", fails);
                            }
                        }
                    }
                }
            }

            if (Want(filters, "sld-del-last-fail"))
            {
                if (!deck.EnsureCount(1, out string err))
                {
                    run.CaseFail("sld-del-last-fail", err);
                }
                else
                {
                    ExpectFailUnchanged(run, "sld-del-last-fail", Del(deck, deck.Ids()[0]), deck, 1, "唯一");
                }
            }

            if (Want(filters, "sld-del-ids-two"))
            {
                if (!deck.EnsureCount(4, out string err))
                {
                    run.CaseFail("sld-del-ids-two", err);
                }
                else
                {
                    IList<string> ids = deck.Ids();
                    string b = ids[1];
                    string c = ids[2];
                    var m = deck.Args();
                    m["action"] = "delete";
                    m["confirm"] = true;
                    m["slide_ids"] = new[] { b, c };
                    ToolResult r = Invoke(m);
                    var fails = new List<string>();
                    if (!Ok(r, fails) || !Align(r, deck, fails))
                    {
                        Finish(run, "sld-del-ids-two", fails);
                    }
                    else
                    {
                        if (GetInt(r, "slide_count") != 2)
                        {
                            fails.Add("应剩 2");
                        }

                        IList<string> del = DeletedList(r);
                        if (del.Count != 2 || del[0] != b || del[1] != c
                            || !string.Equals(GetString(r, "slide_id"), c, StringComparison.Ordinal))
                        {
                            fails.Add("deleted 序应为 b,c，焦点 c");
                        }

                        Finish(run, "sld-del-ids-two", fails);
                    }
                }
            }

            if (Want(filters, "sld-del-wipe-fail"))
            {
                if (!deck.EnsureCount(2, out string err))
                {
                    run.CaseFail("sld-del-wipe-fail", err);
                }
                else
                {
                    IList<string> ids = deck.Ids();
                    var m = deck.Args();
                    m["action"] = "delete";
                    m["confirm"] = true;
                    m["slide_ids"] = new[] { ids[0], ids[1] };
                    ExpectFailUnchanged(run, "sld-del-wipe-fail", m, deck, 2, "唯一");
                }
            }
        }

        private static void RunMove(TestRun run, IList<string> filters, Deck deck)
        {
            if (!Want(filters, "sld-move-"))
            {
                return;
            }

            if (!deck.EnsureCount(3, out string err))
            {
                FailNeeded(
                    run,
                    filters,
                    err,
                    "sld-move-to-end",
                    "sld-move-to-1",
                    "sld-move-same-ok",
                    "sld-move-no-to-fail",
                    "sld-move-to-0-fail",
                    "sld-move-to-over-fail",
                    "sld-move-missing-fail");
                return;
            }

            IList<string> abc = deck.Ids();
            string a = abc[0];
            string b = abc[1];
            string c = abc[2];

            if (Want(filters, "sld-move-no-to-fail"))
            {
                var m = deck.Args();
                m["action"] = "move";
                m["slide_id"] = a;
                ExpectFailUnchanged(run, "sld-move-no-to-fail", m, deck, 3, "to_index");
            }

            if (Want(filters, "sld-move-to-0-fail"))
            {
                var m = Move(deck, a, 0);
                ExpectFailUnchanged(run, "sld-move-to-0-fail", m, deck, 3, "越界");
            }

            if (Want(filters, "sld-move-to-over-fail"))
            {
                var m = Move(deck, a, 4);
                ExpectFailUnchanged(run, "sld-move-to-over-fail", m, deck, 3, "越界");
            }

            if (Want(filters, "sld-move-missing-fail"))
            {
                ExpectFailUnchanged(run, "sld-move-missing-fail", Move(deck, "999999", 1), deck, 3, "不存在");
            }

            if (Want(filters, "sld-move-same-ok"))
            {
                if (!deck.EnsureCount(3, out err))
                {
                    run.CaseFail("sld-move-same-ok", err);
                }
                else
                {
                    abc = deck.Ids();
                    ToolResult r = Invoke(Move(deck, abc[1], 2));
                    var fails = new List<string>();
                    if (!Ok(r, fails) || !Align(r, deck, fails))
                    {
                        Finish(run, "sld-move-same-ok", fails);
                    }
                    else if (!OrderIs(Rows(r), abc[0], abc[1], abc[2]))
                    {
                        fails.Add("同位应保持 a,b,c");
                        Finish(run, "sld-move-same-ok", fails);
                    }
                    else
                    {
                        Finish(run, "sld-move-same-ok", fails);
                    }
                }
            }

            if (Want(filters, "sld-move-to-end"))
            {
                if (!deck.EnsureCount(3, out err))
                {
                    run.CaseFail("sld-move-to-end", err);
                }
                else
                {
                    abc = deck.Ids();
                    a = abc[0];
                    ToolResult r = Invoke(Move(deck, a, 3));
                    var fails = new List<string>();
                    if (!Ok(r, fails) || !Align(r, deck, fails))
                    {
                        Finish(run, "sld-move-to-end", fails);
                    }
                    else if (!OrderIs(Rows(r), abc[1], abc[2], a))
                    {
                        fails.Add("顺序应为 b,c,a");
                        Finish(run, "sld-move-to-end", fails);
                    }
                    else
                    {
                        Finish(run, "sld-move-to-end", fails);
                    }
                }
            }

            if (Want(filters, "sld-move-to-1"))
            {
                if (!deck.EnsureCount(3, out err))
                {
                    run.CaseFail("sld-move-to-1", err);
                }
                else
                {
                    abc = deck.Ids();
                    c = abc[2];
                    ToolResult r = Invoke(Move(deck, c, 1));
                    var fails = new List<string>();
                    if (!Ok(r, fails) || !Align(r, deck, fails))
                    {
                        Finish(run, "sld-move-to-1", fails);
                    }
                    else if (!OrderIs(Rows(r), c, abc[0], abc[1]))
                    {
                        fails.Add("顺序应为 c,a,b");
                        Finish(run, "sld-move-to-1", fails);
                    }
                    else
                    {
                        Finish(run, "sld-move-to-1", fails);
                    }
                }
            }
        }

        private static void RunReorder(TestRun run, IList<string> filters, Deck deck)
        {
            if (!Want(filters, "sld-reorder-"))
            {
                return;
            }

            if (!deck.EnsureCount(3, out string err))
            {
                FailNeeded(
                    run,
                    filters,
                    err,
                    "sld-reorder-reverse",
                    "sld-reorder-identity",
                    "sld-reorder-no-order",
                    "sld-reorder-short-fail",
                    "sld-reorder-dup-fail",
                    "sld-reorder-unknown-fail",
                    "sld-reorder-empty-item-fail");
                return;
            }

            IList<string> abc = deck.Ids();
            string a = abc[0];
            string b = abc[1];
            string c = abc[2];

            if (Want(filters, "sld-reorder-no-order"))
            {
                var m = deck.Args();
                m["action"] = "reorder";
                ExpectFailUnchanged(run, "sld-reorder-no-order", m, deck, 3, "order");
            }

            if (Want(filters, "sld-reorder-short-fail"))
            {
                var m = deck.Args();
                m["action"] = "reorder";
                m["order"] = new[] { a, b };
                ExpectFailUnchanged(run, "sld-reorder-short-fail", m, deck, 3, "页数");
            }

            if (Want(filters, "sld-reorder-dup-fail"))
            {
                var m = deck.Args();
                m["action"] = "reorder";
                m["order"] = new[] { a, a, b };
                ExpectFailUnchanged(run, "sld-reorder-dup-fail", m, deck, 3, "重复");
            }

            if (Want(filters, "sld-reorder-unknown-fail"))
            {
                var m = deck.Args();
                m["action"] = "reorder";
                m["order"] = new[] { a, b, "999999" };
                ExpectFailUnchanged(run, "sld-reorder-unknown-fail", m, deck, 3, "未知");
            }

            if (Want(filters, "sld-reorder-empty-item-fail"))
            {
                var m = deck.Args();
                m["action"] = "reorder";
                m["order"] = new[] { a, "", c };
                ExpectFailUnchanged(run, "sld-reorder-empty-item-fail", m, deck, 3, "空");
            }

            if (Want(filters, "sld-reorder-identity"))
            {
                if (!deck.EnsureCount(3, out err))
                {
                    run.CaseFail("sld-reorder-identity", err);
                }
                else
                {
                    abc = deck.Ids();
                    var m = deck.Args();
                    m["action"] = "reorder";
                    m["order"] = new[] { abc[0], abc[1], abc[2] };
                    ToolResult r = Invoke(m);
                    var fails = new List<string>();
                    if (!Ok(r, fails) || !Align(r, deck, fails) || !OrderIs(Rows(r), abc[0], abc[1], abc[2]))
                    {
                        if (fails.Count == 0)
                        {
                            fails.Add("顺序应不变");
                        }

                        Finish(run, "sld-reorder-identity", fails);
                    }
                    else
                    {
                        Finish(run, "sld-reorder-identity", fails);
                    }
                }
            }

            if (Want(filters, "sld-reorder-reverse"))
            {
                if (!deck.EnsureCount(3, out err))
                {
                    run.CaseFail("sld-reorder-reverse", err);
                }
                else
                {
                    abc = deck.Ids();
                    var m = deck.Args();
                    m["action"] = "reorder";
                    m["order"] = new[] { abc[2], abc[1], abc[0] };
                    ToolResult r = Invoke(m);
                    var fails = new List<string>();
                    if (!Ok(r, fails) || !Align(r, deck, fails) || !OrderIs(Rows(r), abc[2], abc[1], abc[0]))
                    {
                        if (fails.Count == 0)
                        {
                            fails.Add("应为 c,b,a");
                        }

                        Finish(run, "sld-reorder-reverse", fails);
                    }
                    else
                    {
                        Finish(run, "sld-reorder-reverse", fails);
                    }
                }
            }
        }

        private static void RunCross(TestRun run, IList<string> filters, Deck dest, Func<Deck> openSecond)
        {
            bool need = Want(filters, "sld-dup-cross-one")
                || Want(filters, "sld-dup-cross-seq")
                || Want(filters, "sld-dup-cross-host-fail")
                || Want(filters, "sld-dup-unknown-source");
            if (!need)
            {
                return;
            }

            Deck src;
            try
            {
                src = openSecond();
            }
            catch (Exception ex)
            {
                FailNeeded(
                    run,
                    filters,
                    ex.Message,
                    "sld-dup-cross-one",
                    "sld-dup-cross-seq",
                    "sld-dup-cross-host-fail",
                    "sld-dup-unknown-source");
                return;
            }

            if (!src.EnsureCount(2, out string err) || !dest.EnsureCount(1, out err))
            {
                FailNeeded(run, filters, err, "sld-dup-cross-one", "sld-dup-cross-seq");
                return;
            }

            IList<string> sids = src.Ids();
            string s1 = sids[0];
            string s2 = sids[1];

            if (Want(filters, "sld-dup-unknown-source"))
            {
                int dc = dest.Count();
                int sc = src.Count();
                var a = dest.Args();
                a["action"] = "duplicate";
                a["slide_id"] = s1;
                a["source_channel_id"] = "no-such";
                ToolResult r = Invoke(a);
                var fails = new List<string>();
                if (r != null && r.Success)
                {
                    fails.Add("应失败");
                }
                else if ((r == null ? "" : r.Error ?? "").IndexOf("source_channel_id", StringComparison.OrdinalIgnoreCase) < 0)
                {
                    fails.Add("文案应含未知 source_channel_id");
                }

                if (dest.Count() != dc || src.Count() != sc)
                {
                    fails.Add("页数被改了");
                }

                Finish(run, "sld-dup-unknown-source", fails);
            }

            if (Want(filters, "sld-dup-cross-host-fail"))
            {
                ChannelRegistry.Register(
                    new KindStub("sld-stub-other-host", dest.Wpp ? ChannelKind.Ppt : ChannelKind.Wpp),
                    false);
                int dc = dest.Count();
                int sc = src.Count();
                var a = dest.Args();
                a["action"] = "duplicate";
                a["slide_id"] = s1;
                a["source_channel_id"] = "sld-stub-other-host";
                ToolResult r = Invoke(a);
                var fails = new List<string>();
                if (r != null && r.Success)
                {
                    fails.Add("应失败");
                }
                else if ((r == null ? "" : r.Error ?? "").IndexOf("同为", StringComparison.Ordinal) < 0)
                {
                    fails.Add("文案应含同为 ppt/wpp，实际 " + (r == null ? "" : r.Error));
                }

                if (dest.Count() != dc || src.Count() != sc)
                {
                    fails.Add("页数被改了");
                }

                Finish(run, "sld-dup-cross-host-fail", fails);
            }

            if (Want(filters, "sld-dup-cross-one"))
            {
                if (!dest.EnsureCount(1, out err) || !src.EnsureCount(2, out err))
                {
                    run.CaseFail("sld-dup-cross-one", err);
                }
                else
                {
                    s1 = src.Ids()[0];
                    int sc = src.Count();
                    int dc = dest.Count();
                    var a = dest.Args();
                    a["action"] = "duplicate";
                    a["slide_id"] = s1;
                    a["source_channel_id"] = src.Channel.ChannelId;
                    ToolResult r = Invoke(a);
                    var fails = new List<string>();
                    if (!Ok(r, fails) || !Align(r, dest, fails))
                    {
                        Finish(run, "sld-dup-cross-one", fails);
                    }
                    else
                    {
                        if (dest.Count() != dc + 1 || src.Count() != sc)
                        {
                            fails.Add("目标应 +1，源不变");
                        }

                        IList<CreatedRow> cr = Created(r);
                        if (cr.Count < 1 || !string.Equals(cr[0].Source, s1, StringComparison.Ordinal))
                        {
                            fails.Add("created 源应为 s1");
                        }

                        if (!string.Equals(
                                GetString(r, "source_channel_id"),
                                ChannelRegistry.ToPublicId(src.Channel.ChannelId),
                                StringComparison.OrdinalIgnoreCase)
                            && !string.Equals(
                                GetString(r, "source_channel_id"),
                                src.Channel.ChannelId,
                                StringComparison.OrdinalIgnoreCase))
                        {
                            fails.Add("应回显 source_channel_id");
                        }

                        Finish(run, "sld-dup-cross-one", fails);
                    }
                }
            }

            if (Want(filters, "sld-dup-cross-seq"))
            {
                if (!dest.EnsureCount(1, out err) || !src.EnsureCount(2, out err))
                {
                    run.CaseFail("sld-dup-cross-seq", err);
                }
                else
                {
                    sids = src.Ids();
                    s1 = sids[0];
                    s2 = sids[1];
                    int sc = src.Count();
                    int dc = dest.Count();
                    var a = dest.Args();
                    a["action"] = "duplicate";
                    a["slide_ids"] = new[] { s2, s1 };
                    a["source_channel_id"] = src.Channel.ChannelId;
                    ToolResult r = Invoke(a);
                    var fails = new List<string>();
                    if (!Ok(r, fails) || !Align(r, dest, fails))
                    {
                        Finish(run, "sld-dup-cross-seq", fails);
                    }
                    else
                    {
                        if (dest.Count() != dc + 2 || src.Count() != sc)
                        {
                            fails.Add("目标应 +2，源仍 2");
                        }

                        IList<CreatedRow> cr = Created(r);
                        if (cr.Count != 2
                            || !string.Equals(cr[0].Source, s2, StringComparison.Ordinal)
                            || !string.Equals(cr[1].Source, s1, StringComparison.Ordinal))
                        {
                            fails.Add("created 源序应为 s2,s1");
                        }

                        Finish(run, "sld-dup-cross-seq", fails);
                    }
                }
            }
        }

        private static void RunHidden(TestRun run, IList<string> filters, Deck deck)
        {
            if (!Want(filters, "sld-dup-hidden-ok") && !Want(filters, "sld-del-hidden-ok"))
            {
                return;
            }

            if (!deck.EnsureCount(3, out string err) || !deck.TryHide(2, true, out err))
            {
                FailNeeded(run, filters, err ?? "隐藏失败", "sld-dup-hidden-ok", "sld-del-hidden-ok");
                return;
            }

            string hid = deck.Ids()[1];
            if (Want(filters, "sld-dup-hidden-ok"))
            {
                int before = deck.Count();
                ToolResult r = Invoke(Dup(deck, hid));
                var fails = new List<string>();
                if (!Ok(r, fails) || GetInt(r, "slide_count") != before + 1)
                {
                    if (fails.Count == 0)
                    {
                        fails.Add("隐藏页应能拷，+1");
                    }

                    Finish(run, "sld-dup-hidden-ok", fails);
                }
                else
                {
                    Finish(run, "sld-dup-hidden-ok", fails);
                }
            }

            if (Want(filters, "sld-del-hidden-ok"))
            {
                if (!deck.EnsureCount(3, out err) || !deck.TryHide(2, true, out err))
                {
                    run.CaseFail("sld-del-hidden-ok", err ?? "隐藏失败");
                }
                else
                {
                    hid = deck.Ids()[1];
                    ToolResult r = Invoke(Del(deck, hid));
                    var fails = new List<string>();
                    if (!Ok(r, fails) || IndexOf(Rows(r), hid) >= 0)
                    {
                        if (fails.Count == 0)
                        {
                            fails.Add("隐藏页应能删");
                        }

                        Finish(run, "sld-del-hidden-ok", fails);
                    }
                    else
                    {
                        Finish(run, "sld-del-hidden-ok", fails);
                    }
                }
            }
        }

        private static void AssertAddTail(TestRun run, ToolResult r, Deck deck, int before, string id1)
        {
            var fails = new List<string>();
            if (!Ok(r, fails) || !Align(r, deck, fails))
            {
                Finish(run, "sld-add-tail", fails);
                return;
            }

            if (!string.Equals(GetString(r, "action"), "add", StringComparison.Ordinal)
                || GetInt(r, "slide_count") != before + 1)
            {
                fails.Add("应 +1 且 action=add");
            }

            string nid = GetString(r, "slide_id");
            if (string.IsNullOrEmpty(nid) || string.Equals(nid, id1, StringComparison.Ordinal)
                || IndexOf(Rows(r), nid) != GetInt(r, "slide_count"))
            {
                fails.Add("新页应在最后且 id 新");
            }

            Finish(run, "sld-add-tail", fails);
        }

        private static void AssertReturnsSlides(TestRun run, ToolResult r)
        {
            var fails = new List<string>();
            if (!Ok(r, fails))
            {
                Finish(run, "sld-add-returns-slides", fails);
                return;
            }

            IList<SlideRow> rows = Rows(r);
            if (rows.Count != GetInt(r, "slide_count") || rows.Count == 0)
            {
                fails.Add("slides 行数应等于 slide_count");
            }

            for (int i = 0; i < rows.Count; i++)
            {
                if (rows[i].Index != i + 1 || string.IsNullOrEmpty(rows[i].SlideId))
                {
                    fails.Add("每行须有 index 与 slide_id");
                    break;
                }
            }

            Finish(run, "sld-add-returns-slides", fails);
        }

        private static void AssertNewHiddenFalse(TestRun run, ToolResult r)
        {
            var fails = new List<string>();
            if (!Ok(r, fails))
            {
                Finish(run, "sld-add-hidden-false", fails);
                return;
            }

            string nid = GetString(r, "slide_id");
            IList<SlideRow> rows = Rows(r);
            for (int i = 0; i < rows.Count; i++)
            {
                if (string.Equals(rows[i].SlideId, nid, StringComparison.Ordinal) && rows[i].Hidden)
                {
                    fails.Add("新建页不应 hidden");
                    break;
                }
            }

            Finish(run, "sld-add-hidden-false", fails);
        }

        private static void ExpectFailUnchanged(
            TestRun run,
            string id,
            Dictionary<string, object> args,
            Deck deck,
            int expectCount,
            string contains)
        {
            IList<string> before = deck.Ids();
            ToolResult r = Invoke(args);
            var fails = new List<string>();
            if (r == null || r.Success)
            {
                fails.Add("应失败");
            }
            else if ((r.Error ?? "").IndexOf(contains, StringComparison.OrdinalIgnoreCase) < 0)
            {
                fails.Add("文案应含「" + contains + "」实际 " + (r.Error ?? ""));
            }

            if (deck.Count() != expectCount)
            {
                fails.Add("页数被改了");
            }
            else
            {
                IList<string> after = deck.Ids();
                if (after.Count == before.Count)
                {
                    for (int i = 0; i < before.Count; i++)
                    {
                        if (!string.Equals(before[i], after[i], StringComparison.Ordinal))
                        {
                            fails.Add("页序被改了");
                            break;
                        }
                    }
                }
            }

            Finish(run, id, fails);
        }

        private static void FailAll(TestRun run, IList<string> filters, string err, string[] ids)
        {
            FailNeeded(run, filters, err, ids);
        }

        private static void FailNeeded(TestRun run, IList<string> filters, string err, params string[] ids)
        {
            for (int i = 0; i < ids.Length; i++)
            {
                if (Want(filters, ids[i]))
                {
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

        private static Dictionary<string, object> Dup(Deck deck, string slideId)
        {
            var a = deck.Args();
            a["action"] = "duplicate";
            a["slide_id"] = slideId;
            return a;
        }

        private static Dictionary<string, object> Del(Deck deck, string slideId)
        {
            var a = deck.Args();
            a["action"] = "delete";
            a["slide_id"] = slideId;
            a["confirm"] = true;
            return a;
        }

        private static Dictionary<string, object> Move(Deck deck, string slideId, int to)
        {
            var a = deck.Args();
            a["action"] = "move";
            a["slide_id"] = slideId;
            a["to_index"] = to;
            return a;
        }

        private static ToolResult Invoke(Dictionary<string, object> args)
        {
            var tools = new Dictionary<string, Func<Dictionary<string, object>, Task<ToolResult>>>(
                StringComparer.OrdinalIgnoreCase);
            F_ManagePptSlideTool.Register(tools, null);
            if (!tools.TryGetValue("F_manage_ppt_slide", out Func<Dictionary<string, object>, Task<ToolResult>> fn))
            {
                return null;
            }

            return fn(args ?? new Dictionary<string, object>()).GetAwaiter().GetResult();
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

        private static bool Align(ToolResult r, Deck deck, List<string> fails)
        {
            IList<string> com = deck.Ids();
            IList<SlideRow> rows = Rows(r);
            if (GetInt(r, "slide_count") != com.Count || rows.Count != com.Count)
            {
                fails.Add("slide_count/slides 与 COM 不一致");
                return false;
            }

            for (int i = 0; i < com.Count; i++)
            {
                if (rows[i].Index != i + 1 || !string.Equals(rows[i].SlideId, com[i], StringComparison.Ordinal))
                {
                    fails.Add("第 " + (i + 1) + " 页 id 与 COM 不符");
                    return false;
                }
            }

            return true;
        }

        private static bool OrderIs(IList<SlideRow> rows, params string[] ids)
        {
            if (rows == null || rows.Count != ids.Length)
            {
                return false;
            }

            for (int i = 0; i < ids.Length; i++)
            {
                if (!string.Equals(rows[i].SlideId, ids[i], StringComparison.Ordinal))
                {
                    return false;
                }
            }

            return true;
        }

        private static int IndexOf(IList<SlideRow> rows, string id)
        {
            for (int i = 0; i < rows.Count; i++)
            {
                if (string.Equals(rows[i].SlideId, id, StringComparison.Ordinal))
                {
                    return rows[i].Index;
                }
            }

            return -1;
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

        private static int GetInt(ToolResult r, string key)
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

        private static IList<SlideRow> Rows(ToolResult r)
        {
            var list = new List<SlideRow>();
            var d = r != null ? r.Data as Dictionary<string, object> : null;
            if (d == null || !d.TryGetValue("slides", out object raw) || !(raw is System.Collections.IEnumerable en))
            {
                return list;
            }

            foreach (object item in en)
            {
                var one = item as Dictionary<string, object>;
                if (one == null)
                {
                    continue;
                }

                int index = 0;
                if (one.TryGetValue("index", out object iv) && iv != null)
                {
                    try
                    {
                        index = Convert.ToInt32(iv, CultureInfo.InvariantCulture);
                    }
                    catch
                    {
                    }
                }

                bool hidden = false;
                if (one.TryGetValue("hidden", out object hv) && hv != null)
                {
                    try
                    {
                        hidden = Convert.ToBoolean(hv);
                    }
                    catch
                    {
                    }
                }

                list.Add(new SlideRow
                {
                    Index = index,
                    SlideId = one.TryGetValue("slide_id", out object sv) ? Convert.ToString(sv) ?? "" : "",
                    Layout = one.TryGetValue("layout", out object lv) ? Convert.ToString(lv) ?? "" : "",
                    Hidden = hidden
                });
            }

            return list;
        }

        private static IList<CreatedRow> Created(ToolResult r)
        {
            var list = new List<CreatedRow>();
            var d = r != null ? r.Data as Dictionary<string, object> : null;
            if (d == null || !d.TryGetValue("created", out object raw) || !(raw is System.Collections.IEnumerable en))
            {
                return list;
            }

            foreach (object item in en)
            {
                var one = item as Dictionary<string, object>;
                if (one == null)
                {
                    continue;
                }

                int index = 0;
                if (one.TryGetValue("index", out object ix) && ix != null)
                {
                    try
                    {
                        index = Convert.ToInt32(ix, CultureInfo.InvariantCulture);
                    }
                    catch
                    {
                    }
                }

                list.Add(new CreatedRow
                {
                    Source = one.TryGetValue("source_slide_id", out object s) ? Convert.ToString(s) ?? "" : "",
                    NewId = one.TryGetValue("slide_id", out object n) ? Convert.ToString(n) ?? "" : "",
                    Index = index
                });
            }

            return list;
        }

        private static bool DeletedHas(ToolResult r, string id)
        {
            IList<string> list = DeletedList(r);
            for (int i = 0; i < list.Count; i++)
            {
                if (string.Equals(list[i], id, StringComparison.Ordinal))
                {
                    return true;
                }
            }

            return false;
        }

        private static IList<string> DeletedList(ToolResult r)
        {
            var list = new List<string>();
            var d = r != null ? r.Data as Dictionary<string, object> : null;
            if (d == null || !d.TryGetValue("deleted", out object raw) || !(raw is System.Collections.IEnumerable en))
            {
                return list;
            }

            foreach (object item in en)
            {
                list.Add(Convert.ToString(item) ?? "");
            }

            return list;
        }

        private sealed class SlideRow
        {
            public int Index;
            public string SlideId;
            public string Layout;
            public bool Hidden;
        }

        private sealed class CreatedRow
        {
            public string Source;
            public string NewId;
            public int Index;
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
            }

            public IOperationChannel Channel { get; private set; }

            public object Pres { get; set; }

            public bool Wpp { get; private set; }

            public Dictionary<string, object> Args()
            {
                return new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase)
                {
                    ["channel_id"] = Channel.ChannelId
                };
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

            public IList<string> Ids()
            {
                Refresh();
                var ids = new List<string>();
                if (Wpp)
                {
                    object slides = WppCom.GetProperty(Pres, "Slides");
                    int n = Convert.ToInt32(WppCom.GetProperty(slides, "Count") ?? 0);
                    for (int i = 1; i <= n; i++)
                    {
                        object slide = WppCom.GetIndexed(slides, i);
                        ids.Add(Convert.ToString(
                            WppCom.GetProperty(slide, "SlideID") ?? WppCom.GetProperty(slide, "SlideId")) ?? "");
                    }

                    return ids;
                }

                PowerPoint.Presentation ppt = (PowerPoint.Presentation)Pres;
                for (int i = 1; i <= ppt.Slides.Count; i++)
                {
                    ids.Add(ppt.Slides[i].SlideID.ToString(CultureInfo.InvariantCulture));
                }

                return ids;
            }

            public bool EnsureCount(int n, out string error)
            {
                error = null;
                Refresh();
                if (Count() == 0)
                {
                    var add0 = Args();
                    add0["action"] = "add";
                    ToolResult first = Invoke(add0);
                    if (first == null || !first.Success)
                    {
                        error = first == null ? "无法 add 首页" : first.Error;
                        return false;
                    }

                    Refresh();
                }

                while (Count() < n)
                {
                    var add = Args();
                    add["action"] = "add";
                    ToolResult r = Invoke(add);
                    if (r == null || !r.Success)
                    {
                        error = r == null ? "add 失败" : r.Error;
                        return false;
                    }

                    Refresh();
                }

                while (Count() > n)
                {
                    IList<string> ids = Ids();
                    if (ids.Count <= n)
                    {
                        break;
                    }

                    var del = Del(this, ids[ids.Count - 1]);
                    ToolResult r = Invoke(del);
                    if (r == null || !r.Success)
                    {
                        error = r == null ? "删到目标页数失败" : r.Error;
                        return false;
                    }

                    Refresh();
                }

                return Count() == n;
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

            public bool TryApplyText(string slideId, string text, out string error)
            {
                Refresh();
                string html = ContentHtml.TextboxCreate(text, 8, 10, 30, 12);
                if (!ContentHtml.TryParse(html, slideId, out PptHtmlApplyPlan plan, out error))
                {
                    return false;
                }

                plan.AllowCreate = true;
                return Wpp
                    ? PptHtmlWppApplier.TryApply(Pres, plan, "sld-wpp", out _, out error)
                    : PptHtmlPowerPointApplier.TryApply(
                        (PowerPoint.Presentation)Pres, plan, "sld-ppt", out _, out error);
            }

            public bool PageHasText(string slideId, string text, out string error)
            {
                error = null;
                Refresh();
                PptHtmlReadResult tree;
                bool ok = Wpp
                    ? PptHtmlWppReader.TryRead(Pres, slideId, "sld-wpp", "wpp", out tree, out error)
                    : PptHtmlPowerPointReader.TryRead(
                        (PowerPoint.Presentation)Pres, slideId, "sld-ppt", "ppt", out tree, out error);
                if (!ok || tree == null)
                {
                    return false;
                }

                return ForestHasText(tree.Shapes, text);
            }

            private static bool ForestHasText(IList<PptHtmlShapeNode> nodes, string text)
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

                    if (!string.IsNullOrEmpty(n.Text) && n.Text.IndexOf(text, StringComparison.Ordinal) >= 0)
                    {
                        return true;
                    }

                    if (ForestHasText(n.Children, text))
                    {
                        return true;
                    }
                }

                return false;
            }
        }
    }
}
