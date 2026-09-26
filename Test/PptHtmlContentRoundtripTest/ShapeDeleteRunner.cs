using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using WordAddIn1;
using WordAddIn1.OpenFiles;
using WordAddIn1.PresentationHost;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PptHtmlContentRoundtripTest
{
    /// <summary>F_manage_ppt_shape delete。前缀 shp-。</summary>
    internal static class ShapeDeleteRunner
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
                if (f.IndexOf("shp-", StringComparison.OrdinalIgnoreCase) >= 0
                    || string.Equals(f, "shp", StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
            }

            return false;
        }

        public static void Run(TestRun run, IList<string> nameFilters, bool useWpp)
        {
            Console.WriteLine(useWpp ? "删形状 · WPP" : "删形状 · PPT");
            RunParse(run, nameFilters);
            if (NeedCom(nameFilters))
            {
                RunCom(run, nameFilters, useWpp);
            }

            Console.WriteLine("删形状 通过 " + run.CasesPassed
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
                "shp-neg-no-slide", "shp-neg-no-shape", "shp-neg-empty-ids", "shp-neg-bad-id",
                "shp-neg-missing", "shp-neg-cross", "shp-neg-partial", "shp-neg-empty-item",
                "shp-neg-after", "shp-neg-child", "shp-del-", "shp-echo-"
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
                if (string.Equals(f, "shp", StringComparison.OrdinalIgnoreCase)
                    || string.Equals(f, "shp-", StringComparison.OrdinalIgnoreCase))
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
            if (fails == null || fails.Count == 0)
            {
                run.CaseOk(id);
            }
            else
            {
                run.CaseFail(id, string.Join("；", fails));
            }
        }

        private static void Finish(TestRun run, string id, string fail)
        {
            Finish(run, id, string.IsNullOrEmpty(fail) ? new List<string>() : new List<string> { fail });
        }

        private static void RunParse(TestRun run, IList<string> filters)
        {
            Console.WriteLine("--- A 渠道 / 互斥 ---");
            ChannelRegistry.Register(new KindStub("shp-stub-ppt", ChannelKind.Ppt), false);
            ChannelRegistry.Register(new KindStub("shp-stub-word", ChannelKind.Word), false);
            ChannelRegistry.Register(new KindStub("shp-stub-excel", ChannelKind.Excel), false);

            if (Want(filters, "shp-neg-no-channel"))
            {
                ExpectFail(run, "shp-neg-no-channel", new Dictionary<string, object>(), null);
            }

            if (Want(filters, "shp-neg-no-action"))
            {
                ExpectFail(run, "shp-neg-no-action", Args("shp-stub-ppt"), "delete");
            }

            if (Want(filters, "shp-neg-bad-action"))
            {
                var a = Args("shp-stub-ppt");
                a["action"] = "move";
                ExpectFail(run, "shp-neg-bad-action", a, "delete");
            }

            if (Want(filters, "shp-neg-wrong-kind"))
            {
                var a = Args("shp-stub-word");
                a["action"] = "delete";
                a["slide_id"] = "1";
                a["shape_id"] = "s1";
                ExpectFail(run, "shp-neg-wrong-kind", a, "渠道");
            }

            if (Want(filters, "shp-neg-excel"))
            {
                var a = Args("shp-stub-excel");
                a["action"] = "delete";
                a["slide_id"] = "1";
                a["shape_id"] = "s1";
                ExpectFail(run, "shp-neg-excel", a, "渠道");
            }

            Foreign(run, filters, "shp-neg-del-source", "source_channel_id", "x");
            Foreign(run, filters, "shp-neg-del-to-slide", "to_slide_id", "1");
            Foreign(run, filters, "shp-neg-del-left", "left", 8);
            Foreign(run, filters, "shp-neg-del-top", "top", 8);
            Foreign(run, filters, "shp-neg-del-width", "width", 10);
            Foreign(run, filters, "shp-neg-del-height", "height", 10);

            if (Want(filters, "shp-neg-del-source-empty"))
            {
                var a = DelArgs("shp-stub-ppt", "1", "s1");
                a["source_channel_id"] = "";
                ExpectFail(run, "shp-neg-del-source-empty", a, "source_channel_id");
            }
        }

        private static void Foreign(TestRun run, IList<string> filters, string id, string key, object val)
        {
            if (!Want(filters, id))
            {
                return;
            }

            var a = DelArgs("shp-stub-ppt", "1", "s1");
            a[key] = val;
            ExpectFail(run, id, a, key);
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
            Console.WriteLine(useWpp ? "--- B–G COM（WPP）---" : "--- B–G COM ---");
            string assetsDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "shp-assets");
            ContentAssets assets = ContentAssetsIo.Ensure(assetsDir);
            if (useWpp)
            {
                RunComWpp(run, filters, assets);
            }
            else
            {
                RunComPpt(run, filters, assets);
            }
        }

        private static void RunComPpt(TestRun run, IList<string> filters, ContentAssets assets)
        {
            PowerPoint.Application app = null;
            PowerPoint.Presentation pres = null;
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

                pres = app.Presentations.Add(Office.MsoTriState.msoTrue);
                PptChannel ch = ChannelRegistry.CreateOrGetPpt(pres, "", false);
                if (!ch.TryGetLivePresentation(out PowerPoint.Presentation live) || live == null)
                {
                    run.CaseFail("shp-", "渠道拿不到活稿");
                    return;
                }

                RunBlocks(run, filters, new Deck(ch, live, assets, wpp: false));
            }
            catch (Exception ex)
            {
                run.CaseFail("shp-", "COM 失败: " + ex.Message);
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

        private static void RunComWpp(TestRun run, IList<string> filters, ContentAssets assets)
        {
            object app = null;
            object pres = null;
            try
            {
                if (!WppCom.TryCreateApplication(out app, out string startErr))
                {
                    run.CaseFail("shp-", startErr);
                    return;
                }

                object presentations = WppCom.GetProperty(app, "Presentations");
                try
                {
                    pres = WppCom.Invoke(presentations, "Add", true);
                }
                catch
                {
                    pres = WppCom.Invoke(presentations, "Add");
                }

                if (pres == null)
                {
                    run.CaseFail("shp-", "WPP Add 失败");
                    return;
                }

                WppChannel ch = ChannelRegistry.CreateOrGetWpp(pres, "", false);
                if (!ch.TryGetLivePresentation(out object live) || live == null)
                {
                    run.CaseFail("shp-", "渠道拿不到活稿");
                    return;
                }

                RunBlocks(run, filters, new Deck(ch, live, assets, wpp: true));
            }
            catch (Exception ex)
            {
                run.CaseFail("shp-", "COM 失败: " + ex.Message);
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
            RunPre(run, filters, deck);
            RunSingle(run, filters, deck);
            RunMulti(run, filters, deck);
            RunGroup(run, filters, deck);
            RunHiddenPages(run, filters, deck);
            RunEcho(run, filters, deck);
        }

        private static void RunPre(TestRun run, IList<string> filters, Deck deck)
        {
            if (!Want(filters, "shp-neg-no-slide") && !Want(filters, "shp-neg-no-shape")
                && !Want(filters, "shp-neg-empty-ids") && !Want(filters, "shp-neg-bad-id")
                && !Want(filters, "shp-neg-missing") && !Want(filters, "shp-neg-cross")
                && !Want(filters, "shp-neg-partial") && !Want(filters, "shp-neg-empty-item")
                && !Want(filters, "shp-neg-after-ok"))
            {
                return;
            }

            if (!deck.TryTwoBoxes(out string a, out string b, out string err))
            {
                FailNeeded(
                    run,
                    filters,
                    err,
                    "shp-neg-no-slide",
                    "shp-neg-no-shape",
                    "shp-neg-empty-ids",
                    "shp-neg-bad-id",
                    "shp-neg-missing-slide",
                    "shp-neg-missing-shape",
                    "shp-neg-cross-page",
                    "shp-neg-partial-fail",
                    "shp-neg-empty-item",
                    "shp-neg-after-ok");
                return;
            }

            int before = deck.CountIds();
            if (Want(filters, "shp-neg-no-slide"))
            {
                var x = deck.Args();
                x["action"] = "delete";
                x["shape_ids"] = new[] { a };
                ExpectFailUnchanged(run, "shp-neg-no-slide", x, deck, before, a, b, "slide_id");
            }

            if (Want(filters, "shp-neg-no-shape"))
            {
                var x = deck.Args();
                x["action"] = "delete";
                x["slide_id"] = deck.SlideId;
                ExpectFailUnchanged(run, "shp-neg-no-shape", x, deck, before, a, b, "shape_ids");
            }

            if (Want(filters, "shp-neg-empty-ids"))
            {
                var x = deck.Args();
                x["action"] = "delete";
                x["slide_id"] = deck.SlideId;
                x["shape_ids"] = new string[0];
                ExpectFailUnchanged(run, "shp-neg-empty-ids", x, deck, before, a, b, "shape_ids");
            }

            if (Want(filters, "shp-neg-bad-id"))
            {
                ExpectFailUnchanged(
                    run,
                    "shp-neg-bad-id",
                    deck.DelIds("nope"),
                    deck,
                    before,
                    a,
                    b,
                    "非法");
            }

            if (Want(filters, "shp-neg-missing-slide"))
            {
                var x = deck.DelIds(a);
                x["slide_id"] = "999999";
                ExpectFailUnchanged(run, "shp-neg-missing-slide", x, deck, before, a, b, "不存在");
            }

            if (Want(filters, "shp-neg-missing-shape"))
            {
                string ghost = "sid" + deck.SlideId + "-s99999";
                ExpectFailUnchanged(
                    run,
                    "shp-neg-missing-shape",
                    deck.DelIds(ghost),
                    deck,
                    before,
                    a,
                    b,
                    "找不到");
            }

            if (Want(filters, "shp-neg-partial-fail"))
            {
                ExpectFailUnchanged(
                    run,
                    "shp-neg-partial-fail",
                    deck.DelIds(a, "sid" + deck.SlideId + "-s99999"),
                    deck,
                    before,
                    a,
                    b,
                    "找不到");
            }

            if (Want(filters, "shp-neg-empty-item"))
            {
                ExpectFailUnchanged(
                    run,
                    "shp-neg-empty-item",
                    deck.DelIds(a, ""),
                    deck,
                    before,
                    a,
                    b,
                    "非法");
            }

            if (Want(filters, "shp-neg-cross-page"))
            {
                if (!deck.TrySecondForeignBox(out string p2, out err))
                {
                    run.CaseFail("shp-neg-cross-page", err);
                }
                else
                {
                    var x = deck.Args();
                    x["action"] = "delete";
                    x["slide_id"] = deck.SlideId;
                    x["shape_ids"] = new[] { p2 };
                    ToolResult r = Invoke(x);
                    var fails = new List<string>();
                    if (r != null && r.Success)
                    {
                        fails.Add("跨页应失败");
                    }
                    else if ((r == null ? "" : r.Error ?? "").IndexOf("找不到", StringComparison.Ordinal) < 0)
                    {
                        fails.Add("文案应含找不到");
                    }

                    if (!deck.HasId(a) || !deck.HasId(b) || !deck.HasIdOn(deck.SlideId2, p2))
                    {
                        fails.Add("两页形状应都在");
                    }

                    Finish(run, "shp-neg-cross-page", fails);
                }
            }

            if (Want(filters, "shp-neg-after-ok"))
            {
                if (!deck.TryTwoBoxes(out a, out b, out err))
                {
                    run.CaseFail("shp-neg-after-ok", err);
                }
                else
                {
                    Invoke(deck.DelIds("nope"));
                    ToolResult r = Invoke(deck.DelIds(a));
                    var fails = new List<string>();
                    if (!OkDel(r, fails, 1) || deck.HasId(a) || !deck.HasId(b))
                    {
                        if (fails.Count == 0)
                        {
                            fails.Add("失败后再删 a 应只剩 b");
                        }
                    }

                    Finish(run, "shp-neg-after-ok", fails);
                }
            }
        }

        private static void RunSingle(TestRun run, IList<string> filters, Deck deck)
        {
            if (!Want(filters, "shp-del-"))
            {
                return;
            }

            OneTwo(run, filters, deck, "shp-del-one-tb", (d, a, b) =>
            {
                ToolResult r = Invoke(d.DelIds(a));
                return CheckOneGone(r, d, a, b);
            });

            OneTwo(run, filters, deck, "shp-del-shape-id", (d, a, b) =>
            {
                ToolResult r = Invoke(d.DelShape(a));
                return CheckOneGone(r, d, a, b);
            });

            OneTwo(run, filters, deck, "shp-del-short-id", (d, a, b) =>
            {
                if (!PptShapeId.TryParseShape(a, out _, out int n))
                {
                    return "短写无法从 a 解析";
                }

                ToolResult r = Invoke(d.DelIds("s" + n.ToString(CultureInfo.InvariantCulture)));
                var fails = CheckOneGone(r, d, a, b);
                if (fails == null && !DeletedHas(r, a))
                {
                    return "回包应为完整 id " + a;
                }

                return fails;
            });

            OneTwo(run, filters, deck, "shp-del-ids-wins", (d, a, b) =>
            {
                var x = d.DelIds(a);
                x["shape_id"] = b;
                ToolResult r = Invoke(x);
                return CheckOneGone(r, d, a, b);
            });

            OneTwo(run, filters, deck, "shp-del-empty-ids-falls-to-id", (d, a, b) =>
            {
                var x = d.Args();
                x["action"] = "delete";
                x["slide_id"] = d.SlideId;
                x["shape_ids"] = new string[0];
                x["shape_id"] = a;
                return CheckOneGone(Invoke(x), d, a, b);
            });

            if (Want(filters, "shp-del-pic"))
            {
                if (!deck.TryBoxAnd(ContentHtml.PictureCreate(deck.Assets.RedPng, 50, 20, 30, 28), out string box, out string pic, out string err))
                {
                    run.CaseFail("shp-del-pic", err);
                }
                else
                {
                    Finish(run, "shp-del-pic", CheckOneGone(Invoke(deck.DelIds(pic)), deck, pic, box));
                }
            }

            if (Want(filters, "shp-del-table"))
            {
                if (!deck.TryBoxAnd(
                        ContentHtml.TableCreate(new[] { new[] { "1", "2" }, new[] { "3", "4" } }, 8, 40, 40, 24),
                        out string box,
                        out string tbl,
                        out string err))
                {
                    run.CaseFail("shp-del-table", err);
                }
                else
                {
                    Finish(run, "shp-del-table", CheckOneGone(Invoke(deck.DelIds(tbl)), deck, tbl, box));
                }
            }

            if (Want(filters, "shp-del-chart"))
            {
                if (!deck.TryBoxAnd(ContentHtml.PieChart(), out string box, out string pie, out string err))
                {
                    run.CaseFail("shp-del-chart", err ?? "铺饼失败");
                }
                else
                {
                    Finish(run, "shp-del-chart", CheckOneGone(Invoke(deck.DelIds(pie)), deck, pie, box));
                }
            }

            OneTwo(run, filters, deck, "shp-del-sibling-geo", (d, a, b) =>
            {
                if (!d.TryGeo(b, out double left, out double top, out string gerr))
                {
                    return gerr;
                }

                string gone = CheckOneGone(Invoke(d.DelIds(a)), d, a, b);
                if (gone != null)
                {
                    return gone;
                }

                if (!d.TryGeo(b, out double left2, out double top2, out gerr))
                {
                    return gerr;
                }

                if (Math.Abs(left - left2) > 1 || Math.Abs(top - top2) > 1)
                {
                    return "兄弟几何应差不多";
                }

                return null;
            });

            if (Want(filters, "shp-del-last-ok"))
            {
                if (!deck.TryWipeApply(ContentHtml.TextboxCreate("仅一"), out List<string> ids, out string err)
                    || ids.Count < 1)
                {
                    run.CaseFail("shp-del-last-ok", err ?? "铺一框失败");
                }
                else
                {
                    ToolResult r = Invoke(deck.DelIds(ids[0]));
                    var fails = new List<string>();
                    if (!OkDel(r, fails, 1) || deck.HasId(ids[0]) || deck.CountIds() != 0)
                    {
                        if (fails.Count == 0)
                        {
                            fails.Add("删最后一形后应无内容 id");
                        }
                    }

                    Finish(run, "shp-del-last-ok", fails);
                }
            }

            OneTwo(run, filters, deck, "shp-del-id-spaces", (d, a, b) =>
                CheckOneGone(Invoke(d.DelShape("  " + a + "  ")), d, a, b));

            OneTwo(run, filters, deck, "shp-del-action-upper", (d, a, b) =>
            {
                var x = d.DelIds(a);
                x["action"] = "DELETE";
                ToolResult r = Invoke(x);
                string gone = CheckOneGone(r, d, a, b);
                if (gone != null)
                {
                    return gone;
                }

                return string.Equals(GetString(r, "action"), "delete", StringComparison.Ordinal)
                    ? null
                    : "action 应归一 delete";
            });
        }

        private static void RunMulti(TestRun run, IList<string> filters, Deck deck)
        {
            if (!Want(filters, "shp-del-two") && !Want(filters, "shp-del-three")
                && !Want(filters, "shp-del-dup-ok") && !Want(filters, "shp-del-list")
                && !Want(filters, "shp-del-order") && !Want(filters, "shp-del-two-types")
                && !Want(filters, "shp-del-read-gone") && !Want(filters, "shp-del-com-gone"))
            {
                return;
            }

            if (Want(filters, "shp-del-two") || Want(filters, "shp-del-three")
                || Want(filters, "shp-del-dup-ok") || Want(filters, "shp-del-list")
                || Want(filters, "shp-del-order") || Want(filters, "shp-del-read-gone")
                || Want(filters, "shp-del-com-gone"))
            {
                if (!deck.TryThreeBoxes(out string a, out string b, out string c, out string err))
                {
                    FailNeeded(
                        run,
                        filters,
                        err,
                        "shp-del-two",
                        "shp-del-three",
                        "shp-del-dup-ok",
                        "shp-del-list",
                        "shp-del-order",
                        "shp-del-read-gone",
                        "shp-del-com-gone");
                }
                else
                {
                    if (Want(filters, "shp-del-two"))
                    {
                        ToolResult r = Invoke(deck.DelIds(a, c));
                        var fails = new List<string>();
                        if (!OkDel(r, fails, 2) || deck.HasId(a) || deck.HasId(c) || !deck.HasId(b))
                        {
                            if (fails.Count == 0)
                            {
                                fails.Add("应剩 b");
                            }
                        }

                        Finish(run, "shp-del-two", fails);
                    }

                    if (Want(filters, "shp-del-dup-ok")
                        && !deck.TryThreeBoxes(out a, out b, out c, out err))
                    {
                        run.CaseFail("shp-del-dup-ok", err);
                    }
                    else if (Want(filters, "shp-del-dup-ok"))
                    {
                        ToolResult r = Invoke(deck.DelIds(a, a));
                        var fails = new List<string>();
                        if (!OkDel(r, fails, 1) || deck.HasId(a) || !deck.HasId(b) || !deck.HasId(c))
                        {
                            if (fails.Count == 0)
                            {
                                fails.Add("去重后应只删 a");
                            }
                        }

                        Finish(run, "shp-del-dup-ok", fails);
                    }

                    if (Want(filters, "shp-del-list")
                        && !deck.TryThreeBoxes(out a, out b, out c, out err))
                    {
                        run.CaseFail("shp-del-list", err);
                    }
                    else if (Want(filters, "shp-del-list"))
                    {
                        var x = deck.Args();
                        x["action"] = "delete";
                        x["slide_id"] = deck.SlideId;
                        x["shape_ids"] = new List<string> { a };
                        Finish(run, "shp-del-list", CheckOneGone(Invoke(x), deck, a, b));
                    }

                    if (Want(filters, "shp-del-order")
                        && !deck.TryThreeBoxes(out a, out b, out c, out err))
                    {
                        run.CaseFail("shp-del-order", err);
                    }
                    else if (Want(filters, "shp-del-order"))
                    {
                        ToolResult r = Invoke(deck.DelIds(c, a));
                        IList<string> del = DeletedList(r);
                        var fails = new List<string>();
                        if (!OkDel(r, fails, 2) || del.Count != 2 || del[0] != c || del[1] != a)
                        {
                            if (fails.Count == 0)
                            {
                                fails.Add("deleted 序应为 c,a");
                            }
                        }

                        Finish(run, "shp-del-order", fails);
                    }

                    if (Want(filters, "shp-del-three")
                        && !deck.TryThreeBoxes(out a, out b, out c, out err))
                    {
                        run.CaseFail("shp-del-three", err);
                    }
                    else if (Want(filters, "shp-del-three"))
                    {
                        ToolResult r = Invoke(deck.DelIds(a, b, c));
                        var fails = new List<string>();
                        if (!OkDel(r, fails, 3) || deck.CountIds() != 0)
                        {
                            if (fails.Count == 0)
                            {
                                fails.Add("三框应删光");
                            }
                        }

                        Finish(run, "shp-del-three", fails);
                    }

                    if (Want(filters, "shp-del-read-gone")
                        && !deck.TryThreeBoxes(out a, out b, out c, out err))
                    {
                        run.CaseFail("shp-del-read-gone", err);
                    }
                    else if (Want(filters, "shp-del-read-gone"))
                    {
                        Invoke(deck.DelIds(a));
                        var fails = new List<string>();
                        if (deck.HasId(a) || !deck.HasId(b))
                        {
                            fails.Add("读回仍见 a 或丢了 b");
                        }

                        Finish(run, "shp-del-read-gone", fails);
                    }

                    if (Want(filters, "shp-del-com-gone")
                        && !deck.TryThreeBoxes(out a, out b, out c, out err))
                    {
                        run.CaseFail("shp-del-com-gone", err);
                    }
                    else if (Want(filters, "shp-del-com-gone"))
                    {
                        Invoke(deck.DelIds(a));
                        var fails = new List<string>();
                        if (deck.ComHas(a) || !deck.ComHas(b))
                        {
                            fails.Add("COM 仍能找到 a 或丢了 b");
                        }

                        Finish(run, "shp-del-com-gone", fails);
                    }
                }
            }

            if (Want(filters, "shp-del-two-types"))
            {
                if (!deck.TryBoxAnd(ContentHtml.PictureCreate(deck.Assets.RedPng, 50, 20, 30, 28), out string box, out string pic, out string err))
                {
                    run.CaseFail("shp-del-two-types", err);
                }
                else
                {
                    ToolResult r = Invoke(deck.DelIds(box, pic));
                    var fails = new List<string>();
                    if (!OkDel(r, fails, 2) || deck.HasId(box) || deck.HasId(pic))
                    {
                        if (fails.Count == 0)
                        {
                            fails.Add("框和图都应没了");
                        }
                    }

                    Finish(run, "shp-del-two-types", fails);
                }
            }
        }

        private static void RunGroup(TestRun run, IList<string> filters, Deck deck)
        {
            if (!Want(filters, "shp-del-group") && !Want(filters, "shp-del-group-child")
                && !Want(filters, "shp-del-group-keep") && !Want(filters, "shp-del-nested")
                && !Want(filters, "shp-del-group-then") && !Want(filters, "shp-neg-child"))
            {
                return;
            }

            if (Want(filters, "shp-del-group") || Want(filters, "shp-del-group-then-tb"))
            {
                if (!deck.TrySeedGroup(out string g, out string a, out string b, out string err))
                {
                    FailNeeded(run, filters, err, "shp-del-group", "shp-del-group-then-tb");
                }
                else
                {
                    if (Want(filters, "shp-del-group"))
                    {
                        ToolResult r = Invoke(deck.DelIds(g));
                        var fails = new List<string>();
                        if (!OkDel(r, fails, 1) || deck.HasId(g) || deck.HasId(a) || deck.HasId(b))
                        {
                            if (fails.Count == 0)
                            {
                                fails.Add("删组壳后叶子也应没了");
                            }
                        }

                        Finish(run, "shp-del-group", fails);
                    }

                    if (Want(filters, "shp-del-group-then-tb"))
                    {
                        if (!deck.TrySeedGroup(out g, out _, out _, out err))
                        {
                            run.CaseFail("shp-del-group-then-tb", err);
                        }
                        else
                        {
                            Invoke(deck.DelIds(g));
                            var fails = new List<string>();
                            if (!deck.TryWipeApply(ContentHtml.TextboxCreate("组后新框"), out List<string> neu, out err)
                                || neu.Count < 1)
                            {
                                fails.Add(err ?? "删组后再 apply 失败");
                            }

                            Finish(run, "shp-del-group-then-tb", fails);
                        }
                    }
                }
            }

            if (Want(filters, "shp-del-group-child") || Want(filters, "shp-del-group-keep-other")
                || Want(filters, "shp-neg-child-plus-missing"))
            {
                if (!deck.TrySeedGroup(out string g, out string a, out string b, out string err))
                {
                    FailNeeded(
                        run,
                        filters,
                        err,
                        "shp-del-group-child",
                        "shp-del-group-keep-other",
                        "shp-neg-child-plus-missing");
                }
                else
                {
                    if (Want(filters, "shp-neg-child-plus-missing"))
                    {
                        int before = deck.CountIds();
                        ExpectFailUnchanged(
                            run,
                            "shp-neg-child-plus-missing",
                            deck.DelIds(a, "sid" + deck.SlideId + "-s99999"),
                            deck,
                            before,
                            a,
                            b,
                            "找不到");
                        if (!deck.HasId(g) && Want(filters, "shp-neg-child-plus-missing"))
                        {
                            // 组壳读回可能是空壳；叶子还在即可。上面已查 a,b
                        }
                    }

                    if (Want(filters, "shp-del-group-child")
                        && !deck.TrySeedGroup(out g, out a, out b, out err))
                    {
                        run.CaseFail("shp-del-group-child", err);
                    }
                    else if (Want(filters, "shp-del-group-child"))
                    {
                        ToolResult r = Invoke(deck.DelIds(a));
                        var fails = new List<string>();
                        if (!OkDel(r, fails, 1) || deck.HasId(a) || !deck.PageHasText("源正文"))
                        {
                            if (fails.Count == 0)
                            {
                                fails.Add("应删叶子 a，b 文案还在");
                            }
                        }

                        Finish(run, "shp-del-group-child", fails);
                    }

                    if (Want(filters, "shp-del-group-keep-other")
                        && !deck.TrySeedGroup(out g, out a, out b, out err))
                    {
                        run.CaseFail("shp-del-group-keep-other", err);
                    }
                    else if (Want(filters, "shp-del-group-keep-other"))
                    {
                        Invoke(deck.DelIds(a));
                        var fails = new List<string>();
                        if (!deck.PageHasText("源正文"))
                        {
                            fails.Add("未点名叶子正文应仍在");
                        }

                        Finish(run, "shp-del-group-keep-other", fails);
                    }
                }
            }

            if (Want(filters, "shp-del-nested-child"))
            {
                if (!deck.TryNestedLeaf(out string leaf, out string err))
                {
                    run.CaseFail("shp-del-nested-child", err);
                }
                else
                {
                    ToolResult r = Invoke(deck.DelIds(leaf));
                    var fails = new List<string>();
                    if (!OkDel(r, fails, 1) || deck.HasId(leaf))
                    {
                        if (fails.Count == 0)
                        {
                            fails.Add("内组叶子应被删");
                        }
                    }

                    Finish(run, "shp-del-nested-child", fails);
                }
            }
        }

        private static void RunHiddenPages(TestRun run, IList<string> filters, Deck deck)
        {
            if (!Want(filters, "shp-del-hidden") && !Want(filters, "shp-del-other")
                && !Want(filters, "shp-del-page-2"))
            {
                return;
            }

            if (Want(filters, "shp-del-hidden-ok"))
            {
                if (!deck.TryTwoBoxes(out string a, out string b, out string err) || !deck.TryHide(a, false, out err))
                {
                    run.CaseFail("shp-del-hidden-ok", err ?? "隐藏失败");
                }
                else
                {
                    Finish(run, "shp-del-hidden-ok", CheckOneGone(Invoke(deck.DelIds(a)), deck, a, b));
                }
            }

            if (Want(filters, "shp-del-hidden-sibling"))
            {
                if (!deck.TryTwoBoxes(out string a, out string b, out string err) || !deck.TryHide(a, false, out err))
                {
                    run.CaseFail("shp-del-hidden-sibling", err ?? "隐藏失败");
                }
                else
                {
                    ToolResult r = Invoke(deck.DelIds(b));
                    var fails = new List<string>();
                    if (!OkDel(r, fails, 1) || deck.HasId(b) || !deck.ComHas(a))
                    {
                        if (fails.Count == 0)
                        {
                            fails.Add("应删可见 b，隐藏 a 仍在");
                        }
                    }

                    Finish(run, "shp-del-hidden-sibling", fails);
                }
            }

            if (Want(filters, "shp-del-other-page") || Want(filters, "shp-del-page-2"))
            {
                if (!deck.TryTwoPages(out string p1, out string p2, out string err))
                {
                    FailNeeded(run, filters, err, "shp-del-other-page", "shp-del-page-2");
                }
                else
                {
                    if (Want(filters, "shp-del-other-page"))
                    {
                        ToolResult r = Invoke(deck.DelIds(p1));
                        var fails = new List<string>();
                        if (!OkDel(r, fails, 1) || deck.HasId(p1) || !deck.HasIdOn(deck.SlideId2, p2))
                        {
                            if (fails.Count == 0)
                            {
                                fails.Add("删第 1 页不应动第 2 页");
                            }
                        }

                        Finish(run, "shp-del-other-page", fails);
                    }

                    if (Want(filters, "shp-del-page-2"))
                    {
                        if (!deck.TryTwoPages(out p1, out p2, out err))
                        {
                            run.CaseFail("shp-del-page-2", err);
                        }
                        else
                        {
                            var x = deck.Args();
                            x["action"] = "delete";
                            x["slide_id"] = deck.SlideId2;
                            x["shape_ids"] = new[] { p2 };
                            ToolResult r = Invoke(x);
                            var fails = new List<string>();
                            if (!OkDel(r, fails, 1) || deck.HasIdOn(deck.SlideId2, p2) || !deck.HasIdOn(deck.SlideId, p1))
                            {
                                if (fails.Count == 0)
                                {
                                    fails.Add("删第 2 页不应动第 1 页");
                                }
                            }

                            Finish(run, "shp-del-page-2", fails);
                        }
                    }
                }
            }
        }

        private static void RunEcho(TestRun run, IList<string> filters, Deck deck)
        {
            if (!Want(filters, "shp-echo-"))
            {
                return;
            }

            if (!deck.TryTwoBoxes(out string a, out string b, out string err))
            {
                FailNeeded(
                    run,
                    filters,
                    err,
                    "shp-echo-count",
                    "shp-echo-ids",
                    "shp-echo-kind",
                    "shp-echo-display",
                    "shp-echo-slide",
                    "shp-echo-no-group-id");
                return;
            }

            ToolResult r = Invoke(deck.DelIds(a));
            if (Want(filters, "shp-echo-count"))
            {
                var f = new List<string>();
                if (!OkDel(r, f, 1))
                {
                    Finish(run, "shp-echo-count", f);
                }
                else
                {
                    Finish(run, "shp-echo-count", f);
                }
            }

            if (Want(filters, "shp-echo-ids"))
            {
                IList<string> del = DeletedList(r);
                var f = new List<string>();
                if (del.Count != 1 || del[0] != a || del[0].IndexOf("sid", StringComparison.OrdinalIgnoreCase) < 0)
                {
                    f.Add("应恰一条完整 id");
                }

                Finish(run, "shp-echo-ids", f);
            }

            if (Want(filters, "shp-echo-kind"))
            {
                var f = new List<string>();
                string want = deck.Wpp ? "wpp" : "ppt";
                if (!string.Equals(GetString(r, "kind"), want, StringComparison.OrdinalIgnoreCase))
                {
                    f.Add("kind 应为 " + want);
                }

                Finish(run, "shp-echo-kind", f);
            }

            if (Want(filters, "shp-echo-display"))
            {
                string d = GetString(r, "display_contents");
                var f = new List<string>();
                if (d.IndexOf("action=delete", StringComparison.Ordinal) < 0
                    || d.IndexOf("deleted=", StringComparison.Ordinal) < 0)
                {
                    f.Add("display 应含 action=delete 与 deleted=");
                }

                Finish(run, "shp-echo-display", f);
            }

            if (Want(filters, "shp-echo-slide"))
            {
                var f = new List<string>();
                if (!string.Equals(GetString(r, "slide_id"), deck.SlideId, StringComparison.Ordinal))
                {
                    f.Add("slide_id 应回显本页");
                }

                Finish(run, "shp-echo-slide", f);
            }

            if (Want(filters, "shp-echo-no-group-id"))
            {
                var f = new List<string>();
                if (!string.IsNullOrEmpty(GetString(r, "group_shape_id")))
                {
                    f.Add("delete 不应带 group_shape_id");
                }

                Finish(run, "shp-echo-no-group-id", f);
            }
        }

        private static void OneTwo(
            TestRun run,
            IList<string> filters,
            Deck deck,
            string id,
            Func<Deck, string, string, string> body)
        {
            if (!Want(filters, id))
            {
                return;
            }

            if (!deck.TryTwoBoxes(out string a, out string b, out string err))
            {
                run.CaseFail(id, err);
                return;
            }

            string fail = body(deck, a, b);
            Finish(run, id, fail == null ? new List<string>() : new List<string> { fail });
        }

        private static string CheckOneGone(ToolResult r, Deck deck, string gone, string stay)
        {
            var fails = new List<string>();
            if (!OkDel(r, fails, 1))
            {
                return string.Join("；", fails);
            }

            if (deck.HasId(gone) || !deck.HasId(stay))
            {
                return "应无 " + gone + " 且仍有兄弟";
            }

            return null;
        }

        private static void ExpectFailUnchanged(
            TestRun run,
            string id,
            Dictionary<string, object> args,
            Deck deck,
            int expectCount,
            string keepA,
            string keepB,
            string contains)
        {
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

            if (deck.CountIds() != expectCount || !deck.HasId(keepA) || !deck.HasId(keepB))
            {
                fails.Add("形状被改了");
            }

            Finish(run, id, fails);
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

        private static Dictionary<string, object> DelArgs(string channelId, string slideId, string shapeId)
        {
            var a = Args(channelId);
            a["action"] = "delete";
            a["slide_id"] = slideId;
            a["shape_id"] = shapeId;
            return a;
        }

        private static ToolResult Invoke(Dictionary<string, object> args)
        {
            var tools = new Dictionary<string, Func<Dictionary<string, object>, Task<ToolResult>>>(
                StringComparer.OrdinalIgnoreCase);
            F_ManagePptShapeTool.Register(tools, null);
            if (!tools.TryGetValue("F_manage_ppt_shape", out Func<Dictionary<string, object>, Task<ToolResult>> fn))
            {
                return null;
            }

            return fn(args ?? new Dictionary<string, object>()).GetAwaiter().GetResult();
        }

        private static bool OkDel(ToolResult r, List<string> fails, int count)
        {
            if (r == null || !r.Success)
            {
                fails.Add(r == null ? "无回包" : (r.Error ?? "失败"));
                return false;
            }

            if (GetInt(r, "deleted_count") != count)
            {
                fails.Add("deleted_count=" + GetInt(r, "deleted_count") + " 期望 " + count);
                return false;
            }

            return true;
        }

        private static bool DeletedHas(ToolResult r, string id)
        {
            IList<string> list = DeletedList(r);
            for (int i = 0; i < list.Count; i++)
            {
                if (string.Equals(list[i], id, StringComparison.OrdinalIgnoreCase))
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
            if (d == null || !d.TryGetValue("deleted_shape_ids", out object raw)
                || !(raw is System.Collections.IEnumerable en))
            {
                return list;
            }

            foreach (object item in en)
            {
                list.Add(Convert.ToString(item) ?? "");
            }

            return list;
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
            public Deck(IOperationChannel channel, object pres, ContentAssets assets, bool wpp)
            {
                Channel = channel;
                Pres = pres;
                Assets = assets;
                Wpp = wpp;
                Refresh();
                EnsureFirstSlide();
                SlideId = ReadSlideId(1);
            }

            public IOperationChannel Channel { get; }

            public object Pres { get; set; }

            public ContentAssets Assets { get; }

            public bool Wpp { get; }

            public string SlideId { get; private set; }

            public string SlideId2 { get; private set; }

            public Dictionary<string, object> Args()
            {
                return new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase)
                {
                    ["channel_id"] = Channel.ChannelId
                };
            }

            public Dictionary<string, object> DelIds(params string[] ids)
            {
                var a = Args();
                a["action"] = "delete";
                a["slide_id"] = SlideId;
                a["shape_ids"] = ids;
                return a;
            }

            public Dictionary<string, object> DelShape(string shapeId)
            {
                var a = Args();
                a["action"] = "delete";
                a["slide_id"] = SlideId;
                a["shape_id"] = shapeId;
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

            public bool TryTwoBoxes(out string a, out string b, out string error)
            {
                a = b = null;
                if (!TryWipeApply(ContentHtml.TwoTextboxes(), out List<string> ids, out error) || ids.Count < 2)
                {
                    error = error ?? "两框不足";
                    return false;
                }

                a = ids[0];
                b = ids[1];
                return true;
            }

            public bool TryThreeBoxes(out string a, out string b, out string c, out string error)
            {
                a = b = c = null;
                if (!TryWipeApply(ContentHtml.ThreeTextboxes(), out List<string> ids, out error) || ids.Count < 3)
                {
                    error = error ?? "三框不足";
                    return false;
                }

                a = ids[0];
                b = ids[1];
                c = ids[2];
                return true;
            }

            public bool TryBoxAnd(string otherHtml, out string box, out string other, out string error)
            {
                box = other = null;
                if (!TryWipeApply(ContentHtml.TextboxCreate("留框", 8, 8, 28, 10), out List<string> boxes, out error)
                    || boxes.Count < 1)
                {
                    return false;
                }

                box = boxes[0];
                if (!TryApply(otherHtml, out List<string> more, out error) || more.Count < 1)
                {
                    return false;
                }

                other = more[0];
                return true;
            }

            public bool TrySeedGroup(out string groupId, out string a, out string b, out string error)
            {
                groupId = null;
                if (!TryTwoBoxes(out a, out b, out error))
                {
                    return false;
                }

                var g = Args();
                g["action"] = "group";
                g["slide_id"] = SlideId;
                g["shape_ids"] = new[] { a, b };
                ToolResult r = Invoke(g);
                if (r == null || !r.Success)
                {
                    error = r == null ? "group 无回包" : r.Error;
                    return false;
                }

                groupId = GetString(r, "group_shape_id");
                if (string.IsNullOrEmpty(groupId))
                {
                    error = "未回 group_shape_id";
                    return false;
                }

                return true;
            }

            public bool TryNestedLeaf(out string leaf, out string error)
            {
                leaf = null;
                if (!TryWipeApply(ContentHtml.NestedGroups(), out _, out error))
                {
                    return false;
                }

                if (!TryRead(SlideId, null, out PptHtmlReadResult page, out error))
                {
                    return false;
                }

                PptHtmlShapeNode outer = ContentGroupSession.FindGroup(page);
                if (outer == null)
                {
                    error = "顶层无组";
                    return false;
                }

                if (!TryRead(SlideId, outer.ShapeId, out PptHtmlReadResult expand, out error))
                {
                    return false;
                }

                PptHtmlShapeNode inner = ContentGroupSession.FindChildGroup(ContentGroupSession.FindByIdDeep(expand, outer.ShapeId)
                    ?? ContentGroupSession.FindGroup(expand));
                if (inner == null)
                {
                    inner = ContentGroupSession.FindChildGroup(ContentGroupSession.FindGroup(expand));
                }

                if (inner == null)
                {
                    error = "无内组";
                    return false;
                }

                if (!TryRead(SlideId, inner.ShapeId, out PptHtmlReadResult inn, out error))
                {
                    return false;
                }

                leaf = FindTextId(inn, "内组甲") ?? FirstLeaf(inn);
                if (string.IsNullOrEmpty(leaf))
                {
                    error = "内组无叶子";
                    return false;
                }

                return true;
            }

            public bool TrySecondBox(out string id, out string error)
            {
                id = null;
                if (!EnsureSlide2(out error))
                {
                    return false;
                }

                Wipe(SlideId2);
                if (!TryApplyOn(SlideId2, ContentHtml.TextboxCreate("第二页", 10, 10, 30, 12), out List<string> ids, out error)
                    || ids.Count < 1)
                {
                    return false;
                }

                id = ids[0];
                return true;
            }

            /// <summary>第 2 页上一个本页没有的 com id（产品按 slide_id 定页，同号会当成删本页）。</summary>
            public bool TrySecondForeignBox(out string id, out string error)
            {
                id = null;
                if (!EnsureSlide2(out error))
                {
                    return false;
                }

                Wipe(SlideId2);
                if (!TryApplyOn(SlideId2, ContentHtml.ThreeTextboxes(), out List<string> ids, out error)
                    || ids.Count < 3)
                {
                    error = error ?? "第2页三框不足";
                    return false;
                }

                for (int i = 0; i < ids.Count; i++)
                {
                    if (!ComHas(ids[i]))
                    {
                        id = ids[i];
                        return true;
                    }
                }

                error = "第2页没有本页没有的 com id";
                return false;
            }

            public bool TryTwoPages(out string p1, out string p2, out string error)
            {
                p1 = p2 = null;
                if (!TryWipeApply(ContentHtml.TextboxCreate("页一", 8, 8, 28, 10), out List<string> a, out error)
                    || a.Count < 1)
                {
                    return false;
                }

                p1 = a[0];
                return TrySecondBox(out p2, out error);
            }

            public bool TryWipeApply(string html, out List<string> ids, out string error)
            {
                Wipe(SlideId);
                return TryApply(html, out ids, out error);
            }

            public bool TryApply(string html, out List<string> ids, out string error)
            {
                return TryApplyOn(SlideId, html, out ids, out error);
            }

            public bool TryApplyOn(string slideId, string html, out List<string> ids, out string error)
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
                    ? PptHtmlWppApplier.TryApply(Pres, plan, "shp-wpp", out result, out error)
                    : PptHtmlPowerPointApplier.TryApply(
                        (PowerPoint.Presentation)Pres, plan, "shp-ppt", out result, out error);
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

            public bool TryRead(string slideId, string shapeId, out PptHtmlReadResult tree, out string error)
            {
                Refresh();
                if (Wpp)
                {
                    return PptHtmlWppReader.TryRead(
                        Pres, slideId, "shp-wpp", "wpp", shapeId, true, false, out tree, out error);
                }

                return PptHtmlPowerPointReader.TryRead(
                    (PowerPoint.Presentation)Pres, slideId, "shp-ppt", "ppt", shapeId, true, false, out tree, out error);
            }

            public bool HasId(string id)
            {
                return HasIdOn(SlideId, id);
            }

            public bool HasIdOn(string slideId, string id)
            {
                if (string.IsNullOrEmpty(id) || !TryRead(slideId, null, out PptHtmlReadResult tree, out _))
                {
                    return false;
                }

                return TreeHas(tree.Shapes, id);
            }

            public int CountIds()
            {
                if (!TryRead(SlideId, null, out PptHtmlReadResult tree, out _) || tree == null)
                {
                    return 0;
                }

                return CountSid(tree.Shapes);
            }

            public bool PageHasText(string text)
            {
                if (!TryRead(SlideId, null, out PptHtmlReadResult tree, out _))
                {
                    return false;
                }

                return TreeHasText(tree.Shapes, text);
            }

            public bool TryGeo(string id, out double left, out double top, out string error)
            {
                left = top = 0;
                error = null;
                if (!TryRead(SlideId, id, out PptHtmlReadResult tree, out error))
                {
                    return false;
                }

                PptHtmlShapeNode n = FindById(tree.Shapes, id);
                if (n == null)
                {
                    error = "读不到 " + id;
                    return false;
                }

                return TryParsePct(n.Style, "left", out left) && TryParsePct(n.Style, "top", out top);
            }

            public bool ComHas(string shapeId)
            {
                Refresh();
                if (!PptShapeId.TryParseShape(shapeId, out _, out int comId)
                    && !PptShapeId.TryParseShapeComId(shapeId, out comId))
                {
                    return false;
                }

                return FindCom(SlideId, comId) != null;
            }

            public bool TryHide(string shapeId, bool visible, out string error)
            {
                error = null;
                Refresh();
                if (!PptShapeId.TryParseShape(shapeId, out _, out int comId)
                    && !PptShapeId.TryParseShapeComId(shapeId, out comId))
                {
                    error = "id 无法解析";
                    return false;
                }

                object shape = FindCom(SlideId, comId);
                if (shape == null)
                {
                    error = "找不到形状";
                    return false;
                }

                try
                {
                    if (Wpp)
                    {
                        WppCom.TrySetProperty(shape, "Visible", visible ? -1 : 0);
                    }
                    else
                    {
                        ((PowerPoint.Shape)shape).Visible = visible
                            ? Office.MsoTriState.msoTrue
                            : Office.MsoTriState.msoFalse;
                    }

                    return true;
                }
                catch (Exception ex)
                {
                    error = ex.Message;
                    return false;
                }
            }

            public void Wipe(string slideId)
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

            private object FindCom(string slideId, int comId)
            {
                if (Wpp)
                {
                    object slide = FindSlide(slideId);
                    return FindComWpp(WppCom.GetProperty(slide, "Shapes"), comId);
                }

                return FindComPpt(FindSlidePpt(slideId).Shapes, comId);
            }

            private static object FindComWpp(object shapes, int comId)
            {
                if (shapes == null)
                {
                    return null;
                }

                int n = Convert.ToInt32(WppCom.GetProperty(shapes, "Count") ?? 0);
                for (int i = 1; i <= n; i++)
                {
                    object shape = WppCom.GetIndexed(shapes, i);
                    int sid = Convert.ToInt32(WppCom.GetProperty(shape, "Id") ?? 0);
                    if (sid == comId)
                    {
                        return shape;
                    }

                    int type = Convert.ToInt32(WppCom.GetProperty(shape, "Type") ?? 0);
                    if (type == 6)
                    {
                        object inner = FindComWpp(WppCom.GetProperty(shape, "GroupItems"), comId);
                        if (inner != null)
                        {
                            return inner;
                        }
                    }
                }

                return null;
            }

            private static PowerPoint.Shape FindComPpt(PowerPoint.Shapes shapes, int comId)
            {
                for (int i = 1; i <= shapes.Count; i++)
                {
                    PowerPoint.Shape shape = shapes[i];
                    if (shape.Id == comId)
                    {
                        return shape;
                    }

                    if ((int)shape.Type == 6)
                    {
                        PowerPoint.Shape found = FindComPptGroup(shape, comId);
                        if (found != null)
                        {
                            return found;
                        }
                    }
                }

                return null;
            }

            private static PowerPoint.Shape FindComPptGroup(PowerPoint.Shape group, int comId)
            {
                try
                {
                    for (int i = 1; i <= group.GroupItems.Count; i++)
                    {
                        PowerPoint.Shape one = group.GroupItems[i];
                        if (one.Id == comId)
                        {
                            return one;
                        }

                        if ((int)one.Type == 6)
                        {
                            PowerPoint.Shape found = FindComPptGroup(one, comId);
                            if (found != null)
                            {
                                return found;
                            }
                        }
                    }
                }
                catch
                {
                }

                return null;
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

            private static bool TreeHas(IList<PptHtmlShapeNode> nodes, string id)
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

                    if (string.Equals(n.ShapeId, id, StringComparison.OrdinalIgnoreCase))
                    {
                        return true;
                    }

                    if (TreeHas(n.Children, id))
                    {
                        return true;
                    }
                }

                return false;
            }

            private static int CountSid(IList<PptHtmlShapeNode> nodes)
            {
                int c = 0;
                if (nodes == null)
                {
                    return 0;
                }

                for (int i = 0; i < nodes.Count; i++)
                {
                    PptHtmlShapeNode n = nodes[i];
                    if (n == null)
                    {
                        continue;
                    }

                    if (!string.IsNullOrEmpty(n.ShapeId)
                        && n.ShapeId.IndexOf("-s", StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        c++;
                    }

                    c += CountSid(n.Children);
                }

                return c;
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

                    if (!string.IsNullOrEmpty(n.Text) && n.Text.IndexOf(text, StringComparison.Ordinal) >= 0)
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

            private static PptHtmlShapeNode FindById(IList<PptHtmlShapeNode> nodes, string id)
            {
                if (nodes == null)
                {
                    return null;
                }

                for (int i = 0; i < nodes.Count; i++)
                {
                    PptHtmlShapeNode n = nodes[i];
                    if (n == null)
                    {
                        continue;
                    }

                    if (string.Equals(n.ShapeId, id, StringComparison.OrdinalIgnoreCase))
                    {
                        return n;
                    }

                    PptHtmlShapeNode c = FindById(n.Children, id);
                    if (c != null)
                    {
                        return c;
                    }
                }

                return null;
            }

            private static string FindTextId(PptHtmlReadResult tree, string text)
            {
                return FindTextId(tree != null ? tree.Shapes : null, text);
            }

            private static string FindTextId(IList<PptHtmlShapeNode> nodes, string text)
            {
                if (nodes == null)
                {
                    return null;
                }

                for (int i = 0; i < nodes.Count; i++)
                {
                    PptHtmlShapeNode n = nodes[i];
                    if (n == null)
                    {
                        continue;
                    }

                    if (!string.IsNullOrEmpty(n.Text) && n.Text.IndexOf(text, StringComparison.Ordinal) >= 0
                        && !string.IsNullOrEmpty(n.ShapeId))
                    {
                        return n.ShapeId;
                    }

                    string c = FindTextId(n.Children, text);
                    if (c != null)
                    {
                        return c;
                    }
                }

                return null;
            }

            private static string FirstLeaf(PptHtmlReadResult tree)
            {
                return FirstLeaf(tree != null ? tree.Shapes : null);
            }

            private static string FirstLeaf(IList<PptHtmlShapeNode> nodes)
            {
                if (nodes == null)
                {
                    return null;
                }

                for (int i = 0; i < nodes.Count; i++)
                {
                    PptHtmlShapeNode n = nodes[i];
                    if (n == null)
                    {
                        continue;
                    }

                    if (!string.Equals(n.ShapeType, "group", StringComparison.OrdinalIgnoreCase)
                        && !string.IsNullOrEmpty(n.ShapeId)
                        && n.ShapeId.IndexOf("-s", StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        return n.ShapeId;
                    }

                    string c = FirstLeaf(n.Children);
                    if (c != null)
                    {
                        return c;
                    }
                }

                return null;
            }

            private static bool TryParsePct(string style, string key, out double value)
            {
                value = 0;
                if (string.IsNullOrEmpty(style))
                {
                    return false;
                }

                Match m = Regex.Match(
                    style,
                    key + "\\s*:\\s*([0-9.]+)%",
                    RegexOptions.IgnoreCase);
                return m.Success
                    && double.TryParse(m.Groups[1].Value, NumberStyles.Float, CultureInfo.InvariantCulture, out value);
            }
        }
    }
}
