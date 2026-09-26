using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Threading.Tasks;
using WordAddIn1;
using WordAddIn1.OpenFiles;
using WordAddIn1.PresentationHost;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PptHtmlContentRoundtripTest
{
    /// <summary>F_capture_document_page_image PPT/WPP 整页与页九宫格。前缀 pg-。</summary>
    internal static class PageImageRunner
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
                if (f.IndexOf("pg-", StringComparison.OrdinalIgnoreCase) >= 0
                    || string.Equals(f, "pg", StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
            }

            return false;
        }

        public static void Run(TestRun run, IList<string> nameFilters, bool useWpp)
        {
            Console.WriteLine(useWpp ? "整页截图 · WPP" : "整页截图 · PPT");
            RunParse(run, nameFilters);
            if (NeedCom(nameFilters))
            {
                RunCom(run, nameFilters, useWpp);
            }

            Console.WriteLine("整页截图 通过 " + run.CasesPassed
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
                "pg-one-", "pg-mid-", "pg-last-", "pg-page-str", "pg-echo-", "pg-size-",
                "pg-hidden-", "pg-after-", "pg-empty-shape", "pg-grid-", "pg-span-one",
                "pg-neg-over", "pg-neg-from-over", "pg-neg-to-over", "pg-neg-after-ok",
                "pg-neg-fail-no", "pg-path-"
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
                if (string.Equals(f, "pg", StringComparison.OrdinalIgnoreCase)
                    || string.Equals(f, "pg-", StringComparison.OrdinalIgnoreCase))
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

        private static void RunParse(TestRun run, IList<string> filters)
        {
            Console.WriteLine("--- A 渠道 / 互斥 ---");
            ChannelRegistry.Register(new KindStub("pg-stub-ppt", ChannelKind.Ppt), false);
            ChannelRegistry.Register(new KindStub("pg-stub-word", ChannelKind.Word), false);
            ChannelRegistry.Register(new KindStub("pg-stub-excel", ChannelKind.Excel), false);

            ExpectFail(run, filters, "pg-neg-no-channel", new Dictionary<string, object>(), null);
            ExpectFail(run, filters, "pg-neg-no-page", Args("pg-stub-ppt"), "page_number");

            var word = Args("pg-stub-word");
            word["page_from"] = 1;
            word["page_to"] = 2;
            ExpectFail(run, filters, "pg-neg-word", word, "page_from");

            var excel = Args("pg-stub-excel");
            excel["page_number"] = 1;
            ExpectFail(run, filters, "pg-neg-excel", excel, "page_number");

            var shape = Args("pg-stub-ppt");
            shape["page_number"] = 1;
            shape["shape_id"] = "sid1-s2";
            ExpectFail(run, filters, "pg-neg-shape-id", shape, "F_capture_ppt_shape");

            var shapeIds = Args("pg-stub-ppt");
            shapeIds["page_number"] = 1;
            shapeIds["shape_ids"] = new string[0];
            ExpectFail(run, filters, "pg-neg-shape-ids", shapeIds, "F_capture_ppt_shape");

            ExpectFail(run, filters, "pg-neg-slide-id", Page("pg-stub-ppt", 1, "slide_id", "256"), "slide_id");
            ExpectFail(
                run,
                filters,
                "pg-neg-page-numbers",
                Page("pg-stub-ppt", 1, "page_numbers", new[] { 1, 2 }),
                "page_numbers");
            ExpectFail(run, filters, "pg-neg-storage", Page("pg-stub-ppt", 1, "storage_doc_uuid", "x"), "storage_doc_uuid");
            ExpectFail(run, filters, "pg-neg-source-channel", Page("pg-stub-ppt", 1, "source_channel_id", "p2"), "channel_id");
            ExpectFail(run, filters, "pg-neg-name-kb", Page("pg-stub-ppt", 1, "document_name", "x"), "名");
            ExpectFail(run, filters, "pg-neg-sheet", Page("pg-stub-ppt", 1, "sheet", "Sheet1"), "sheet");
            ExpectFail(run, filters, "pg-neg-range", Page("pg-stub-ppt", 1, "range", "A1:B2"), "range");

            var both = Args("pg-stub-ppt");
            both["page_number"] = 1;
            both["page_from"] = 1;
            both["page_to"] = 2;
            ExpectFail(run, filters, "pg-neg-page-and-span", both, "不能同时传");

            var fromOnly = Args("pg-stub-ppt");
            fromOnly["page_from"] = 1;
            ExpectFail(run, filters, "pg-neg-from-only", fromOnly, "成对");

            var toOnly = Args("pg-stub-ppt");
            toOnly["page_to"] = 3;
            ExpectFail(run, filters, "pg-neg-to-only", toOnly, "成对");

            var gt = Args("pg-stub-ppt");
            gt["page_from"] = 5;
            gt["page_to"] = 2;
            ExpectFail(run, filters, "pg-neg-from-gt-to", gt, "不能大于");

            var ten = Args("pg-stub-ppt");
            ten["page_from"] = 1;
            ten["page_to"] = 10;
            ExpectFail(run, filters, "pg-neg-span-10", ten, "9");

            var zero = Args("pg-stub-ppt");
            zero["page_number"] = 0;
            ExpectFail(run, filters, "pg-neg-page-0", zero, "≥1");

            var bad = Args("pg-stub-ppt");
            bad["page_number"] = "abc";
            ExpectFail(run, filters, "pg-neg-page-not-int", bad, "整数");
        }

        private static void RunCom(TestRun run, IList<string> filters, bool useWpp)
        {
            Console.WriteLine(useWpp ? "--- B–E COM（WPP）---" : "--- B–E COM ---");
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
                    run.CaseFail("pg-", "渠道拿不到活稿");
                    return;
                }

                RunBlocks(run, filters, new Deck(ch, live, wpp: false));
            }
            catch (Exception ex)
            {
                run.CaseFail("pg-", "COM 失败: " + ex.Message);
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

        private static void RunComWpp(TestRun run, IList<string> filters)
        {
            object app = null;
            object pres = null;
            try
            {
                if (!WppCom.TryCreateApplication(out app, out string startErr))
                {
                    run.CaseFail("pg-", startErr);
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
                    run.CaseFail("pg-", "WPP Add 失败");
                    return;
                }

                WppChannel ch = ChannelRegistry.CreateOrGetWpp(pres, "", false);
                if (!ch.TryGetLivePresentation(out object live) || live == null)
                {
                    run.CaseFail("pg-", "渠道拿不到活稿");
                    return;
                }

                RunBlocks(run, filters, new Deck(ch, live, wpp: true));
            }
            catch (Exception ex)
            {
                run.CaseFail("pg-", "COM 失败: " + ex.Message);
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
            RunOnePage(run, filters, deck);
            RunThreePage(run, filters, deck);
            RunTenPage(run, filters, deck);
        }

        private static void RunOnePage(TestRun run, IList<string> filters, Deck deck)
        {
            if (!NeedOne(filters))
            {
                return;
            }

            if (!deck.EnsureCount(1, out string err))
            {
                FailNeeded(run, filters, err,
                    "pg-one-first", "pg-echo-kind", "pg-echo-channel", "pg-echo-no-layout",
                    "pg-size-ok", "pg-after-apply", "pg-empty-shape-id",
                    "pg-path-write", "pg-path-ext-fix", "pg-path-still-b64", "pg-path-omit");
                return;
            }

            if (Want(filters, "pg-one-first"))
            {
                ToolResult r = Invoke(deck.Page(1));
                var f = OkImage(r);
                if (GetInt(r, "page_number") != 1)
                {
                    f.Add("page_number 应为 1");
                }

                if (GetInt(r, "slide_count") != 1)
                {
                    f.Add("slide_count 应为 1");
                }

                if (string.IsNullOrEmpty(GetString(r, "slide_id")))
                {
                    f.Add("应有 slide_id");
                }

                Finish(run, "pg-one-first", f);
            }

            if (Want(filters, "pg-echo-kind"))
            {
                ToolResult r = Invoke(deck.Page(1));
                var f = OkImage(r);
                string want = deck.Wpp ? "wpp" : "ppt";
                if (!string.Equals(GetString(r, "source_kind"), want, StringComparison.OrdinalIgnoreCase))
                {
                    f.Add("source_kind 应为 " + want);
                }

                Finish(run, "pg-echo-kind", f);
            }

            if (Want(filters, "pg-echo-channel"))
            {
                ToolResult r = Invoke(deck.Page(1));
                var f = OkImage(r);
                string want = ChannelRegistry.ToPublicId(deck.Channel.ChannelId) ?? deck.Channel.ChannelId;
                if (!string.Equals(GetString(r, "channel_id"), want, StringComparison.Ordinal))
                {
                    f.Add("channel_id 应为 " + want + " 实际 " + GetString(r, "channel_id"));
                }

                Finish(run, "pg-echo-channel", f);
            }

            if (Want(filters, "pg-echo-no-layout"))
            {
                ToolResult r = Invoke(deck.Page(1));
                var f = OkImage(r);
                if (string.Equals(GetString(r, "layout"), "contact_sheet", StringComparison.OrdinalIgnoreCase))
                {
                    f.Add("单张不应带 layout=contact_sheet");
                }

                if (GetPages(r).Count > 0)
                {
                    f.Add("单张不应带 pages[]");
                }

                Finish(run, "pg-echo-no-layout", f);
            }

            if (Want(filters, "pg-size-ok"))
            {
                ToolResult r = Invoke(deck.Page(1));
                var f = OkImage(r);
                int w = GetInt(r, "width_px");
                int h = GetInt(r, "height_px");
                if (Math.Max(w, h) > 800)
                {
                    f.Add("长边应 ≤800 实际 " + Math.Max(w, h));
                }

                string fmt = (GetString(r, "format") ?? "").ToLowerInvariant();
                if (fmt != "png" && fmt != "jpeg" && fmt != "jpg")
                {
                    f.Add("format 应为 png 或 jpeg");
                }

                Finish(run, "pg-size-ok", f);
            }

            if (Want(filters, "pg-after-apply"))
            {
                if (!deck.TryApplyBox(1, out err))
                {
                    run.CaseFail("pg-after-apply", err);
                }
                else
                {
                    Finish(run, "pg-after-apply", OkImage(Invoke(deck.Page(1))));
                }
            }

            if (Want(filters, "pg-empty-shape-id"))
            {
                var a = deck.Page(1);
                a["shape_id"] = "";
                Finish(run, "pg-empty-shape-id", OkImage(Invoke(a)));
            }

            RunPath(run, filters, deck);
        }

        private static void RunPath(TestRun run, IList<string> filters, Deck deck)
        {
            if (!Want(filters, "pg-path-write") && !Want(filters, "pg-path-ext-fix")
                && !Want(filters, "pg-path-still-b64") && !Want(filters, "pg-path-omit"))
            {
                return;
            }

            string dir = Path.Combine(Path.GetTempPath(), "EasyWrite", "pg-test", Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(dir);
            try
            {
                if (Want(filters, "pg-path-write"))
                {
                    string png = Path.Combine(dir, "pg-one.png");
                    var a = deck.Page(1);
                    a["path"] = png;
                    ToolResult r = Invoke(a);
                    var f = OkImage(r);
                    if (string.IsNullOrEmpty(GetString(r, "path")))
                    {
                        f.Add("data.path 应非空");
                    }

                    if (!File.Exists(png) || new FileInfo(png).Length == 0)
                    {
                        f.Add("应写出非空 png");
                    }

                    Finish(run, "pg-path-write", f);
                }

                if (Want(filters, "pg-path-ext-fix") || Want(filters, "pg-path-still-b64"))
                {
                    string bmp = Path.Combine(dir, "pg-fix.bmp");
                    var a = deck.Page(1);
                    a["path"] = bmp;
                    ToolResult r = Invoke(a);
                    var f = OkImage(r);
                    string written = GetString(r, "path");
                    bool exists = FileExistsNonEmpty(written)
                        || FileExistsNonEmpty(Path.ChangeExtension(bmp, ".png"))
                        || FileExistsNonEmpty(Path.ChangeExtension(bmp, ".jpg"))
                        || FileExistsNonEmpty(Path.ChangeExtension(bmp, ".jpeg"));
                    if (Want(filters, "pg-path-ext-fix"))
                    {
                        var fx = new List<string>(f);
                        if (!exists)
                        {
                            fx.Add("改后缀后文件应存在");
                        }

                        Finish(run, "pg-path-ext-fix", fx);
                    }

                    if (Want(filters, "pg-path-still-b64"))
                    {
                        Finish(run, "pg-path-still-b64", f);
                    }
                }

                if (Want(filters, "pg-path-omit"))
                {
                    ToolResult r = Invoke(deck.Page(1));
                    var f = OkImage(r);
                    Finish(run, "pg-path-omit", f);
                }
            }
            finally
            {
                try
                {
                    Directory.Delete(dir, true);
                }
                catch
                {
                }
            }
        }

        private static void RunThreePage(TestRun run, IList<string> filters, Deck deck)
        {
            if (!NeedThree(filters))
            {
                return;
            }

            if (!deck.EnsureCount(3, out string err))
            {
                FailNeeded(run, filters, err,
                    "pg-one-of-3", "pg-mid-of-3", "pg-last-of-3", "pg-page-str", "pg-hidden-ok",
                    "pg-grid-2", "pg-grid-3", "pg-grid-pages-ids", "pg-grid-echo-span",
                    "pg-grid-no-shape-mode", "pg-span-one-is-single", "pg-grid-has-image",
                    "pg-grid-order", "pg-neg-over-fail", "pg-neg-from-over-fail",
                    "pg-neg-to-over-fail", "pg-neg-after-ok", "pg-neg-fail-no-b64",
                    "pg-neg-over-keeps-deck");
                return;
            }

            string id1 = deck.IdAt(1);
            string id2 = deck.IdAt(2);
            string id3 = deck.IdAt(3);

            if (Want(filters, "pg-one-of-3"))
            {
                ToolResult r = Invoke(deck.Page(1));
                var f = OkImage(r);
                if (GetString(r, "slide_id") != id1)
                {
                    f.Add("slide_id 应为 " + id1);
                }

                if (GetInt(r, "slide_count") != 3)
                {
                    f.Add("slide_count 应为 3");
                }

                Finish(run, "pg-one-of-3", f);
            }

            if (Want(filters, "pg-mid-of-3"))
            {
                ToolResult r = Invoke(deck.Page(2));
                var f = OkImage(r);
                if (GetString(r, "slide_id") != id2)
                {
                    f.Add("slide_id 应为 " + id2);
                }

                Finish(run, "pg-mid-of-3", f);
            }

            if (Want(filters, "pg-last-of-3"))
            {
                ToolResult r = Invoke(deck.Page(3));
                var f = OkImage(r);
                if (GetString(r, "slide_id") != id3)
                {
                    f.Add("slide_id 应为 " + id3);
                }

                Finish(run, "pg-last-of-3", f);
            }

            if (Want(filters, "pg-page-str"))
            {
                var a = deck.Args();
                a["page_number"] = "2";
                ToolResult r = Invoke(a);
                var f = OkImage(r);
                if (GetInt(r, "page_number") != 2 || GetString(r, "slide_id") != id2)
                {
                    f.Add("字符串 2 应等同第 2 页");
                }

                Finish(run, "pg-page-str", f);
            }

            if (Want(filters, "pg-hidden-ok"))
            {
                if (!deck.TryHide(2, true, out err))
                {
                    run.CaseFail("pg-hidden-ok", err);
                }
                else
                {
                    ToolResult r = Invoke(deck.Page(2));
                    var f = OkImage(r);
                    if (GetString(r, "slide_id") != id2)
                    {
                        f.Add("隐藏页 slide_id 应为 " + id2);
                    }

                    Finish(run, "pg-hidden-ok", f);
                }
            }

            if (Want(filters, "pg-grid-2"))
            {
                ToolResult r = Invoke(deck.Span(1, 2));
                var f = OkSheet(r, 2);
                Finish(run, "pg-grid-2", f);
            }

            if (Want(filters, "pg-grid-3"))
            {
                ToolResult r = Invoke(deck.Span(1, 3));
                var f = OkSheet(r, 3);
                IList<Dictionary<string, object>> pages = GetPages(r);
                if (pages.Count == 3
                    && (GetDictInt(pages[0], "index") != 1
                        || GetDictInt(pages[1], "index") != 2
                        || GetDictInt(pages[2], "index") != 3))
                {
                    f.Add("pages.index 应为 1,2,3");
                }

                Finish(run, "pg-grid-3", f);
            }

            if (Want(filters, "pg-grid-pages-ids"))
            {
                ToolResult r = Invoke(deck.Span(1, 3));
                var f = OkSheet(r, 3);
                IList<Dictionary<string, object>> pages = GetPages(r);
                if (pages.Count == 3
                    && (GetDictString(pages[0], "slide_id") != id1
                        || GetDictString(pages[1], "slide_id") != id2
                        || GetDictString(pages[2], "slide_id") != id3))
                {
                    f.Add("pages[].slide_id 应对齐 COM");
                }

                Finish(run, "pg-grid-pages-ids", f);
            }

            if (Want(filters, "pg-grid-echo-span"))
            {
                ToolResult r = Invoke(deck.Span(2, 3));
                var f = OkSheet(r, 2);
                if (GetInt(r, "page_from") != 2 || GetInt(r, "page_to") != 3 || GetInt(r, "page_number") != 2)
                {
                    f.Add("应回显 from=2 to=3 page_number=2");
                }

                Finish(run, "pg-grid-echo-span", f);
            }

            if (Want(filters, "pg-grid-no-shape-mode"))
            {
                ToolResult r = Invoke(deck.Span(1, 2));
                var f = OkSheet(r, 2);
                string mode = GetString(r, "capture_mode");
                if (mode.IndexOf("shape", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    f.Add("页宫格不应带 capture_mode=" + mode);
                }

                Finish(run, "pg-grid-no-shape-mode", f);
            }

            if (Want(filters, "pg-span-one-is-single"))
            {
                ToolResult r = Invoke(deck.Span(2, 2));
                var f = OkImage(r);
                if (string.Equals(GetString(r, "layout"), "contact_sheet", StringComparison.OrdinalIgnoreCase))
                {
                    f.Add("一段 1 页应收成单张");
                }

                if (GetString(r, "slide_id") != id2)
                {
                    f.Add("slide_id 应为 " + id2);
                }

                Finish(run, "pg-span-one-is-single", f);
            }

            if (Want(filters, "pg-grid-has-image"))
            {
                Finish(run, "pg-grid-has-image", OkSheet(Invoke(deck.Span(1, 2)), 2));
            }

            if (Want(filters, "pg-grid-order"))
            {
                ToolResult r = Invoke(deck.Span(2, 3));
                var f = OkSheet(r, 2);
                IList<Dictionary<string, object>> pages = GetPages(r);
                if (pages.Count == 2
                    && (GetDictInt(pages[0], "index") != 2 || GetDictInt(pages[1], "index") != 3))
                {
                    f.Add("pages 应为升序 2,3");
                }

                Finish(run, "pg-grid-order", f);
            }

            RunOver(run, filters, deck, id1, id2, id3);
        }

        private static void RunOver(TestRun run, IList<string> filters, Deck deck, string id1, string id2, string id3)
        {
            if (Want(filters, "pg-neg-over-fail"))
            {
                ToolResult r = Invoke(deck.Page(4));
                var f = ExpectFailImage(r, "超出范围");
                if (GetInt(r, "slide_count") != 0 && GetInt(r, "slide_count") != 3)
                {
                    f.Add("越界 data.slide_count 若有应为 3");
                }

                Finish(run, "pg-neg-over-fail", f);
            }

            if (Want(filters, "pg-neg-from-over-fail"))
            {
                Finish(run, "pg-neg-from-over-fail", ExpectFailImage(Invoke(deck.Span(5, 6)), "超出范围"));
            }

            if (Want(filters, "pg-neg-to-over-fail"))
            {
                Finish(run, "pg-neg-to-over-fail", ExpectFailImage(Invoke(deck.Span(2, 5)), "超出范围"));
            }

            if (Want(filters, "pg-neg-after-ok"))
            {
                Invoke(deck.Page(99));
                Finish(run, "pg-neg-after-ok", OkImage(Invoke(deck.Page(1))));
            }

            if (Want(filters, "pg-neg-fail-no-b64"))
            {
                ToolResult r = Invoke(deck.Page(4));
                var f = new List<string>();
                if (r != null && r.Success)
                {
                    f.Add("应失败");
                }

                if (HasImage(r))
                {
                    f.Add("失败不应带可用图");
                }

                Finish(run, "pg-neg-fail-no-b64", f);
            }

            if (Want(filters, "pg-neg-over-keeps-deck"))
            {
                Invoke(deck.Page(4));
                var f = new List<string>();
                if (deck.Count() != 3 || deck.IdAt(1) != id1 || deck.IdAt(2) != id2 || deck.IdAt(3) != id3)
                {
                    f.Add("越界后页数或 id 变了");
                }

                Finish(run, "pg-neg-over-keeps-deck", f);
            }
        }

        private static void RunTenPage(TestRun run, IList<string> filters, Deck deck)
        {
            if (!NeedTen(filters))
            {
                return;
            }

            if (!deck.EnsureCount(10, out string err))
            {
                FailNeeded(run, filters, err, "pg-grid-9", "pg-grid-seg", "pg-grid-tail", "pg-grid-count");
                return;
            }

            string id10 = deck.IdAt(10);

            if (Want(filters, "pg-grid-9"))
            {
                Finish(run, "pg-grid-9", OkSheet(Invoke(deck.Span(1, 9)), 9));
            }

            if (Want(filters, "pg-grid-seg"))
            {
                ToolResult r = Invoke(deck.Span(5, 9));
                var f = OkSheet(r, 5);
                if (GetInt(r, "page_from") != 5 || GetInt(r, "page_to") != 9)
                {
                    f.Add("应回显 5–9");
                }

                IList<Dictionary<string, object>> pages = GetPages(r);
                if (pages.Count > 0 && GetDictInt(pages[0], "index") != 5)
                {
                    f.Add("pages[0].index 应为 5");
                }

                Finish(run, "pg-grid-seg", f);
            }

            if (Want(filters, "pg-grid-tail"))
            {
                ToolResult r = Invoke(deck.Span(8, 10));
                var f = OkSheet(r, 3);
                IList<Dictionary<string, object>> pages = GetPages(r);
                if (pages.Count != 3
                    || GetDictInt(pages[2], "index") != 10
                    || GetDictString(pages[2], "slide_id") != id10)
                {
                    f.Add("末格应为 index=10 且 slide_id 对");
                }

                Finish(run, "pg-grid-tail", f);
            }

            if (Want(filters, "pg-grid-count"))
            {
                ToolResult r = Invoke(deck.Span(1, 9));
                var f = OkSheet(r, 9);
                if (GetInt(r, "slide_count") != 10)
                {
                    f.Add("slide_count 应为稿页数 10");
                }

                Finish(run, "pg-grid-count", f);
            }
        }

        private static bool NeedOne(IList<string> filters)
        {
            return Want(filters, "pg-one-first") || Want(filters, "pg-echo-") || Want(filters, "pg-size-ok")
                || Want(filters, "pg-after-apply") || Want(filters, "pg-empty-shape-id") || Want(filters, "pg-path-");
        }

        private static bool NeedThree(IList<string> filters)
        {
            return Want(filters, "pg-one-of-3") || Want(filters, "pg-mid-of-3") || Want(filters, "pg-last-of-3")
                || Want(filters, "pg-page-str") || Want(filters, "pg-hidden-ok") || Want(filters, "pg-grid-2")
                || Want(filters, "pg-grid-3") || Want(filters, "pg-grid-pages-ids") || Want(filters, "pg-grid-echo-span")
                || Want(filters, "pg-grid-no-shape-mode") || Want(filters, "pg-span-one-is-single")
                || Want(filters, "pg-grid-has-image") || Want(filters, "pg-grid-order")
                || Want(filters, "pg-neg-over") || Want(filters, "pg-neg-from-over") || Want(filters, "pg-neg-to-over")
                || Want(filters, "pg-neg-after-ok") || Want(filters, "pg-neg-fail-no") || Want(filters, "pg-neg-over-keeps");
        }

        private static bool NeedTen(IList<string> filters)
        {
            return Want(filters, "pg-grid-9") || Want(filters, "pg-grid-seg")
                || Want(filters, "pg-grid-tail") || Want(filters, "pg-grid-count");
        }

        private static void ExpectFail(
            TestRun run,
            IList<string> filters,
            string id,
            Dictionary<string, object> args,
            string contains)
        {
            if (!Want(filters, id))
            {
                return;
            }

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

            if (HasImage(r))
            {
                fails.Add("失败不应带图");
            }

            Finish(run, id, fails);
        }

        private static List<string> ExpectFailImage(ToolResult r, string contains)
        {
            var f = new List<string>();
            if (r == null || r.Success)
            {
                f.Add("应失败");
            }
            else if ((r.Error ?? "").IndexOf(contains, StringComparison.OrdinalIgnoreCase) < 0)
            {
                f.Add("文案应含「" + contains + "」实际 " + (r.Error ?? ""));
            }

            if (HasImage(r))
            {
                f.Add("失败不应带可用图");
            }

            return f;
        }

        private static List<string> OkImage(ToolResult r)
        {
            var f = new List<string>();
            if (r == null || !r.Success)
            {
                f.Add(r == null ? "无回包" : (r.Error ?? "失败"));
                return f;
            }

            if (!HasImage(r))
            {
                f.Add("应有非空 image_base64");
            }

            if (GetInt(r, "width_px") <= 0 || GetInt(r, "height_px") <= 0)
            {
                f.Add("宽高应 > 0");
            }

            return f;
        }

        private static List<string> OkSheet(ToolResult r, int pageCount)
        {
            var f = OkImage(r);
            if (!string.Equals(GetString(r, "layout"), "contact_sheet", StringComparison.OrdinalIgnoreCase))
            {
                f.Add("layout 应为 contact_sheet");
            }

            if (GetPages(r).Count != pageCount)
            {
                f.Add("pages 长度应为 " + pageCount + " 实际 " + GetPages(r).Count);
            }

            return f;
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

        private static Dictionary<string, object> Page(string channelId, int page, string extra, object extraVal)
        {
            var a = Args(channelId);
            a["page_number"] = page;
            a[extra] = extraVal;
            return a;
        }

        private static ToolResult Invoke(Dictionary<string, object> args)
        {
            var tools = new Dictionary<string, Func<Dictionary<string, object>, Task<ToolResult>>>(
                StringComparer.OrdinalIgnoreCase);
            F_CaptureDocumentPageImageTool.Register(tools, null);
            if (!tools.TryGetValue("F_capture_document_page_image", out Func<Dictionary<string, object>, Task<ToolResult>> fn))
            {
                return null;
            }

            return fn(args ?? new Dictionary<string, object>()).GetAwaiter().GetResult();
        }

        private static bool HasImage(ToolResult r)
        {
            if (r == null || r.Data == null)
            {
                return false;
            }

            var d = r.Data as Dictionary<string, object>;
            if (d == null || !d.TryGetValue("image_base64", out object v) || v == null)
            {
                return false;
            }

            return Convert.ToString(v).Length > 32;
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
                int n;
                return int.TryParse(Convert.ToString(v), NumberStyles.Integer, CultureInfo.InvariantCulture, out n)
                    ? n
                    : 0;
            }
        }

        private static IList<Dictionary<string, object>> GetPages(ToolResult r)
        {
            var list = new List<Dictionary<string, object>>();
            var d = r != null ? r.Data as Dictionary<string, object> : null;
            if (d == null || !d.TryGetValue("pages", out object raw) || raw == null)
            {
                return list;
            }

            var enumerable = raw as IEnumerable;
            if (enumerable == null || raw is string)
            {
                return list;
            }

            foreach (object item in enumerable)
            {
                var one = item as Dictionary<string, object>;
                if (one != null)
                {
                    list.Add(one);
                }
            }

            return list;
        }

        private static string GetDictString(Dictionary<string, object> d, string key)
        {
            if (d == null || !d.TryGetValue(key, out object v) || v == null)
            {
                return "";
            }

            return Convert.ToString(v) ?? "";
        }

        private static int GetDictInt(Dictionary<string, object> d, string key)
        {
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
                int n;
                return int.TryParse(Convert.ToString(v), NumberStyles.Integer, CultureInfo.InvariantCulture, out n)
                    ? n
                    : 0;
            }
        }

        private static bool FileExistsNonEmpty(string path)
        {
            return !string.IsNullOrEmpty(path) && File.Exists(path) && new FileInfo(path).Length > 0;
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

            public Dictionary<string, object> Args()
            {
                return new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase)
                {
                    ["channel_id"] = Channel.ChannelId
                };
            }

            public Dictionary<string, object> Page(int pageNumber)
            {
                var a = Args();
                a["page_number"] = pageNumber;
                return a;
            }

            public Dictionary<string, object> Span(int from, int to)
            {
                var a = Args();
                a["page_from"] = from;
                a["page_to"] = to;
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

            public bool TryApplyBox(int index1, out string error)
            {
                error = null;
                Refresh();
                string slideId = IdAt(index1);
                if (!ContentHtml.TryParse(ContentHtml.TextboxCreate("整页截图"), slideId, out PptHtmlApplyPlan plan, out error))
                {
                    return false;
                }

                plan.AllowCreate = true;
                PptHtmlApplyResult result;
                return Wpp
                    ? PptHtmlWppApplier.TryApply(Pres, plan, "pg-wpp", out result, out error)
                    : PptHtmlPowerPointApplier.TryApply(
                        (PowerPoint.Presentation)Pres, plan, "pg-ppt", out result, out error);
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
