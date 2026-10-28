using System;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.Globalization;
using System.IO;
using System.Reflection;
using System.Security.Cryptography;
using System.Threading;
using System.Threading.Tasks;
using WordAddIn1;
using WordAddIn1.OpenFiles;
using WordAddIn1.PresentationHost;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PptHtmlContentRoundtripTest
{
    /// <summary>F_ppt_background。前缀 bg-。手验 bg-hand- 不进本运行器。</summary>
    internal static class BackgroundRunner
    {
        private const string SolidHex = "#1E5AA8";

        private static readonly HashSet<string> Scored = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        private static readonly string[] ComIds =
        {
            "bg-neg-bad-action", "bg-neg-apply-alias", "bg-neg-list-slide", "bg-neg-list-ids",
            "bg-neg-list-write", "bg-neg-get-no-slide", "bg-neg-get-write", "bg-neg-set-no-slide",
            "bg-neg-set-no-kind", "bg-neg-bad-kind", "bg-neg-solid-no-color", "bg-neg-solid-short",
            "bg-neg-solid-name", "bg-neg-solid-path", "bg-neg-pic-no-path", "bg-neg-pic-color",
            "bg-neg-clear-no-slide", "bg-neg-clear-write", "bg-neg-sid", "bg-neg-fit",
            "bg-neg-mix-stretch", "bg-neg-mix-tile", "bg-neg-align", "bg-neg-mirror", "bg-neg-scale",
            "bg-list-kinds", "bg-list-order", "bg-get-master", "bg-get-solid", "bg-get-theme",
            "bg-get-picture", "bg-get-other", "bg-get-shape", "bg-get-missing",
            "bg-set-solid", "bg-set-solid-lower", "bg-set-solid-over-pic", "bg-set-theme-fixed",
            "bg-set-other-page", "bg-set-over-other",
            "bg-set-pic-default", "bg-set-pic-abs", "bg-set-stretch", "bg-set-stretch-neg",
            "bg-set-stretch-omit", "bg-copy-pic",
            "bg-set-tile-default", "bg-set-tile", "bg-set-tile-mirror",
            "bg-clear", "bg-clear-master-kept", "bg-fail-keeps", "bg-missing-keeps",
            "bg-action-upper", "bg-echo-kind", "bg-echo-display", "bg-hidden",
            "bg-text-kept", "bg-trn-kept"
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
                if (f.IndexOf("bg-", StringComparison.OrdinalIgnoreCase) >= 0
                    || string.Equals(f, "bg", StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
            }

            return false;
        }

        public static void Run(TestRun run, IList<string> nameFilters, bool useWpp)
        {
            Scored.Clear();
            Console.WriteLine(useWpp ? "幻灯片背景 · WPP" : "幻灯片背景 · PPT");
            RunParse(run, nameFilters);
            if (NeedCom(nameFilters))
            {
                RunCom(run, nameFilters, useWpp);
            }

            Console.WriteLine("幻灯片背景 通过 " + run.CasesPassed
                + "  失败 " + run.CasesFailed
                + "  跳过 " + run.CasesSkipped);
        }

        private static void RunParse(TestRun run, IList<string> filters)
        {
            Console.WriteLine("--- A 渠道（stub）---");
            ChannelRegistry.Register(new KindStub("bg-stub-ppt", ChannelKind.Ppt), false);
            ChannelRegistry.Register(new KindStub("bg-stub-word", ChannelKind.Word), false);
            ChannelRegistry.Register(new KindStub("bg-stub-excel", ChannelKind.Excel), false);

            if (Want(filters, "bg-neg-no-channel"))
            {
                ExpectFail(run, "bg-neg-no-channel", new Dictionary<string, object>(), null);
            }

            if (Want(filters, "bg-neg-no-action"))
            {
                ExpectFail(run, "bg-neg-no-action", Args("bg-stub-ppt"), "list|get|set|clear");
            }

            if (Want(filters, "bg-neg-word"))
            {
                var a = Args("bg-stub-word");
                a["action"] = "list";
                ExpectFail(run, "bg-neg-word", a, "渠道");
            }

            if (Want(filters, "bg-neg-excel"))
            {
                var a = Args("bg-stub-excel");
                a["action"] = "list";
                ExpectFail(run, "bg-neg-excel", a, "渠道");
            }

            if (Want(filters, "bg-neg-pic-http"))
            {
                var a = Args("bg-stub-ppt");
                a["action"] = "set";
                a["slide_id"] = "999999";
                a["kind"] = "picture";
                a["path"] = "https://example.com/a.png";
                ExpectFail(run, "bg-neg-pic-http", a, "http");
            }

            if (Want(filters, "bg-neg-pic-data"))
            {
                var a = Args("bg-stub-ppt");
                a["action"] = "set";
                a["slide_id"] = "999999";
                a["kind"] = "picture";
                a["path"] = "data:image/png;base64,AAAA";
                ExpectFail(run, "bg-neg-pic-data", a, "data");
            }

            if (Want(filters, "bg-neg-pic-missing"))
            {
                var a = Args("bg-stub-ppt");
                a["action"] = "set";
                a["slide_id"] = "999999";
                a["kind"] = "picture";
                a["path"] = @"C:\bg-no-such-file.png";
                ExpectFail(run, "bg-neg-pic-missing", a, "不存在|不可用");
            }
        }

        private static void RunCom(TestRun run, IList<string> filters, bool useWpp)
        {
            Console.WriteLine(useWpp ? "--- COM（WPP）---" : "--- COM ---");
            if (!UserService.Instance.EnsureEffectiveRootInteractive(null))
            {
                FailNeeded(run, filters, "工作区未就绪");
                return;
            }

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

                var deck = new Deck(ch, live, false);
                if (!deck.Prepare(out string err))
                {
                    FailNeeded(run, filters, err ?? "铺页失败");
                    return;
                }

                RunCases(run, filters, deck);
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

                var deck = new Deck(ch, live, true);
                if (!deck.Prepare(out string err))
                {
                    FailNeeded(run, filters, err ?? "铺页失败");
                    return;
                }

                RunCases(run, filters, deck);
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

        private static void RunCases(TestRun run, IList<string> filters, Deck deck)
        {
            try
            {
                RunNeg(run, filters, deck);
                RunRead(run, filters, deck);
                RunSolid(run, filters, deck);
                RunPicture(run, filters, deck);
                RunTile(run, filters, deck);
                RunTail(run, filters, deck);
            }
            catch (Exception ex)
            {
                FailNeeded(run, filters, "COM 中断: " + ex.Message);
            }
        }

        private static void RunNeg(TestRun run, IList<string> filters, Deck deck)
        {
            if (Want(filters, "bg-neg-bad-action"))
            {
                var a = deck.Args();
                a["action"] = "add";
                ExpectFail(run, "bg-neg-bad-action", a, null);
            }

            if (Want(filters, "bg-neg-apply-alias"))
            {
                var a = deck.Args();
                a["action"] = "apply";
                a["slide_id"] = deck.IdS;
                a["kind"] = "solid";
                a["color"] = "#112233";
                ExpectFail(run, "bg-neg-apply-alias", a, null);
            }

            if (Want(filters, "bg-neg-list-slide"))
            {
                var a = deck.List();
                a["slide_id"] = deck.IdM;
                ExpectFail(run, "bg-neg-list-slide", a, "slide_id");
            }

            if (Want(filters, "bg-neg-list-ids"))
            {
                var a = deck.List();
                a["slide_ids"] = new[] { deck.IdM };
                ExpectFail(run, "bg-neg-list-ids", a, "slide_id");
            }

            if (Want(filters, "bg-neg-list-write"))
            {
                var a = deck.List();
                a["kind"] = "solid";
                ExpectFail(run, "bg-neg-list-write", a, null);
            }

            if (Want(filters, "bg-neg-get-no-slide"))
            {
                var a = deck.Args();
                a["action"] = "get";
                ExpectFail(run, "bg-neg-get-no-slide", a, "slide_id");
            }

            if (Want(filters, "bg-neg-get-write"))
            {
                var a = deck.Get(deck.IdM);
                a["color"] = "#112233";
                ExpectFail(run, "bg-neg-get-write", a, null);
            }

            if (Want(filters, "bg-neg-set-no-slide"))
            {
                var a = deck.Args();
                a["action"] = "set";
                a["kind"] = "solid";
                a["color"] = "#112233";
                ExpectFail(run, "bg-neg-set-no-slide", a, "slide_id");
            }

            if (Want(filters, "bg-neg-set-no-kind"))
            {
                var a = deck.Args();
                a["action"] = "set";
                a["slide_id"] = "999999";
                ExpectFail(run, "bg-neg-set-no-kind", a, "kind");
            }

            if (Want(filters, "bg-neg-bad-kind"))
            {
                var a = deck.SetSolid("999999", "#112233");
                a["kind"] = "gradient";
                ExpectFail(run, "bg-neg-bad-kind", a, "solid|picture");
            }

            if (Want(filters, "bg-neg-solid-no-color"))
            {
                var a = deck.Args();
                a["action"] = "set";
                a["slide_id"] = deck.IdS;
                a["kind"] = "solid";
                ExpectFail(run, "bg-neg-solid-no-color", a, "color");
            }

            if (Want(filters, "bg-neg-solid-short"))
            {
                ExpectFail(run, "bg-neg-solid-short", deck.SetSolid(deck.IdS, "#FFF"), "#RRGGBB");
            }

            if (Want(filters, "bg-neg-solid-name"))
            {
                ExpectFail(run, "bg-neg-solid-name", deck.SetSolid(deck.IdS, "red"), "#RRGGBB");
            }

            if (Want(filters, "bg-neg-solid-path"))
            {
                var a = deck.SetSolid(deck.IdS, "#112233");
                a["path"] = "bg-red.png";
                var f = new List<string>();
                ToolResult r = Invoke(a);
                if (r == null || r.Success)
                {
                    f.Add(r == null ? "无回包" : "应失败却成功");
                }

                ExpectHex(deck, deck.IdS, SolidHex, f);
                Finish(run, "bg-neg-solid-path", f);
            }

            if (Want(filters, "bg-neg-pic-no-path"))
            {
                var a = deck.Args();
                a["action"] = "set";
                a["slide_id"] = deck.IdS;
                a["kind"] = "picture";
                ExpectFail(run, "bg-neg-pic-no-path", a, "path");
            }

            if (Want(filters, "bg-neg-pic-color"))
            {
                var a = deck.SetPic(deck.IdS, "bg-red.png");
                a["color"] = "#112233";
                var f = new List<string>();
                ToolResult r = Invoke(a);
                if (r == null || r.Success)
                {
                    f.Add(r == null ? "无回包" : "应失败却成功");
                }

                ExpectHex(deck, deck.IdS, SolidHex, f);
                Finish(run, "bg-neg-pic-color", f);
            }

            if (Want(filters, "bg-neg-clear-no-slide"))
            {
                var a = deck.Args();
                a["action"] = "clear";
                ExpectFail(run, "bg-neg-clear-no-slide", a, "slide_id");
            }

            if (Want(filters, "bg-neg-clear-write"))
            {
                var a = deck.Clear(deck.IdS);
                a["color"] = "#112233";
                ExpectFail(run, "bg-neg-clear-write", a, null);
            }

            if (Want(filters, "bg-neg-sid"))
            {
                ExpectFail(run, "bg-neg-sid", deck.Get("sid256"), "slide_id");
            }

            if (Want(filters, "bg-neg-fit"))
            {
                var a = deck.SetPic(deck.IdS, "bg-red.png");
                a["fit"] = "cover";
                ExpectFail(run, "bg-neg-fit", a, "stretch|tile");
            }

            if (Want(filters, "bg-neg-mix-stretch"))
            {
                var a = deck.SetPic(deck.IdS, "bg-red.png");
                a["fit"] = "stretch";
                a["tile_scale_x"] = 40d;
                var f = new List<string>();
                ToolResult r = Invoke(a);
                if (r == null || r.Success)
                {
                    f.Add(r == null ? "无回包" : "应失败却成功");
                }

                ExpectHex(deck, deck.IdS, SolidHex, f);
                Finish(run, "bg-neg-mix-stretch", f);
            }

            if (Want(filters, "bg-neg-mix-tile"))
            {
                var a = deck.SetPic(deck.IdS, "bg-red.png");
                a["fit"] = "tile";
                a["stretch_left"] = 10d;
                var f = new List<string>();
                ToolResult r = Invoke(a);
                if (r == null || r.Success)
                {
                    f.Add(r == null ? "无回包" : "应失败却成功");
                }

                ExpectHex(deck, deck.IdS, SolidHex, f);
                Finish(run, "bg-neg-mix-tile", f);
            }

            if (Want(filters, "bg-neg-align"))
            {
                var a = deck.SetPic(deck.IdS, "bg-red.png");
                a["fit"] = "tile";
                a["tile_align"] = "middle";
                ExpectFail(run, "bg-neg-align", a, "tl");
            }

            if (Want(filters, "bg-neg-mirror"))
            {
                var a = deck.SetPic(deck.IdS, "bg-red.png");
                a["fit"] = "tile";
                a["tile_mirror"] = "both";
                ExpectFail(run, "bg-neg-mirror", a, "none|x|y|xy");
            }

            if (Want(filters, "bg-neg-scale"))
            {
                var a = deck.SetPic(deck.IdS, "bg-red.png");
                a["fit"] = "tile";
                a["tile_scale_x"] = 0d;
                var f = new List<string>();
                ToolResult r = Invoke(a);
                if (r == null || r.Success)
                {
                    f.Add(r == null ? "无回包" : "应失败却成功");
                }
                else if (!ContainsAny(r.Error, "大于 0"))
                {
                    f.Add("文案应含「大于 0」实际 " + (r.Error ?? ""));
                }

                ExpectHex(deck, deck.IdS, SolidHex, f);
                Finish(run, "bg-neg-scale", f);
            }
        }

        private static void RunRead(TestRun run, IList<string> filters, Deck deck)
        {
            if (Want(filters, "bg-list-kinds") || Want(filters, "bg-list-order"))
            {
                ToolResult listed = Invoke(deck.List());
                if (Want(filters, "bg-list-kinds"))
                {
                    var f = new List<string>();
                    if (!Ok(listed, f))
                    {
                        Finish(run, "bg-list-kinds", f);
                    }
                    else
                    {
                        string blob = Newtonsoft.Json.JsonConvert.SerializeObject(listed.Data);
                        if (blob.IndexOf("base64", StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            f.Add("回包含 base64");
                        }

                        CheckKind(Rows(listed), deck.IdM, true, "master", null, false, f, "母版页");
                        CheckKind(Rows(listed), deck.IdS, false, "solid", SolidHex, false, f, "纯色页");
                        CheckKind(Rows(listed), deck.IdP, false, "picture", null, true, f, "图片页");
                        CheckKind(Rows(listed), deck.IdO, false, "other", null, false, f, "渐变页");
                        CheckKind(Rows(listed), deck.IdShape, true, "master", null, false, f, "铺满形状页");
                        Finish(run, "bg-list-kinds", f);
                    }
                }

                if (Want(filters, "bg-list-order"))
                {
                    var f = new List<string>();
                    if (!Ok(listed, f))
                    {
                        Finish(run, "bg-list-order", f);
                    }
                    else
                    {
                        IList<Dictionary<string, object>> rows = Rows(listed);
                        if (rows.Count != deck.Count())
                        {
                            f.Add("长度应为 " + deck.Count() + " 实际 " + rows.Count);
                        }

                        int prev = 0;
                        for (int i = 0; i < rows.Count; i++)
                        {
                            int index = GetInt(rows[i], "index");
                            if (index != prev + 1)
                            {
                                f.Add("index 不连续 实际 " + index);
                                break;
                            }

                            prev = index;
                            string expect = deck.IdAt(index);
                            if (!Eq(Get(rows[i], "slide_id"), expect))
                            {
                                f.Add("index " + index + " slide_id 应为 " + expect);
                            }
                        }

                        Finish(run, "bg-list-order", f);
                    }
                }
            }

            if (Want(filters, "bg-get-master"))
            {
                int before = CountExport();
                var f = new List<string>();
                ToolResult r = Invoke(deck.Get(deck.IdM));
                if (!Ok(r, f))
                {
                    Finish(run, "bg-get-master", f);
                }
                else
                {
                    Dictionary<string, object> row = First(r);
                    if (row == null || !Eq(Get(row, "kind"), "master"))
                    {
                        f.Add("kind 应为 master");
                    }

                    if (row != null && (Has(row, "path") || Has(row, "color")))
                    {
                        f.Add("跟随母版不应带 path 或 color");
                    }

                    if (CountExport() != before)
                    {
                        f.Add("跟随母版不应新导出文件");
                    }

                    Finish(run, "bg-get-master", f);
                }
            }

            if (Want(filters, "bg-get-solid"))
            {
                var f = new List<string>();
                ToolResult r = Invoke(deck.Get(deck.IdS));
                if (!Ok(r, f))
                {
                    Finish(run, "bg-get-solid", f);
                }
                else
                {
                    Dictionary<string, object> row = First(r);
                    if (row == null || !Eq(Get(row, "kind"), "solid") || !Eq(Get(row, "color"), SolidHex))
                    {
                        f.Add("应为 solid " + SolidHex);
                    }

                    ExpectHex(deck, deck.IdS, SolidHex, f);
                    Finish(run, "bg-get-solid", f);
                }
            }

            if (Want(filters, "bg-get-theme"))
            {
                var f = new List<string>();
                ToolResult r = Invoke(deck.Get(deck.IdT));
                if (!Ok(r, f))
                {
                    Finish(run, "bg-get-theme", f);
                }
                else
                {
                    Dictionary<string, object> row = First(r);
                    string color = row == null ? "" : Get(row, "color");
                    if (row == null || !Eq(Get(row, "kind"), "solid") || !IsHex(color))
                    {
                        f.Add("主题色应读成 #RRGGBB 实际 " + color);
                    }

                    string blob = Newtonsoft.Json.JsonConvert.SerializeObject(r.Data);
                    if (blob.IndexOf("accent", StringComparison.OrdinalIgnoreCase) >= 0
                        || blob.IndexOf("scheme", StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        f.Add("回包带了主题槽名");
                    }

                    Finish(run, "bg-get-theme", f);
                }
            }

            if (Want(filters, "bg-get-picture"))
            {
                var f = new List<string>();
                ToolResult r = Invoke(deck.Get(deck.IdP));
                if (!Ok(r, f))
                {
                    Finish(run, "bg-get-picture", f);
                }
                else
                {
                    CheckStretch(First(r), 0, 0, 0, 0, f);
                    string path = Get(First(r), "path");
                    if (path.IndexOf("ppt_backgrounds/", StringComparison.OrdinalIgnoreCase) < 0)
                    {
                        f.Add("path 应在 ppt_backgrounds 下 实际 " + path);
                    }

                    ExpectFileHash(path, deck.RedHash, f);
                    ExpectPackageHash(deck, deck.IdP, deck.RedHash, f);
                    Finish(run, "bg-get-picture", f);
                }
            }

            if (Want(filters, "bg-get-other"))
            {
                var f = new List<string>();
                ToolResult r = Invoke(deck.Get(deck.IdO));
                if (!Ok(r, f))
                {
                    Finish(run, "bg-get-other", f);
                }
                else
                {
                    Dictionary<string, object> row = First(r);
                    if (row == null || !Eq(Get(row, "kind"), "other") || GetBool(row, "follow_master"))
                    {
                        f.Add("应为 other 且已脱离母版");
                    }

                    if (row != null && Has(row, "path"))
                    {
                        f.Add("other 不应带 path");
                    }

                    Finish(run, "bg-get-other", f);
                }
            }

            if (Want(filters, "bg-get-shape"))
            {
                var f = new List<string>();
                ToolResult r = Invoke(deck.Get(deck.IdShape));
                if (!Ok(r, f))
                {
                    Finish(run, "bg-get-shape", f);
                }
                else
                {
                    Dictionary<string, object> row = First(r);
                    if (row == null || !Eq(Get(row, "kind"), "master"))
                    {
                        f.Add("铺满形状不应当成背景 实际 " + (row == null ? "" : Get(row, "kind")));
                    }

                    if (row != null && Has(row, "path"))
                    {
                        f.Add("不应导出形状图");
                    }

                    Finish(run, "bg-get-shape", f);
                }
            }

            if (Want(filters, "bg-get-missing"))
            {
                ExpectFail(run, "bg-get-missing", deck.Get("999999"), "不存在");
            }
        }

        private static void RunSolid(TestRun run, IList<string> filters, Deck deck)
        {
            if (Want(filters, "bg-set-solid"))
            {
                deck.ForceMaster(1);
                var f = new List<string>();
                ToolResult r = Invoke(deck.SetSolid(deck.IdAt(1), SolidHex));
                if (!Ok(r, f))
                {
                    Finish(run, "bg-set-solid", f);
                }
                else
                {
                    ExpectHex(deck, deck.IdAt(1), SolidHex, f);
                    if (deck.Follows(1))
                    {
                        f.Add("应脱离母版");
                    }

                    ToolResult again = Invoke(deck.Get(deck.IdAt(1)));
                    if (!Ok(again, f) || !Eq(Get(First(again), "color"), SolidHex))
                    {
                        f.Add("再 get 色值不对");
                    }

                    Finish(run, "bg-set-solid", f);
                    deck.ForceMaster(1);
                }
            }

            if (Want(filters, "bg-set-solid-lower"))
            {
                var f = new List<string>();
                ToolResult r = Invoke(deck.SetSolid(deck.IdW, "#1e5aa8"));
                if (!Ok(r, f))
                {
                    Finish(run, "bg-set-solid-lower", f);
                }
                else if (!Eq(Get(First(r), "color"), SolidHex))
                {
                    f.Add("回包应为 " + SolidHex + " 实际 " + Get(First(r), "color"));
                    Finish(run, "bg-set-solid-lower", f);
                }
                else
                {
                    Finish(run, "bg-set-solid-lower", f);
                }
            }

            if (Want(filters, "bg-set-solid-over-pic"))
            {
                var f = new List<string>();
                ToolResult r = Invoke(deck.SetSolid(deck.IdP, "#00AA00"));
                if (!Ok(r, f))
                {
                    Finish(run, "bg-set-solid-over-pic", f);
                }
                else
                {
                    ToolResult again = Invoke(deck.Get(deck.IdP));
                    if (!Ok(again, f))
                    {
                        Finish(run, "bg-set-solid-over-pic", f);
                    }
                    else
                    {
                        Dictionary<string, object> row = First(again);
                        if (row == null || !Eq(Get(row, "kind"), "solid") || Has(row, "path"))
                        {
                            f.Add("盖掉图片后应为 solid 且无 path");
                        }

                        string hash = deck.PackageHash(deck.IndexOf(deck.IdP));
                        if (hash != null && Eq(hash, deck.RedHash))
                        {
                            f.Add("包里仍是红 PNG");
                        }

                        Finish(run, "bg-set-solid-over-pic", f);
                    }
                }

                deck.PaintPicture(3, deck.RedPath);
            }

            if (Want(filters, "bg-set-theme-fixed"))
            {
                var f = new List<string>();
                ToolResult got = Invoke(deck.Get(deck.IdT));
                string color = Ok(got, f) ? Get(First(got), "color") : "";
                if (!IsHex(color))
                {
                    f.Add("先 get 没有 #RRGGBB");
                    Finish(run, "bg-set-theme-fixed", f);
                }
                else
                {
                    ToolResult set = Invoke(deck.SetSolid(deck.IdT, color));
                    if (!Ok(set, f))
                    {
                        Finish(run, "bg-set-theme-fixed", f);
                    }
                    else if (deck.ColorType(deck.IndexOf(deck.IdT)) == (int)Office.MsoColorType.msoColorTypeScheme)
                    {
                        f.Add("再写回后仍是主题槽");
                        Finish(run, "bg-set-theme-fixed", f);
                    }
                    else
                    {
                        ExpectHex(deck, deck.IdT, color, f);
                        Finish(run, "bg-set-theme-fixed", f);
                    }
                }

                deck.PaintTheme(6);
            }

            if (Want(filters, "bg-set-other-page"))
            {
                deck.ForceMaster(1);
                deck.PaintSolid(2, SolidHex);
                string before = KindOf(Invoke(deck.Get(deck.IdAt(1))));
                var f = new List<string>();
                ToolResult r = Invoke(deck.SetSolid(deck.IdS, "#223344"));
                if (!Ok(r, f))
                {
                    Finish(run, "bg-set-other-page", f);
                }
                else
                {
                    string after = KindOf(Invoke(deck.Get(deck.IdAt(1))));
                    if (!Eq(after, before))
                    {
                        f.Add("未传的页被改了 前 " + before + " 后 " + after);
                    }

                    Finish(run, "bg-set-other-page", f);
                }

                deck.PaintSolid(2, SolidHex);
            }

            if (Want(filters, "bg-set-over-other"))
            {
                var f = new List<string>();
                ToolResult r = Invoke(deck.SetSolid(deck.IdO, "#445566"));
                if (!Ok(r, f))
                {
                    Finish(run, "bg-set-over-other", f);
                }
                else
                {
                    ToolResult again = Invoke(deck.Get(deck.IdO));
                    if (!Ok(again, f) || !Eq(Get(First(again), "kind"), "solid") || !Eq(Get(First(again), "color"), "#445566"))
                    {
                        f.Add("覆盖渐变后应为 #445566");
                    }

                    Finish(run, "bg-set-over-other", f);
                }

                deck.PaintGradient(4);
            }
        }

        private static void RunPicture(TestRun run, IList<string> filters, Deck deck)
        {
            if (Want(filters, "bg-set-pic-default"))
            {
                deck.ForceMaster(7);
                var f = new List<string>();
                ToolResult r = Invoke(deck.SetPic(deck.IdW, "bg-red.png"));
                if (!Ok(r, f))
                {
                    Finish(run, "bg-set-pic-default", f);
                }
                else
                {
                    if (!Eq(Get(First(r), "path"), "bg-red.png"))
                    {
                        f.Add("回包 path 应为传入相对路径 实际 " + Get(First(r), "path"));
                    }

                    CheckStretch(First(r), 0, 0, 0, 0, f);
                    if (deck.Follows(7))
                    {
                        f.Add("应脱离母版");
                    }

                    ExpectPackage(deck, deck.IdW, Stretch(0, 0, 0, 0), deck.RedHash, f);
                    Finish(run, "bg-set-pic-default", f);
                }
            }

            if (Want(filters, "bg-set-pic-abs"))
            {
                var f = new List<string>();
                ToolResult r = Invoke(deck.SetPic(deck.IdW, deck.BluePath));
                if (!Ok(r, f))
                {
                    Finish(run, "bg-set-pic-abs", f);
                }
                else
                {
                    ExpectPackageHash(deck, deck.IdW, deck.BlueHash, f);
                    Finish(run, "bg-set-pic-abs", f);
                }
            }

            if (Want(filters, "bg-set-stretch"))
            {
                var a = deck.SetPic(deck.IdW, "bg-red.png");
                a["fit"] = "stretch";
                a["stretch_left"] = 10d;
                a["stretch_right"] = 10d;
                a["stretch_top"] = 0d;
                a["stretch_bottom"] = 0d;
                var f = new List<string>();
                ToolResult r = Invoke(a);
                if (!Ok(r, f))
                {
                    Finish(run, "bg-set-stretch", f);
                }
                else
                {
                    CheckStretch(First(Invoke(deck.Get(deck.IdW))), 10, 10, 0, 0, f);
                    ExpectPackage(deck, deck.IdW, Stretch(10, 10, 0, 0), deck.RedHash, f);
                    Finish(run, "bg-set-stretch", f);
                }
            }

            if (Want(filters, "bg-set-stretch-neg"))
            {
                var a = deck.SetPic(deck.IdW, "bg-red.png");
                a["fit"] = "stretch";
                a["stretch_left"] = -5d;
                a["stretch_right"] = 0d;
                a["stretch_top"] = 0d;
                a["stretch_bottom"] = 0d;
                var f = new List<string>();
                ToolResult r = Invoke(a);
                if (!Ok(r, f))
                {
                    Finish(run, "bg-set-stretch-neg", f);
                }
                else
                {
                    CheckStretch(First(Invoke(deck.Get(deck.IdW))), -5, 0, 0, 0, f);
                    ExpectPackage(deck, deck.IdW, Stretch(-5, 0, 0, 0), null, f);
                    Finish(run, "bg-set-stretch-neg", f);
                }
            }

            if (Want(filters, "bg-set-stretch-omit"))
            {
                var a = deck.SetPic(deck.IdW, "bg-red.png");
                a["fit"] = "stretch";
                a["stretch_left"] = 8d;
                var f = new List<string>();
                ToolResult r = Invoke(a);
                if (!Ok(r, f))
                {
                    Finish(run, "bg-set-stretch-omit", f);
                }
                else
                {
                    CheckStretch(First(Invoke(deck.Get(deck.IdW))), 8, 0, 0, 0, f);
                    Finish(run, "bg-set-stretch-omit", f);
                }
            }

            if (Want(filters, "bg-copy-pic"))
            {
                deck.PaintPicture(3, deck.RedPath);
                var f = new List<string>();
                ToolResult src = Invoke(deck.Get(deck.IdP));
                if (!Ok(src, f))
                {
                    Finish(run, "bg-copy-pic", f);
                }
                else
                {
                    Dictionary<string, object> row = First(src);
                    var a = deck.SetPic(deck.IdW, Get(row, "path"));
                    a["fit"] = Get(row, "fit");
                    CopyNum(row, a, "stretch_left");
                    CopyNum(row, a, "stretch_right");
                    CopyNum(row, a, "stretch_top");
                    CopyNum(row, a, "stretch_bottom");
                    ToolResult set = Invoke(a);
                    if (!Ok(set, f))
                    {
                        Finish(run, "bg-copy-pic", f);
                    }
                    else
                    {
                        ExpectPackageHash(deck, deck.IdW, deck.RedHash, f);
                        ExpectPackage(deck, deck.IdW, Stretch(0, 0, 0, 0), null, f);
                        Finish(run, "bg-copy-pic", f);
                    }
                }
            }
        }

        private static void RunTile(TestRun run, IList<string> filters, Deck deck)
        {
            if (Want(filters, "bg-set-tile-default"))
            {
                var a = deck.SetPic(deck.IdW, "bg-red.png");
                a["fit"] = "tile";
                var f = new List<string>();
                ToolResult r = Invoke(a);
                if (!Ok(r, f))
                {
                    Finish(run, "bg-set-tile-default", f);
                }
                else
                {
                    CheckTile(First(Invoke(deck.Get(deck.IdW))), 100, 100, 0, 0, "tl", "none", f);
                    ToolResult listed = Invoke(deck.List());
                    Dictionary<string, object> row = Find(Rows(listed), deck.IdW);
                    if (row == null || !Eq(Get(row, "kind"), "picture"))
                    {
                        f.Add("list 应为 picture 不是 other");
                    }

                    ExpectPackage(deck, deck.IdW, Tile(100, 100, 0, 0, "tl", "none"), null, f);
                    Finish(run, "bg-set-tile-default", f);
                }
            }

            if (Want(filters, "bg-set-tile"))
            {
                var a = deck.SetPic(deck.IdW, "bg-red.png");
                a["fit"] = "tile";
                a["tile_scale_x"] = 40d;
                a["tile_scale_y"] = 40d;
                a["tile_align"] = "br";
                a["tile_offset_x"] = 12d;
                a["tile_offset_y"] = 6d;
                a["tile_mirror"] = "none";
                var f = new List<string>();
                ToolResult r = Invoke(a);
                if (!Ok(r, f))
                {
                    Finish(run, "bg-set-tile", f);
                }
                else
                {
                    CheckTile(First(Invoke(deck.Get(deck.IdW))), 40, 40, 12, 6, "br", "none", f);
                    ExpectPackage(deck, deck.IdW, Tile(40, 40, 12, 6, "br", "none"), null, f);
                    Finish(run, "bg-set-tile", f);
                }
            }

            if (Want(filters, "bg-set-tile-mirror"))
            {
                var a = deck.SetPic(deck.IdW, "bg-red.png");
                a["fit"] = "tile";
                a["tile_scale_x"] = 40d;
                a["tile_scale_y"] = 40d;
                a["tile_align"] = "br";
                a["tile_offset_x"] = 12d;
                a["tile_offset_y"] = 6d;
                a["tile_mirror"] = "xy";
                var f = new List<string>();
                ToolResult r = Invoke(a);
                if (!Ok(r, f))
                {
                    Finish(run, "bg-set-tile-mirror", f);
                }
                else
                {
                    ToolResult gotMirror = Invoke(deck.Get(deck.IdW));
                    if (!Ok(gotMirror, f))
                    {
                        f.Add(deck.FillDebug(deck.IndexOf(deck.IdW)));
                    }
                    else
                    {
                        CheckTile(First(gotMirror), 40, 40, 12, 6, "br", "xy", f);
                        if (f.Count > 0)
                        {
                            f.Add(deck.FillDebug(deck.IndexOf(deck.IdW)));
                        }
                    }

                    ExpectPackage(deck, deck.IdW, Tile(40, 40, 12, 6, "br", "xy"), null, f);
                    Finish(run, "bg-set-tile-mirror", f);
                }
            }
        }

        private static void RunTail(TestRun run, IList<string> filters, Deck deck)
        {
            if (Want(filters, "bg-clear") || Want(filters, "bg-clear-master-kept"))
            {
                deck.ForceMaster(1);
                deck.PaintSolid(2, SolidHex);
                ToolResult before = Invoke(deck.Get(deck.IdAt(1)));
                var f = new List<string>();
                ToolResult r = Invoke(deck.Clear(deck.IdS));
                if (Want(filters, "bg-clear"))
                {
                    if (!Ok(r, f))
                    {
                        Finish(run, "bg-clear", f);
                    }
                    else
                    {
                        Dictionary<string, object> row = First(r);
                        if (row == null || !GetBool(row, "follow_master") || !Eq(Get(row, "kind"), "master"))
                        {
                            f.Add("clear 后应为跟随母版");
                        }

                        if (!deck.Follows(2))
                        {
                            f.Add("COM 仍未跟随母版");
                        }

                        if (!deck.Follows(1))
                        {
                            f.Add("另一张跟随页被改了");
                        }

                        Finish(run, "bg-clear", f);
                    }
                }

                if (Want(filters, "bg-clear-master-kept"))
                {
                    var g = new List<string>();
                    ToolResult again = Invoke(deck.Get(deck.IdAt(1)));
                    if (!Ok(again, g) || !Eq(KindOf(again), "master") || !deck.Follows(1))
                    {
                        g.Add("从未改过的页不再跟随母版");
                    }

                    if (!Eq(KindOf(before), KindOf(again)))
                    {
                        g.Add("跟随页 kind 变了");
                    }

                    Finish(run, "bg-clear-master-kept", g);
                }

                deck.PaintSolid(2, SolidHex);
            }

            if (Want(filters, "bg-fail-keeps"))
            {
                deck.PaintSolid(2, SolidHex);
                var a = deck.SetPic(deck.IdS, "bg-red.png");
                a["color"] = "#112233";
                var f = new List<string>();
                ToolResult r = Invoke(a);
                if (r == null || r.Success)
                {
                    f.Add(r == null ? "无回包" : "应失败却成功");
                }

                ExpectHex(deck, deck.IdS, SolidHex, f);
                ToolResult again = Invoke(deck.Get(deck.IdS));
                if (!Ok(again, f) || !Eq(Get(First(again), "color"), SolidHex))
                {
                    f.Add("失败后 get 不再是原色");
                }

                Finish(run, "bg-fail-keeps", f);
            }

            if (Want(filters, "bg-missing-keeps"))
            {
                string before = Newtonsoft.Json.JsonConvert.SerializeObject(Rows(Invoke(deck.List())));
                var f = new List<string>();
                ToolResult r = Invoke(deck.SetSolid("999999", "#010101"));
                if (r == null || r.Success)
                {
                    f.Add(r == null ? "无回包" : "应失败却成功");
                }
                else if (!ContainsAny(r.Error, "不存在"))
                {
                    f.Add("文案应含不存在 实际 " + r.Error);
                }

                string after = Newtonsoft.Json.JsonConvert.SerializeObject(Rows(Invoke(deck.List())));
                if (!Eq(before, after))
                {
                    f.Add("失败后 list 变了");
                }

                Finish(run, "bg-missing-keeps", f);
            }

            if (Want(filters, "bg-action-upper"))
            {
                var a = deck.SetSolid(deck.IdW, SolidHex);
                a["action"] = "SET";
                var f = new List<string>();
                ToolResult r = Invoke(a);
                if (!Ok(r, f))
                {
                    Finish(run, "bg-action-upper", f);
                }
                else if (!string.Equals(GetData(r, "action"), "set", StringComparison.Ordinal))
                {
                    f.Add("回包 action 应为 set 实际 " + GetData(r, "action"));
                    Finish(run, "bg-action-upper", f);
                }
                else
                {
                    Finish(run, "bg-action-upper", f);
                }
            }

            if (Want(filters, "bg-echo-kind"))
            {
                var f = new List<string>();
                ToolResult r = Invoke(deck.List());
                if (!Ok(r, f))
                {
                    Finish(run, "bg-echo-kind", f);
                }
                else if (!Eq(GetData(r, "kind"), deck.Wpp ? "wpp" : "ppt"))
                {
                    f.Add("kind 应为 " + (deck.Wpp ? "wpp" : "ppt") + " 实际 " + GetData(r, "kind"));
                    Finish(run, "bg-echo-kind", f);
                }
                else
                {
                    Finish(run, "bg-echo-kind", f);
                }
            }

            if (Want(filters, "bg-echo-display"))
            {
                deck.PaintSolid(2, SolidHex);
                var f = new List<string>();
                ToolResult r = Invoke(deck.Get(deck.IdS));
                if (!Ok(r, f))
                {
                    Finish(run, "bg-echo-display", f);
                }
                else
                {
                    string text = GetData(r, "display_contents");
                    if (text.IndexOf(deck.IdS, StringComparison.Ordinal) < 0 || text.IndexOf("solid", StringComparison.OrdinalIgnoreCase) < 0)
                    {
                        f.Add("display_contents 应含 slide_id 与 solid");
                    }

                    Finish(run, "bg-echo-display", f);
                }
            }

            if (Want(filters, "bg-hidden"))
            {
                var f = new List<string>();
                if (!deck.TryHide(2, true, out string err))
                {
                    f.Add(err ?? "设隐藏失败");
                    Finish(run, "bg-hidden", f);
                }
                else
                {
                    ToolResult listed = Invoke(deck.List());
                    ToolResult got = Invoke(deck.Get(deck.IdS));
                    if (!Ok(listed, f) || Find(Rows(listed), deck.IdS) == null)
                    {
                        f.Add("隐藏页不在 list");
                    }

                    if (!Ok(got, f))
                    {
                        f.Add("隐藏页 get 失败");
                    }

                    if (!deck.IsHidden(2))
                    {
                        f.Add("工具把隐藏取消了");
                    }

                    deck.TryHide(2, false, out _);
                    Finish(run, "bg-hidden", f);
                }
            }

            if (Want(filters, "bg-text-kept"))
            {
                var f = new List<string>();
                if (!deck.TryText(7, "背景源", out string err))
                {
                    f.Add(err ?? "加文本失败");
                    Finish(run, "bg-text-kept", f);
                }
                else
                {
                    ToolResult r = Invoke(deck.SetSolid(deck.IdW, "#334455"));
                    if (!Ok(r, f))
                    {
                        Finish(run, "bg-text-kept", f);
                    }
                    else if (deck.SlideText(7).IndexOf("背景源", StringComparison.Ordinal) < 0)
                    {
                        f.Add("文本被改掉 实际 " + deck.SlideText(7));
                        Finish(run, "bg-text-kept", f);
                    }
                    else
                    {
                        Finish(run, "bg-text-kept", f);
                    }
                }
            }

            if (Want(filters, "bg-trn-kept"))
            {
                var f = new List<string>();
                var setTr = deck.Args();
                setTr["action"] = "set";
                setTr["slide_id"] = deck.IdW;
                setTr["effect"] = "fade";
                ToolResult tr = InvokeTr(setTr);
                if (!Ok(tr, f))
                {
                    Finish(run, "bg-trn-kept", f);
                }
                else
                {
                    ToolResult bg = Invoke(deck.SetSolid(deck.IdW, "#556677"));
                    if (!Ok(bg, f))
                    {
                        Finish(run, "bg-trn-kept", f);
                    }
                    else
                    {
                        ToolResult again = InvokeTr(deck.Get(deck.IdW));
                        string effect = "";
                        var d = again != null && again.Success ? again.Data as Dictionary<string, object> : null;
                        if (d != null && d.TryGetValue("transitions", out object raw) && raw is IEnumerable en)
                        {
                            foreach (object item in en)
                            {
                                var one = item as Dictionary<string, object>;
                                if (one != null)
                                {
                                    effect = Get(one, "effect");
                                    break;
                                }
                            }
                        }

                        if (!Eq(effect, "fade"))
                        {
                            f.Add("过渡不再是 fade 实际 " + effect + " " + (again == null ? "" : again.Error));
                        }

                        Finish(run, "bg-trn-kept", f);
                    }
                }
            }
        }

        private static void ExpectHex(Deck deck, string slideId, string hex, List<string> fails)
        {
            string actual = deck.ReadSolidHex(deck.IndexOf(slideId));
            if (!Eq(actual, hex))
            {
                fails.Add("COM 色值应为 " + hex + " 实际 " + actual);
            }
        }

        private static void ExpectFileHash(string relative, string expectHash, List<string> fails)
        {
            if (string.IsNullOrEmpty(relative))
            {
                fails.Add("没有 path");
                return;
            }

            string full = Path.Combine(
                WorkspacePathResolver.GetSessionDirectory(),
                relative.Replace('/', Path.DirectorySeparatorChar));
            if (!File.Exists(full))
            {
                fails.Add("导出文件不存在 " + relative);
                return;
            }

            string hash = Sha(full);
            if (!Eq(hash, expectHash))
            {
                fails.Add("导出哈希不等于源图");
            }
        }

        private static void ExpectPackageHash(Deck deck, string slideId, string expectHash, List<string> fails)
        {
            string hash = deck.PackageHash(deck.IndexOf(slideId));
            if (hash == null)
            {
                fails.Add("包里没有背景图");
            }
            else if (!Eq(hash, expectHash))
            {
                fails.Add("包内图片哈希不等于源图");
            }
        }

        private static void ExpectPackage(
            Deck deck,
            string slideId,
            PptBackgroundPackage.PicturePlacement expect,
            string hash,
            List<string> fails)
        {
            PptBackgroundPackage.PicturePlacement actual = deck.ReadPlacement(deck.IndexOf(slideId));
            if (actual == null)
            {
                fails.Add("包里读不到背景图摆放");
                return;
            }

            if (!PptBackgroundPackage.Matches(actual, expect, out string detail))
            {
                fails.Add(detail ?? "摆放不一致");
            }

            if (hash != null)
            {
                ExpectPackageHash(deck, slideId, hash, fails);
            }
        }

        private static void CheckKind(
            IList<Dictionary<string, object>> rows,
            string slideId,
            bool follow,
            string kind,
            string color,
            bool picture,
            List<string> fails,
            string label)
        {
            Dictionary<string, object> row = Find(rows, slideId);
            if (row == null)
            {
                fails.Add(label + " 不在 list");
                return;
            }

            if (GetBool(row, "follow_master") != follow || !Eq(Get(row, "kind"), kind))
            {
                fails.Add(label + " 应为 follow=" + follow + " " + kind
                    + " 实际 follow=" + GetBool(row, "follow_master") + " " + Get(row, "kind"));
            }

            if (color != null && !Eq(Get(row, "color"), color))
            {
                fails.Add(label + " 色值应为 " + color + " 实际 " + Get(row, "color"));
            }

            if (picture && (Has(row, "path") || Has(row, "fit")))
            {
                fails.Add(label + " list 不应带 path 或 fit");
            }
        }

        private static void CheckStretch(
            Dictionary<string, object> row,
            double left,
            double right,
            double top,
            double bottom,
            List<string> fails)
        {
            if (row == null || !Eq(Get(row, "fit"), "stretch"))
            {
                fails.Add("fit 应为 stretch");
                return;
            }

            if (!Near(GetNum(row, "stretch_left"), left, 0.1)
                || !Near(GetNum(row, "stretch_right"), right, 0.1)
                || !Near(GetNum(row, "stretch_top"), top, 0.1)
                || !Near(GetNum(row, "stretch_bottom"), bottom, 0.1))
            {
                fails.Add("拉伸偏移应为 "
                    + left + "," + right + "," + top + "," + bottom
                    + " 实际 "
                    + GetNum(row, "stretch_left") + ","
                    + GetNum(row, "stretch_right") + ","
                    + GetNum(row, "stretch_top") + ","
                    + GetNum(row, "stretch_bottom"));
            }
        }

        private static void CheckTile(
            Dictionary<string, object> row,
            double scaleX,
            double scaleY,
            double ox,
            double oy,
            string align,
            string mirror,
            List<string> fails)
        {
            if (row == null || !Eq(Get(row, "fit"), "tile"))
            {
                fails.Add("fit 应为 tile 实际 " + (row == null ? "(无行)" : Get(row, "fit") + " kind=" + Get(row, "kind")));
                return;
            }

            if (!Near(GetNum(row, "tile_scale_x"), scaleX, 0.5)
                || !Near(GetNum(row, "tile_scale_y"), scaleY, 0.5)
                || !Near(GetNum(row, "tile_offset_x"), ox, 0.5)
                || !Near(GetNum(row, "tile_offset_y"), oy, 0.5)
                || !Eq(Get(row, "tile_align"), align)
                || !Eq(Get(row, "tile_mirror"), mirror))
            {
                fails.Add("平铺应为 scale " + scaleX + " align " + align + " mirror " + mirror
                    + " 实际 scale " + GetNum(row, "tile_scale_x")
                    + " align " + Get(row, "tile_align")
                    + " mirror " + Get(row, "tile_mirror")
                    + " offset " + GetNum(row, "tile_offset_x") + "," + GetNum(row, "tile_offset_y"));
            }
        }

        private static PptBackgroundPackage.PicturePlacement Stretch(double left, double right, double top, double bottom)
        {
            return new PptBackgroundPackage.PicturePlacement
            {
                Fit = "stretch",
                StretchLeft = left,
                StretchRight = right,
                StretchTop = top,
                StretchBottom = bottom
            };
        }

        private static PptBackgroundPackage.PicturePlacement Tile(
            double scaleX,
            double scaleY,
            double ox,
            double oy,
            string align,
            string mirror)
        {
            return new PptBackgroundPackage.PicturePlacement
            {
                Fit = "tile",
                TileScaleX = scaleX,
                TileScaleY = scaleY,
                TileOffsetX = ox,
                TileOffsetY = oy,
                TileAlign = align,
                TileMirror = mirror
            };
        }

        private static void CopyNum(Dictionary<string, object> from, Dictionary<string, object> to, string key)
        {
            if (from != null && Has(from, key))
            {
                to[key] = GetNum(from, key);
            }
        }

        private static int CountExport()
        {
            string dir = Path.Combine(WorkspacePathResolver.GetSessionDirectory(), "ppt_backgrounds");
            if (!Directory.Exists(dir))
            {
                return 0;
            }

            return Directory.GetFiles(dir).Length;
        }

        private static string KindOf(ToolResult r)
        {
            return Get(First(r), "kind");
        }

        private static bool IsHex(string color)
        {
            if (string.IsNullOrEmpty(color) || color.Length != 7 || color[0] != '#')
            {
                return false;
            }

            for (int i = 1; i < 7; i++)
            {
                char c = color[i];
                bool ok = (c >= '0' && c <= '9') || (c >= 'A' && c <= 'F');
                if (!ok)
                {
                    return false;
                }
            }

            return true;
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
                if (string.Equals(f, "bg", StringComparison.OrdinalIgnoreCase)
                    || string.Equals(f, "bg-", StringComparison.OrdinalIgnoreCase))
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
            else if (contains != null && !ContainsAny(r.Error, contains))
            {
                fails.Add("文案应含「" + contains + "」实际 " + (r.Error ?? ""));
            }

            Finish(run, id, fails);
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
            F_PptBackgroundTool.Register(tools, null);
            if (!tools.TryGetValue("F_ppt_background", out Func<Dictionary<string, object>, Task<ToolResult>> fn))
            {
                return null;
            }

            return fn(args ?? new Dictionary<string, object>()).GetAwaiter().GetResult();
        }

        private static ToolResult InvokeTr(Dictionary<string, object> args)
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

        private static bool Eq(string a, string b)
        {
            return string.Equals(a ?? "", b ?? "", StringComparison.OrdinalIgnoreCase);
        }

        private static bool Near(double actual, double expect, double tol)
        {
            return Math.Abs(actual - expect) <= tol;
        }

        private static bool Has(Dictionary<string, object> row, string key)
        {
            if (row == null)
            {
                return false;
            }

            foreach (KeyValuePair<string, object> kv in row)
            {
                if (string.Equals(kv.Key, key, StringComparison.OrdinalIgnoreCase) && kv.Value != null)
                {
                    return true;
                }
            }

            return false;
        }

        private static string Get(Dictionary<string, object> row, string key)
        {
            if (row == null)
            {
                return "";
            }

            foreach (KeyValuePair<string, object> kv in row)
            {
                if (string.Equals(kv.Key, key, StringComparison.OrdinalIgnoreCase) && kv.Value != null)
                {
                    return Convert.ToString(kv.Value, CultureInfo.InvariantCulture) ?? "";
                }
            }

            return "";
        }

        private static string GetData(ToolResult r, string key)
        {
            return Get(r != null ? r.Data as Dictionary<string, object> : null, key);
        }

        private static bool GetBool(Dictionary<string, object> row, string key)
        {
            string text = Get(row, key);
            return string.Equals(text, "true", StringComparison.OrdinalIgnoreCase)
                || text == "-1"
                || text == "1";
        }

        private static int GetInt(Dictionary<string, object> row, string key)
        {
            string text = Get(row, key);
            if (int.TryParse(text, NumberStyles.Integer, CultureInfo.InvariantCulture, out int n))
            {
                return n;
            }

            return 0;
        }

        private static double GetNum(Dictionary<string, object> row, string key)
        {
            string text = Get(row, key);
            if (double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out double n))
            {
                return n;
            }

            return double.NaN;
        }

        private static IList<Dictionary<string, object>> Rows(ToolResult r)
        {
            var list = new List<Dictionary<string, object>>();
            var d = r != null ? r.Data as Dictionary<string, object> : null;
            if (d == null || !d.TryGetValue("backgrounds", out object raw) || !(raw is IEnumerable en))
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

        private static Dictionary<string, object> First(ToolResult r)
        {
            IList<Dictionary<string, object>> list = Rows(r);
            return list.Count > 0 ? list[0] : null;
        }

        private static Dictionary<string, object> Find(IList<Dictionary<string, object>> rows, string slideId)
        {
            if (rows == null)
            {
                return null;
            }

            for (int i = 0; i < rows.Count; i++)
            {
                if (Eq(Get(rows[i], "slide_id"), slideId))
                {
                    return rows[i];
                }
            }

            return null;
        }

        private static string Describe(Exception ex)
        {
            var sb = new System.Text.StringBuilder();
            Exception cur = ex;
            int guard = 0;
            while (cur != null && guard < 6)
            {
                if (sb.Length > 0)
                {
                    sb.Append(" <- ");
                }

                sb.Append(cur.GetType().Name);
                sb.Append(": ");
                sb.Append(cur.Message);
                cur = cur.InnerException;
                guard++;
            }

            return sb.ToString();
        }

        private static string Sha(string path)
        {
            using (SHA256 sha = SHA256.Create())
            using (FileStream fs = File.OpenRead(path))
            {
                return BitConverter.ToString(sha.ComputeHash(fs)).Replace("-", "");
            }
        }

        private static void SavePng(string path, Color color)
        {
            using (var bmp = new Bitmap(32, 32))
            using (Graphics g = Graphics.FromImage(bmp))
            {
                g.Clear(color);
                bmp.Save(path, ImageFormat.Png);
            }
        }

        private static object Prop(object target, string name)
        {
            return target.GetType().InvokeMember(
                name,
                BindingFlags.GetProperty | BindingFlags.Instance | BindingFlags.Public,
                null,
                target,
                null);
        }

        private static void SetProp(object target, string name, object value)
        {
            target.GetType().InvokeMember(
                name,
                BindingFlags.SetProperty | BindingFlags.Instance | BindingFlags.Public,
                null,
                target,
                new[] { value });
        }

        private static bool IsTrue(object value)
        {
            if (value == null)
            {
                return false;
            }

            if (value is bool b)
            {
                return b;
            }

            try
            {
                int n = Convert.ToInt32(value);
                return n != 0;
            }
            catch
            {
                return false;
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
            }

            public IOperationChannel Channel { get; }

            public object Pres { get; private set; }

            public bool Wpp { get; }

            public string RedPath { get; private set; }

            public string BluePath { get; private set; }

            public string RedHash { get; private set; }

            public string BlueHash { get; private set; }

            public string IdM
            {
                get { return IdAt(1); }
            }

            public string IdS
            {
                get { return IdAt(2); }
            }

            public string IdP
            {
                get { return IdAt(3); }
            }

            public string IdO
            {
                get { return IdAt(4); }
            }

            public string IdShape
            {
                get { return IdAt(5); }
            }

            public string IdT
            {
                get { return IdAt(6); }
            }

            public string IdW
            {
                get { return IdAt(7); }
            }

            public bool Prepare(out string error)
            {
                error = null;
                string step = "png";
                try
                {
                    string session = WorkspacePathResolver.GetSessionDirectory();
                    RedPath = Path.Combine(session, "bg-red.png");
                    BluePath = Path.Combine(Path.GetTempPath(), "bg-blue-" + Guid.NewGuid().ToString("N") + ".png");
                    SavePng(RedPath, Color.FromArgb(255, 220, 20, 20));
                    SavePng(BluePath, Color.FromArgb(255, 20, 60, 220));
                    RedHash = Sha(RedPath);
                    BlueHash = Sha(BluePath);
                    step = "slides";
                    if (!EnsureCount(7, out error))
                    {
                        return false;
                    }

                    step = "master";
                    ForceMaster(1);
                    step = "solid";
                    PaintSolid(2, SolidHex);
                    step = "picture";
                    PaintPicture(3, RedPath);
                    step = "gradient";
                    PaintGradient(4);
                    step = "shape-master";
                    ForceMaster(5);
                    step = "shape";
                    AddFullPicture(5, RedPath);
                    step = "theme";
                    PaintTheme(6);
                    step = "writer";
                    ForceMaster(7);
                    return true;
                }
                catch (Exception ex)
                {
                    error = step + ": " + Describe(ex);
                    return false;
                }
            }

            public Dictionary<string, object> Args()
            {
                return new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase)
                {
                    ["channel_id"] = Channel.ChannelId
                };
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

            public Dictionary<string, object> SetSolid(string slideId, string color)
            {
                var a = Args();
                a["action"] = "set";
                a["slide_id"] = slideId;
                a["kind"] = "solid";
                a["color"] = color;
                return a;
            }

            public Dictionary<string, object> SetPic(string slideId, string path)
            {
                var a = Args();
                a["action"] = "set";
                a["slide_id"] = slideId;
                a["kind"] = "picture";
                a["path"] = path;
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
                object slides = Prop(Pres, "Slides");
                return Convert.ToInt32(Prop(slides, "Count") ?? 0);
            }

            public string IdAt(int index1)
            {
                Refresh();
                object slide = Slide(index1);
                return Convert.ToString(Prop(slide, "SlideID"), CultureInfo.InvariantCulture) ?? "";
            }

            public int IndexOf(string slideId)
            {
                int n = Count();
                for (int i = 1; i <= n; i++)
                {
                    if (string.Equals(IdAt(i), slideId, StringComparison.Ordinal))
                    {
                        return i;
                    }
                }

                return 0;
            }

            public bool Follows(int index1)
            {
                return IsTrue(Prop(Slide(index1), "FollowMasterBackground"));
            }

            public string ReadSolidHex(int index1)
            {
                object fill = Fill(Slide(index1));
                int v = Convert.ToInt32(Prop(Prop(fill, "ForeColor"), "RGB"));
                int r = v & 0xFF;
                int g = (v >> 8) & 0xFF;
                int b = (v >> 16) & 0xFF;
                return "#" + r.ToString("X2") + g.ToString("X2") + b.ToString("X2");
            }

            public string FillDebug(int index1)
            {
                try
                {
                    object fill = Fill(Slide(index1));
                    string type = Convert.ToString(Prop(fill, "Type"), CultureInfo.InvariantCulture);
                    string texture = "";
                    try
                    {
                        texture = Convert.ToString(Prop(fill, "TextureType"), CultureInfo.InvariantCulture);
                    }
                    catch (Exception ex)
                    {
                        texture = ex.GetType().Name;
                    }

                    return "fillType=" + type + " textureType=" + texture + " follow=" + Follows(index1);
                }
                catch (Exception ex)
                {
                    return Describe(ex);
                }
            }

            public int ColorType(int index1)
            {
                object fill = Fill(Slide(index1));
                return Convert.ToInt32(Prop(Prop(fill, "ForeColor"), "Type"));
            }

            public void ForceMaster(int index1)
            {
                SetProp(Slide(index1), "FollowMasterBackground", -1);
            }

            public void PaintSolid(int index1, string hex)
            {
                object slide = Slide(index1);
                SetProp(slide, "FollowMasterBackground", 0);
                object fill = Fill(slide);
                WppCom.Invoke(fill, "Solid");
                SetProp(Prop(fill, "ForeColor"), "RGB", HexToBgr(hex));
            }

            public void PaintPicture(int index1, string path)
            {
                object slide = Slide(index1);
                SetProp(slide, "FollowMasterBackground", 0);
                object fill = Fill(slide);
                WppCom.Invoke(fill, "UserPicture", path);
                SetProp(fill, "TextureTile", 0);
            }

            public void PaintGradient(int index1)
            {
                object slide = Slide(index1);
                SetProp(slide, "FollowMasterBackground", 0);
                object fill = Fill(slide);
                WppCom.Invoke(fill, "Patterned", 1);
            }

            public void PaintTheme(int index1)
            {
                object slide = Slide(index1);
                SetProp(slide, "FollowMasterBackground", 0);
                object fill = Fill(slide);
                WppCom.Invoke(fill, "Solid");
                SetProp(
                    Prop(fill, "ForeColor"),
                    "ObjectThemeColor",
                    (int)Office.MsoThemeColorIndex.msoThemeColorAccent1);
            }

            public void AddFullPicture(int index1, string path)
            {
                object slide = Slide(index1);
                object setup = Prop(Pres, "PageSetup");
                float w = Convert.ToSingle(Prop(setup, "SlideWidth"));
                float h = Convert.ToSingle(Prop(setup, "SlideHeight"));
                string picPath = Path.Combine(Path.GetTempPath(), "bg-shape-" + Guid.NewGuid().ToString("N") + ".png");
                File.Copy(path, picPath, true);
                if (!Wpp)
                {
                    try
                    {
                        PowerPoint.Presentation ppt = (PowerPoint.Presentation)Pres;
                        ppt.Slides[index1].Shapes.AddPicture(
                            picPath,
                            Office.MsoTriState.msoFalse,
                            Office.MsoTriState.msoTrue,
                            0f,
                            0f,
                            w,
                            h);
                    }
                    catch (Exception ex)
                    {
                        throw new InvalidOperationException(
                            "path=" + picPath + " exists=" + File.Exists(picPath) + " w=" + w + " h=" + h + " " + ex.Message,
                            ex);
                    }

                    return;
                }

                object shapes = Prop(slide, "Shapes");
                WppCom.Invoke(shapes, "AddPicture", path, false, true, 0f, 0f, w, h);
            }

            public bool TryHide(int index1, bool hidden, out string error)
            {
                error = null;
                try
                {
                    object tr = Prop(Slide(index1), "SlideShowTransition");
                    SetProp(tr, "Hidden", hidden ? -1 : 0);
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
                try
                {
                    object tr = Prop(Slide(index1), "SlideShowTransition");
                    return IsTrue(Prop(tr, "Hidden"));
                }
                catch
                {
                    return false;
                }
            }

            public bool TryText(int index1, string text, out string error)
            {
                error = null;
                try
                {
                    object shapes = Prop(Slide(index1), "Shapes");
                    object box = WppCom.Invoke(shapes, "AddTextbox", 1, 30f, 30f, 200f, 40f);
                    object range = Prop(Prop(box, "TextFrame"), "TextRange");
                    SetProp(range, "Text", text);
                    return true;
                }
                catch (Exception ex)
                {
                    error = ex.Message;
                    return false;
                }
            }

            public string SlideText(int index1)
            {
                try
                {
                    object shapes = Prop(Slide(index1), "Shapes");
                    int n = Convert.ToInt32(Prop(shapes, "Count") ?? 0);
                    var sb = new System.Text.StringBuilder();
                    for (int i = 1; i <= n; i++)
                    {
                        object shape = WppCom.GetIndexed(shapes, i);
                        try
                        {
                            object range = Prop(Prop(shape, "TextFrame"), "TextRange");
                            sb.Append(Convert.ToString(Prop(range, "Text")));
                        }
                        catch
                        {
                        }
                    }

                    return sb.ToString();
                }
                catch
                {
                    return "";
                }
            }

            public PptBackgroundPackage.PicturePlacement ReadPlacement(int index1)
            {
                string path = PptHtmlOoxmlIo.TrySaveCopyTemp(Pres, "bg-oracle", out _);
                if (string.IsNullOrEmpty(path))
                {
                    return null;
                }

                try
                {
                    if (!PptBackgroundPackage.TryReadPicture(path, index1, null, out PptBackgroundPackage.PicturePlacement placement, out _))
                    {
                        return null;
                    }

                    return placement;
                }
                finally
                {
                    PptHtmlOoxmlIo.TryDeleteFile(path);
                }
            }

            public string PackageHash(int index1)
            {
                string path = PptHtmlOoxmlIo.TrySaveCopyTemp(Pres, "bg-hash", out _);
                if (string.IsNullOrEmpty(path))
                {
                    return null;
                }

                string export = Path.Combine(Path.GetTempPath(), "bg-export-" + Guid.NewGuid().ToString("N"));
                try
                {
                    if (!PptBackgroundPackage.TryReadPicture(
                            path,
                            index1,
                            export,
                            out PptBackgroundPackage.PicturePlacement placement,
                            out _))
                    {
                        return null;
                    }

                    if (placement == null || string.IsNullOrEmpty(placement.RelativePath))
                    {
                        return null;
                    }

                    string name = Path.GetFileName(placement.RelativePath.Replace('/', Path.DirectorySeparatorChar));
                    string full = Path.Combine(export, name);
                    if (!File.Exists(full))
                    {
                        return null;
                    }

                    return Sha(full);
                }
                finally
                {
                    PptHtmlOoxmlIo.TryDeleteFile(path);
                    try
                    {
                        if (Directory.Exists(export))
                        {
                            Directory.Delete(export, true);
                        }
                    }
                    catch
                    {
                    }
                }
            }

            private bool EnsureCount(int n, out string error)
            {
                error = null;
                Refresh();
                if (Count() == 0)
                {
                    TryAddBlank(out _);
                }

                while (Count() < n)
                {
                    if (!TryAddBlank(out error))
                    {
                        return false;
                    }
                }

                return Count() >= n;
            }

            private bool TryAddBlank(out string error)
            {
                error = null;
                Refresh();
                try
                {
                    object slides = Prop(Pres, "Slides");
                    int n = Convert.ToInt32(Prop(slides, "Count") ?? 0);
                    try
                    {
                        WppCom.Invoke(slides, "Add", n + 1, 12);
                    }
                    catch
                    {
                        WppCom.Invoke(slides, "Add", n + 1, (int)PowerPoint.PpSlideLayout.ppLayoutBlank);
                    }

                    return true;
                }
                catch (Exception ex)
                {
                    error = ex.Message;
                    return false;
                }
            }

            private object Slide(int index1)
            {
                Refresh();
                object slides = Prop(Pres, "Slides");
                return WppCom.GetIndexed(slides, index1);
            }

            private static object Fill(object slide)
            {
                return Prop(Prop(slide, "Background"), "Fill");
            }

            private static int HexToBgr(string hex)
            {
                string s = (hex ?? "").Trim();
                if (s.StartsWith("#", StringComparison.Ordinal))
                {
                    s = s.Substring(1);
                }

                int rgb = int.Parse(s, NumberStyles.HexNumber, CultureInfo.InvariantCulture);
                int r = (rgb >> 16) & 0xFF;
                int g = (rgb >> 8) & 0xFF;
                int b = rgb & 0xFF;
                return r | (g << 8) | (b << 16);
            }
        }
    }
}
