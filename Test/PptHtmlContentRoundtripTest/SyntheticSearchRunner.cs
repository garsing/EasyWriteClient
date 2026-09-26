using System;
using System.Collections.Generic;
using System.Globalization;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using WordAddIn1;
using WordAddIn1.OpenFiles;
using WordAddIn1.PresentationHost;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PptHtmlContentRoundtripTest
{
    /// <summary>空白页 apply 再搜。前缀 srch-（不含 srch-fx-）。</summary>
    internal static class SyntheticSearchRunner
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
                if (f.IndexOf("srch-fx", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    continue;
                }

                if (f.IndexOf("srch-", StringComparison.OrdinalIgnoreCase) >= 0
                    || string.Equals(f, "srch", StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
            }

            return false;
        }

        public static void Run(TestRun run, IList<string> nameFilters, bool useWpp)
        {
            Console.WriteLine(useWpp
                ? "合成页搜索 · WPP · apply 再 query / fields"
                : "合成页搜索 · PPT · apply 再 query / fields");
            RunParseBlock(run, nameFilters);
            RunToolMutex(run, nameFilters);
            if (NeedCom(nameFilters))
            {
                RunCom(run, nameFilters, useWpp);
            }

            Console.WriteLine("合成页 通过 " + run.CasesPassed
                + "  失败 " + run.CasesFailed
                + "  跳过 " + run.CasesSkipped);
        }

        private static bool NeedCom(IList<string> filters)
        {
            string[] com =
            {
                "srch-text-", "srch-type-", "srch-and-", "srch-any-", "srch-short-",
                "srch-shape-id", "srch-leaf-", "srch-group-", "srch-parent-",
                "srch-geo-", "srch-zero-", "srch-chart-", "srch-fields-skeleton",
                "srch-fields-fill", "srch-fields-no-extra", "srch-fields-group",
                "srch-pic-"
            };
            if (IsBroad(filters))
            {
                return true;
            }

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
                if (string.Equals(f, "srch", StringComparison.OrdinalIgnoreCase)
                    || string.Equals(f, "srch-", StringComparison.OrdinalIgnoreCase))
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
                if (f.Length == 0)
                {
                    continue;
                }

                if (id.IndexOf(f, StringComparison.OrdinalIgnoreCase) >= 0
                    || f.IndexOf(id, StringComparison.OrdinalIgnoreCase) >= 0)
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

        private static void RunParseBlock(TestRun run, IList<string> filters)
        {
            Console.WriteLine("--- A 解析（不启 COM）---");
            if (Want(filters, "srch-parse-query-ok"))
            {
                var fails = new List<string>();
                AssertParseOk(fails, 1, null, expectAny: false);
                AssertParseOk(fails, 8, "and", expectAny: false);
                Finish(run, "srch-parse-query-ok", fails);
            }

            if (Want(filters, "srch-parse-empty-fail"))
            {
                var fails = new List<string>();
                AssertParseFail(fails, ArgsQuery(0), "query 须 1");
                Finish(run, "srch-parse-empty-fail", fails);
            }

            if (Want(filters, "srch-parse-over8-fail"))
            {
                var fails = new List<string>();
                AssertParseFail(fails, ArgsQuery(9), "query 须 1");
                Finish(run, "srch-parse-over8-fail", fails);
            }

            if (Want(filters, "srch-parse-bad-regex-fail"))
            {
                var fails = new List<string>();
                var args = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase)
                {
                    ["query"] = new List<object>
                    {
                        new Dictionary<string, object>
                        {
                            ["attr"] = "text",
                            ["pattern"] = "("
                        }
                    }
                };
                AssertParseFail(fails, args, "attr=");
                Finish(run, "srch-parse-bad-regex-fail", fails);
            }

            if (Want(filters, "srch-parse-short-and-query-fail"))
            {
                var fails = new List<string>();
                var args = ArgsQuery(1);
                args["attr"] = "text";
                args["pattern"] = "电";
                AssertParseFail(fails, args, "不能同时传");
                Finish(run, "srch-parse-short-and-query-fail", fails);
            }

            if (Want(filters, "srch-parse-fields-all-fail"))
            {
                var fails = new List<string>();
                var args = ArgsQuery(1);
                args["fields"] = new List<object> { "all" };
                AssertParseFail(fails, args, "all/detail");
                var args2 = ArgsQuery(1);
                args2["fields"] = new List<object> { "detail" };
                AssertParseFail(fails, args2, "all/detail");
                Finish(run, "srch-parse-fields-all-fail", fails);
            }

            if (Want(filters, "srch-parse-fields-unknown-fail"))
            {
                var fails = new List<string>();
                var args = ArgsQuery(1);
                args["fields"] = new List<object> { "not-a-field" };
                AssertParseFail(fails, args, "未知名");
                Finish(run, "srch-parse-fields-unknown-fail", fails);
            }

            if (Want(filters, "srch-parse-picture-flag"))
            {
                var fails = new List<string>();
                AssertPictureFlag(fails, "picture", true);
                AssertPictureFlag(fails, "^picture$", true);
                AssertPictureFlag(fails, "chart", false);
                AssertPictureFlag(fails, "text", false);
                Finish(run, "srch-parse-picture-flag", fails);
            }

            if (Want(filters, "srch-fields-dup-fail"))
            {
                var fails = new List<string>();
                var args = ArgsQuery(1);
                args["fields"] = new List<object> { "data-fill", "data-fill" };
                AssertParseFail(fails, args, "不要重复");
                Finish(run, "srch-fields-dup-fail", fails);
            }
        }

        private static void RunToolMutex(TestRun run, IList<string> filters)
        {
            Console.WriteLine("--- E 互斥（工具层）---");
            if (Want(filters, "srch-neg-with-full"))
            {
                var fails = new List<string>();
                var args = ArgsQuery(1);
                args["full"] = true;
                AssertToolFail(fails, args, "full");
                Finish(run, "srch-neg-with-full", fails);
            }

            if (Want(filters, "srch-neg-with-shape-id"))
            {
                var fails = new List<string>();
                var args = ArgsQuery(1);
                args["shape_id"] = "sid1-s2";
                AssertToolFail(fails, args, "shape_id");
                Finish(run, "srch-neg-with-shape-id", fails);
            }

            if (Want(filters, "srch-neg-with-path"))
            {
                var fails = new List<string>();
                var args = ArgsQuery(1);
                args["path"] = "out.html";
                AssertToolFail(fails, args, "path");
                Finish(run, "srch-neg-with-path", fails);
            }

            if (Want(filters, "srch-neg-fields-alone"))
            {
                var fails = new List<string>();
                var args = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase)
                {
                    ["fields"] = new List<object> { "data-fill" }
                };
                if (PptHtmlNodeSearch.IsSearchArgs(args))
                {
                    fails.Add("只传 fields 不应被当成搜索");
                }

                if (!PptHtmlNodeSearch.HasFieldsKey(args))
                {
                    fails.Add("应识别 fields");
                }

                Finish(run, "srch-neg-fields-alone", fails);
            }
        }

        private static void RunCom(TestRun run, IList<string> filters, bool useWpp)
        {
            Console.WriteLine(useWpp ? "--- B–G COM（WPP）---" : "--- B–G COM ---");
            string assetsDir = System.IO.Path.Combine(
                AppDomain.CurrentDomain.BaseDirectory, "srch-synth-assets");
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
                var session = new Session(pres, assets, wpp: false);
                RunPages(run, filters, session);
            }
            catch (Exception ex)
            {
                run.CaseFail("srch-", "合成页 COM 失败: " + ex.Message);
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
                    run.CaseFail("srch-", startErr);
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
                    run.CaseFail("srch-", "WPP Presentations.Add 失败");
                    return;
                }

                var session = new Session(pres, assets, wpp: true);
                RunPages(run, filters, session);
            }
            catch (Exception ex)
            {
                run.CaseFail("srch-", "合成页 COM 失败: " + ex.Message);
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

        private static void RunPages(TestRun run, IList<string> filters, Session session)
        {
            RunPageB(run, filters, session);
            RunPageC(run, filters, session);
            RunPageD(run, filters, session);
            RunPageF(run, filters, session);
            RunPageG(run, filters, session);
        }

        private static void RunPageB(TestRun run, IList<string> filters, Session s)
        {
            bool need = Want(filters, "srch-text-one")
                || Want(filters, "srch-type-picture")
                || Want(filters, "srch-type-textbox")
                || Want(filters, "srch-and-text-type")
                || Want(filters, "srch-any-text-or-type")
                || Want(filters, "srch-short-attr")
                || Want(filters, "srch-shape-id")
                || Want(filters, "srch-zero-hit");
            if (!need)
            {
                return;
            }

            if (!s.TryNewSlide(out string err))
            {
                FailNeeded(run, filters, err, "srch-text-one", "srch-type-picture",
                    "srch-type-textbox", "srch-and-text-type", "srch-any-text-or-type",
                    "srch-short-attr", "srch-shape-id", "srch-zero-hit");
                return;
            }

            string html = ContentHtml.Section(
                "  <div data-shape-type=\"textbox\" style=\"" + ContentHtml.Geo(8, 12, 28, 14) + "\">电力</div>\n"
                + "  <div data-shape-type=\"textbox\" style=\"" + ContentHtml.Geo(8, 40, 28, 14) + "\">其它</div>\n"
                + ContentHtml.MixPicMarkup(s.Assets.RedPng, 50, 20, 40, 30));
            if (!s.TryApply(html, out _, out err) || !s.TryReadSearch(out PptHtmlReadResult tree, out err))
            {
                FailNeeded(run, filters, err, "srch-text-one", "srch-type-picture",
                    "srch-type-textbox", "srch-and-text-type", "srch-any-text-or-type",
                    "srch-short-attr", "srch-shape-id", "srch-zero-hit");
                return;
            }

            string idDian = FindIdByText(tree, "电力");
            string idQi = FindIdByText(tree, "其它");
            string idPic = FindIdByType(tree, "picture");
            if (string.IsNullOrEmpty(idDian) || string.IsNullOrEmpty(idQi) || string.IsNullOrEmpty(idPic))
            {
                FailNeeded(run, filters, "铺页未找到 电力/其它/图", "srch-text-one", "srch-type-picture",
                    "srch-type-textbox", "srch-and-text-type", "srch-any-text-or-type",
                    "srch-short-attr", "srch-shape-id", "srch-zero-hit");
                return;
            }

            if (Want(filters, "srch-text-one"))
            {
                SearchAssert(
                    run,
                    s,
                    "srch-text-one",
                    Query("text", "电"),
                    1,
                    new[] { idDian },
                    (display, _) => display.IndexOf("电力", StringComparison.Ordinal) >= 0
                        ? null
                        : "正文应含「电力」");
            }

            if (Want(filters, "srch-type-picture"))
            {
                SearchAssert(run, s, "srch-type-picture", Query("data-shape-type", "picture"), 1, new[] { idPic }, null);
            }

            if (Want(filters, "srch-type-textbox"))
            {
                SearchAssert(
                    run,
                    s,
                    "srch-type-textbox",
                    Query("data-shape-type", "textbox"),
                    2,
                    new[] { idDian, idQi },
                    null);
            }

            if (Want(filters, "srch-and-text-type"))
            {
                var args = Query("text", "电");
                AddClause(args, "data-shape-type", "textbox");
                args["match"] = "and";
                SearchAssert(run, s, "srch-and-text-type", args, 1, new[] { idDian }, null);
            }

            if (Want(filters, "srch-any-text-or-type"))
            {
                var args = Query("text", "电");
                AddClause(args, "data-shape-type", "picture");
                args["match"] = "or";
                SearchAssert(run, s, "srch-any-text-or-type", args, 2, new[] { idDian, idPic }, null);
            }

            if (Want(filters, "srch-short-attr"))
            {
                var args = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase)
                {
                    ["attr"] = "text",
                    ["pattern"] = "其它"
                };
                SearchAssert(run, s, "srch-short-attr", args, 1, new[] { idQi }, null);
            }

            if (Want(filters, "srch-shape-id"))
            {
                SearchAssert(
                    run,
                    s,
                    "srch-shape-id",
                    Query("shape_id", Regex.Escape(idPic)),
                    1,
                    new[] { idPic },
                    null);
            }

            if (Want(filters, "srch-zero-hit"))
            {
                SearchAssert(run, s, "srch-zero-hit", Query("text", "zzz"), 0, new string[0], (display, _) =>
                    HasTag(display) ? "0 命中不应有标签节点" : null);
            }
        }

        private static void RunPageC(TestRun run, IList<string> filters, Session s)
        {
            bool needGroup = Want(filters, "srch-leaf-not-group")
                || Want(filters, "srch-group-shell-merge")
                || Want(filters, "srch-geo-page-pct")
                || Want(filters, "srch-fields-group-shell");
            if (needGroup)
            {
                if (!s.TryNewSlide(out string err)
                    || !s.TryApply(
                        ContentHtml.TwoTextboxes("甲", "乙", 10, 12, 40, 40),
                        out List<string> ids,
                        out err))
                {
                    FailNeeded(run, filters, err, "srch-leaf-not-group", "srch-group-shell-merge",
                        "srch-geo-page-pct", "srch-fields-group-shell");
                }
                else
                {
                    var groupSess = new ContentGroupSession
                    {
                        Wpp = s.Wpp,
                        Presentation = s.Pres,
                        Slide = s.Slide,
                        SlideId = s.SlideId
                    };
                    if (!groupSess.TryGroup(ids.ToArray(), out _, out err)
                        || !s.TryReadSearch(out PptHtmlReadResult tree, out err))
                    {
                        FailNeeded(run, filters, err, "srch-leaf-not-group", "srch-group-shell-merge",
                            "srch-geo-page-pct", "srch-fields-group-shell");
                    }
                    else
                    {
                        string idJia = FindIdByText(tree, "甲");
                        string idYi = FindIdByText(tree, "乙");
                        string idGroup = FindParentId(tree, idJia);
                        if (Want(filters, "srch-leaf-not-group"))
                        {
                            SearchAssert(
                                run,
                                s,
                                "srch-leaf-not-group",
                                Query("text", ".+"),
                                2,
                                new[] { idJia, idYi },
                                (display, hit) =>
                                    LeafIds(hit.Forest).Contains(idGroup) ? "组不应算命中" : null);
                        }

                        if (Want(filters, "srch-group-shell-merge"))
                        {
                            var fails = new List<string>();
                            if (!s.TrySearch(Query("text", ".+"), out SearchHit hit, out string se))
                            {
                                fails.Add(se);
                            }
                            else
                            {
                                CheckCommon(fails, hit, 2);
                                PptHtmlShapeNode shell = FindNode(hit.Forest, idGroup);
                                if (shell == null || !IsGroup(shell))
                                {
                                    fails.Add("应有一个 group 壳 " + idGroup);
                                }
                                else if (shell.Children == null || shell.Children.Count != 2)
                                {
                                    fails.Add("壳下应有两个孩子");
                                }
                            }

                            Finish(run, "srch-group-shell-merge", fails);
                        }

                        if (Want(filters, "srch-geo-page-pct"))
                        {
                            var fails = new List<string>();
                            if (!s.TrySearch(Query("text", "甲"), out SearchHit hit, out string se))
                            {
                                fails.Add(se);
                            }
                            else
                            {
                                CheckCommon(fails, hit, 1);
                                PptHtmlShapeNode shell = FindNode(hit.Forest, idGroup);
                                PptHtmlShapeNode leaf = FindNode(hit.Forest, idJia);
                                if (shell == null || leaf == null)
                                {
                                    fails.Add("缺壳或叶子");
                                }
                                else if (!LooksPagePct(shell.Style) || !LooksPagePct(leaf.Style))
                                {
                                    fails.Add("left/top 应是页 %");
                                }
                                else if (!ContentGroupSession.TopNearGroup(leaf, shell))
                                {
                                    fails.Add("叶子 top 应靠近组 top，不是 0%");
                                }
                            }

                            Finish(run, "srch-geo-page-pct", fails);
                        }

                        if (Want(filters, "srch-fields-group-shell"))
                        {
                            var args = Query("text", ".+");
                            args["fields"] = new List<object> { "data-fill" };
                            var fails = new List<string>();
                            if (!s.TrySearch(args, out SearchHit hit, out string se))
                            {
                                fails.Add(se);
                            }
                            else
                            {
                                CheckCommon(fails, hit, 2);
                                PptHtmlShapeNode shell = FindNode(hit.Forest, idGroup);
                                if (shell != null && !string.IsNullOrEmpty(shell.Fill))
                                {
                                    fails.Add("组壳不应套叶子 fields");
                                }
                            }

                            Finish(run, "srch-fields-group-shell", fails);
                        }
                    }
                }
            }

            if (!Want(filters, "srch-parent-one-level"))
            {
                return;
            }

            if (!s.TryNewSlide(out string nerr)
                || !s.TryApply(ContentHtml.NestedGroups(), out _, out nerr)
                || !s.TryReadSearch(out PptHtmlReadResult nest, out nerr))
            {
                run.CaseFail("srch-parent-one-level", nerr ?? "铺嵌套组失败");
                return;
            }

            string idJia2 = FindIdByText(nest, "内组甲");
            string idInner = FindParentId(nest, idJia2);
            string idOuter = FindParentId(nest, idInner);
            var pfails = new List<string>();
            if (!s.TrySearch(Query("text", ".+"), out SearchHit phit, out string perr))
            {
                pfails.Add(perr);
            }
            else
            {
                CheckCommon(pfails, phit, 3);
                PptHtmlShapeNode inner = FindNode(phit.Forest, idInner);
                if (inner == null)
                {
                    pfails.Add("应带一层直接父组");
                }

                if (!string.IsNullOrEmpty(idOuter) && FindNode(phit.Forest, idOuter) != null
                    && FindChild(FindNode(phit.Forest, idOuter), idInner) != null)
                {
                    pfails.Add("祖父不应包住内组");
                }
            }

            Finish(run, "srch-parent-one-level", pfails);
        }

        private static void RunPageD(TestRun run, IList<string> filters, Session s)
        {
            if (!Want(filters, "srch-chart-title-only"))
            {
                return;
            }

            string pie = ContentHtml.ChartFrag(
                "pie2d",
                8,
                12,
                40,
                40,
                new[] { "甲类目", "乙类目", "丙类目" },
                new[] { "40", "35", "25" },
                "份额");
            pie = pie.Replace("data-chart-type=\"pie2d\"", "data-chart-type=\"pie2d\" data-title=\"电力占比\"");
            string page = ContentHtml.Section(
                pie + "\n"
                + ContentHtml.MixPageTableMarkup(new[] { new[] { "甲类目", "1" }, new[] { "其它格", "2" } }, 52, 12, 40, 40));
            if (!s.TryNewSlide(out string err)
                || !s.TryApply(page, out _, out err)
                || !s.TryReadSearch(out PptHtmlReadResult tree, out err))
            {
                run.CaseFail("srch-chart-title-only", err ?? "铺饼+表失败");
                return;
            }

            string idChart = FindIdByType(tree, "chart");
            string idTable = FindIdByType(tree, "table");
            var fails = new List<string>();
            if (!s.TrySearch(Query("text", "电力占比"), out SearchHit hit, out string se))
            {
                fails.Add(se);
            }
            else
            {
                CheckCommon(fails, hit, 1);
                if (!HitIds(hit.Display).Contains(idChart))
                {
                    fails.Add("应中饼 " + idChart + " 实际 " + string.Join(",", HitIds(hit.Display)));
                }
            }

            if (!s.TrySearch(Query("text", "甲类目"), out SearchHit tableHit, out string te))
            {
                fails.Add(te);
            }
            else if (!HitIds(tableHit.Display).Contains(idTable) || HitIds(tableHit.Display).Contains(idChart))
            {
                fails.Add("甲类目应中页表、不中饼内嵌表");
            }

            Finish(run, "srch-chart-title-only", fails);
        }

        private static void RunPageF(TestRun run, IList<string> filters, Session s)
        {
            bool need = Want(filters, "srch-fields-skeleton")
                || Want(filters, "srch-fields-fill")
                || Want(filters, "srch-fields-no-extra");
            if (!need)
            {
                return;
            }

            string html = ContentHtml.TextboxCreate(
                "填色框",
                10,
                12,
                36,
                16,
                "data-fill=\"#C00000\"");
            if (!s.TryNewSlide(out string err)
                || !s.TryApply(html, out _, out err)
                || !s.TryReadSearch(out PptHtmlReadResult tree, out err))
            {
                FailNeeded(run, filters, err, "srch-fields-skeleton", "srch-fields-fill", "srch-fields-no-extra");
                return;
            }

            if (Want(filters, "srch-fields-skeleton"))
            {
                var fails = new List<string>();
                if (!s.TrySearch(Query("text", "填色"), out SearchHit hit, out string se))
                {
                    fails.Add(se);
                }
                else
                {
                    CheckCommon(fails, hit, 1);
                    string frag = ExtractHtml(hit.Display);
                    if (frag.IndexOf("ShapeId=", StringComparison.Ordinal) < 0
                        || frag.IndexOf("data-shape-type=", StringComparison.Ordinal) < 0
                        || frag.IndexOf("style=", StringComparison.Ordinal) < 0)
                    {
                        fails.Add("骨架缺 ShapeId/type/style");
                    }

                    if (frag.IndexOf("data-font-size", StringComparison.Ordinal) >= 0)
                    {
                        fails.Add("骨架不应带详细字号皮");
                    }
                }

                Finish(run, "srch-fields-skeleton", fails);
            }

            if (Want(filters, "srch-fields-fill") || Want(filters, "srch-fields-no-extra"))
            {
                var args = Query("text", "填色");
                args["fields"] = new List<object> { "data-fill" };
                if (!s.TrySearch(args, out SearchHit hit, out string se))
                {
                    if (Want(filters, "srch-fields-fill"))
                    {
                        run.CaseFail("srch-fields-fill", se);
                    }

                    if (Want(filters, "srch-fields-no-extra"))
                    {
                        run.CaseFail("srch-fields-no-extra", se);
                    }
                }
                else
                {
                    if (Want(filters, "srch-fields-fill"))
                    {
                        var fails = new List<string>();
                        CheckCommon(fails, hit, 1);
                        string frag = ExtractHtml(hit.Display);
                        if (frag.IndexOf("data-fill=", StringComparison.Ordinal) < 0)
                        {
                            fails.Add("应有 data-fill");
                        }

                        if (frag.IndexOf("ShapeId=", StringComparison.Ordinal) < 0
                            || frag.IndexOf("data-shape-type=", StringComparison.Ordinal) < 0
                            || frag.IndexOf("style=", StringComparison.Ordinal) < 0)
                        {
                            fails.Add("ShapeId/type/style 仍在");
                        }

                        Finish(run, "srch-fields-fill", fails);
                    }

                    if (Want(filters, "srch-fields-no-extra"))
                    {
                        var fails = new List<string>();
                        string frag = ExtractHtml(hit.Display);
                        if (frag.IndexOf("data-src=", StringComparison.Ordinal) >= 0
                            || frag.IndexOf("data-chart-", StringComparison.Ordinal) >= 0)
                        {
                            fails.Add("没点的重皮不应出现");
                        }

                        Finish(run, "srch-fields-no-extra", fails);
                    }
                }
            }
        }

        private static void RunPageG(TestRun run, IList<string> filters, Session s)
        {
            bool needRank = Want(filters, "srch-pic-area-rank") || Want(filters, "srch-pic-flat");
            if (needRank)
            {
                string body = ContentHtml.MixPicMarkup(s.Assets.RedPng, 6, 8, 40, 30) + "\n"
                    + ContentHtml.MixPicMarkup(s.Assets.BluePng, 50, 8, 20, 20) + "\n"
                    + "  <div data-shape-type=\"group\" style=\"" + ContentHtml.Geo(6, 50, 30, 24) + "\">\n"
                    + ContentHtml.MixPicMarkup(s.Assets.GreenPng, 8, 52, 8, 8) + "\n"
                    + "  <div data-shape-type=\"textbox\" style=\"" + ContentHtml.Geo(18, 52, 14, 10) + "\">组辅</div>\n"
                    + "  </div>";
                if (!s.TryNewSlide(out string err)
                    || !s.TryApply(ContentHtml.Section(body), out _, out err)
                    || !s.TryReadSearch(out PptHtmlReadResult tree, out err))
                {
                    FailNeeded(run, filters, err, "srch-pic-area-rank", "srch-pic-flat");
                }
                else
                {
                    var args = Query("data-shape-type", "picture");
                    if (Want(filters, "srch-pic-area-rank"))
                    {
                        var fails = new List<string>();
                        if (!s.TrySearch(args, out SearchHit hit, out string se))
                        {
                            fails.Add(se);
                        }
                        else
                        {
                            CheckCommon(fails, hit, 3);
                            List<PptHtmlShapeNode> pics = CollectPics(hit.Forest);
                            if (pics.Count != 3)
                            {
                                fails.Add("应中三张 picture");
                            }
                            else
                            {
                                if (pics[0].SearchAreaRank != 1)
                                {
                                    fails.Add("最大应 rank=1");
                                }

                                for (int i = 1; i < pics.Count; i++)
                                {
                                    if ((pics[i - 1].SearchArea ?? 0) + 0.0001 < (pics[i].SearchArea ?? 0))
                                    {
                                        fails.Add("顺序应按面积降序");
                                        break;
                                    }
                                }

                                for (int i = 0; i < pics.Count; i++)
                                {
                                    string area = ExtractAttr(ExtractHtml(hit.Display), pics[i].ShapeId, "data-area");
                                    if (!Regex.IsMatch(area ?? "", @"^[0-9.]+$"))
                                    {
                                        fails.Add("data-area 应为纯数字");
                                        break;
                                    }
                                }
                            }
                        }

                        Finish(run, "srch-pic-area-rank", fails);
                    }

                    if (Want(filters, "srch-pic-flat"))
                    {
                        var fails = new List<string>();
                        if (!s.TrySearch(args, out SearchHit hit, out string se))
                        {
                            fails.Add(se);
                        }
                        else
                        {
                            CheckCommon(fails, hit, 3);
                            if (HasGroup(hit.Forest))
                            {
                                fails.Add("picture 向应扁平，不套组壳");
                            }
                        }

                        Finish(run, "srch-pic-flat", fails);
                    }
                }
            }

            if (Want(filters, "srch-pic-rasterized"))
            {
                if (!s.TryNewSlide(out string err)
                    || !s.TryApply(ContentHtml.PictureCreate(s.Assets.RedPng, 10, 10, 20, 16), out _, out err)
                    || !s.TryAddFreeform(out err))
                {
                    run.CaseFail("srch-pic-rasterized", err ?? "铺 freeform 失败");
                }
                else
                {
                    var fails = new List<string>();
                    if (!s.TrySearch(Query("data-rasterized-from", ".+"), out SearchHit some, out string se2))
                    {
                        fails.Add(se2);
                    }
                    else if (some.MatchCount < 1)
                    {
                        fails.Add("有 rasterize 时应能按属性筛中；引擎不默认排除 freeform");
                    }

                    Finish(run, "srch-pic-rasterized", fails);
                }
            }

            if (!Want(filters, "srch-pic-truncate"))
            {
                return;
            }

            var sb = new StringBuilder();
            sb.Append(ContentHtml.MixPicMarkup(s.Assets.RedPng, 2, 2, 36, 24)).Append('\n');
            for (int i = 0; i < 80; i++)
            {
                double left = 2 + (i % 10) * 9.5;
                double top = 30 + (i / 10) * 7.5;
                sb.Append(ContentHtml.MixPicMarkup(s.Assets.BluePng, left, top, 8, 6)).Append('\n');
            }

            if (!s.TryNewSlide(out string terr)
                || !s.TryApply(ContentHtml.Section(sb.ToString()), out _, out terr))
            {
                run.CaseFail("srch-pic-truncate", terr ?? "铺 81 图失败");
                return;
            }

            var tfails = new List<string>();
            if (!s.TrySearch(Query("data-shape-type", "picture"), out SearchHit thit, out string tse))
            {
                tfails.Add(tse ?? "截断不应失败");
            }
            else
            {
                if (thit.MatchCount != 80)
                {
                    tfails.Add("应回前 80，实际 " + thit.MatchCount);
                }

                if (thit.MatchTotal < 81 || !thit.PictureOriented)
                {
                    tfails.Add("截断前应 ≥81");
                }

                if (thit.Display.IndexOf("截断", StringComparison.Ordinal) < 0
                    && thit.Display.IndexOf("truncated", StringComparison.OrdinalIgnoreCase) < 0)
                {
                    tfails.Add("说明应含截断/truncated");
                }

                List<PptHtmlShapeNode> pics = CollectPics(thit.Forest);
                if (pics.Count > 0 && pics[0].SearchAreaRank != 1)
                {
                    tfails.Add("rank 仍按截断前面积，最大为 1");
                }
            }

            Finish(run, "srch-pic-truncate", tfails);
        }

        private static void SearchAssert(
            TestRun run,
            Session s,
            string id,
            Dictionary<string, object> args,
            int expectCount,
            IList<string> expectIds,
            Func<string, SearchHit, string> extra)
        {
            var fails = new List<string>();
            if (!s.TrySearch(args, out SearchHit hit, out string err))
            {
                fails.Add(err);
                Finish(run, id, fails);
                return;
            }

            CheckCommon(fails, hit, expectCount);
            HashSet<string> actual = LeafIds(hit.Forest);
            if (expectCount == 0)
            {
                if (actual.Count != 0)
                {
                    fails.Add("0 命中却有 ShapeId");
                }
            }
            else
            {
                for (int i = 0; i < expectIds.Count; i++)
                {
                    if (!string.IsNullOrEmpty(expectIds[i]) && !actual.Contains(expectIds[i]))
                    {
                        fails.Add("缺 id " + expectIds[i]);
                    }
                }

                if (actual.Count != expectCount)
                {
                    fails.Add("命中 id 数 " + actual.Count + " 期望 " + expectCount);
                }
            }

            if (extra != null)
            {
                string more = extra(hit.Display, hit);
                if (!string.IsNullOrEmpty(more))
                {
                    fails.Add(more);
                }
            }

            Finish(run, id, fails);
        }

        private static void CheckCommon(List<string> fails, SearchHit hit, int expectCount)
        {
            if (hit.Display.IndexOf("<section", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                fails.Add("成功回包禁止 <section>");
            }

            if (hit.Display.IndexOf("matches[", StringComparison.Ordinal) >= 0)
            {
                fails.Add("成功回包禁止 matches[]");
            }

            if (hit.MatchCount != expectCount)
            {
                fails.Add("match_count=" + hit.MatchCount + " 期望 " + expectCount);
            }
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

        private static Dictionary<string, object> ArgsQuery(int n)
        {
            var items = new List<object>();
            for (int i = 0; i < n; i++)
            {
                items.Add(new Dictionary<string, object>
                {
                    ["attr"] = "text",
                    ["pattern"] = "a" + i.ToString(CultureInfo.InvariantCulture)
                });
            }

            return new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase)
            {
                ["query"] = items
            };
        }

        private static Dictionary<string, object> Query(string attr, string pattern)
        {
            return new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase)
            {
                ["query"] = new List<object>
                {
                    new Dictionary<string, object>
                    {
                        ["attr"] = attr,
                        ["pattern"] = pattern
                    }
                }
            };
        }

        private static void AddClause(Dictionary<string, object> args, string attr, string pattern)
        {
            var list = (List<object>)args["query"];
            list.Add(new Dictionary<string, object>
            {
                ["attr"] = attr,
                ["pattern"] = pattern
            });
        }

        private static void AssertParseOk(List<string> fails, int n, string match, bool expectAny)
        {
            var args = ArgsQuery(n);
            if (!string.IsNullOrEmpty(match))
            {
                args["match"] = match;
            }

            if (!PptHtmlNodeSearch.TryParseRequest(args, out PptHtmlSearchRequest req, out string err))
            {
                fails.Add("n=" + n + " 解析失败: " + err);
                return;
            }

            if (req.MatchAny != expectAny)
            {
                fails.Add("n=" + n + " match 空应为 and");
            }
        }

        private static void AssertParseFail(List<string> fails, Dictionary<string, object> args, string contains)
        {
            if (PptHtmlNodeSearch.TryParseRequest(args, out _, out string err))
            {
                fails.Add("应失败却成功");
                return;
            }

            if (string.IsNullOrEmpty(err) || err.IndexOf(contains, StringComparison.OrdinalIgnoreCase) < 0)
            {
                fails.Add("失败文案应含「" + contains + "」实际 " + (err ?? ""));
            }
        }

        private static void AssertPictureFlag(List<string> fails, string pattern, bool expect)
        {
            if (!PptHtmlNodeSearch.TryParseRequest(
                    Query("data-shape-type", pattern),
                    out PptHtmlSearchRequest req,
                    out string err))
            {
                fails.Add(pattern + " 解析失败: " + err);
                return;
            }

            bool got = PptHtmlNodeSearch.IsPictureOrientedSearch(req);
            if (got != expect)
            {
                fails.Add("pattern=" + pattern + " picture向=" + got + " 期望 " + expect);
            }
        }

        private static void AssertToolFail(List<string> fails, Dictionary<string, object> args, string contains)
        {
            ToolResult result = InvokeHandleSearch(args);
            if (result == null)
            {
                fails.Add("HandleSearch 不可调");
                return;
            }

            if (result.Success)
            {
                fails.Add("应失败却成功");
                return;
            }

            string err = result.Error ?? "";
            if (err.IndexOf(contains, StringComparison.OrdinalIgnoreCase) < 0)
            {
                fails.Add("文案应含「" + contains + "」实际 " + err);
            }
        }

        private static ToolResult InvokeHandleSearch(Dictionary<string, object> args)
        {
            MethodInfo mi = typeof(F_ReadPptHtmlTool).GetMethod(
                "HandleSearch",
                BindingFlags.NonPublic | BindingFlags.Static);
            if (mi == null)
            {
                return null;
            }

            return (ToolResult)mi.Invoke(null, new object[] { args, new DummyPptChannel(), "1" });
        }

        private static string FindIdByText(PptHtmlReadResult tree, string text)
        {
            PptHtmlShapeNode n = Find(tree != null ? tree.Shapes : null, node =>
                !IsGroup(node)
                && !string.IsNullOrEmpty(node.Text)
                && node.Text.IndexOf(text, StringComparison.Ordinal) >= 0);
            return n != null ? n.ShapeId : null;
        }

        private static string FindIdByType(PptHtmlReadResult tree, string type)
        {
            PptHtmlShapeNode n = Find(tree != null ? tree.Shapes : null, node =>
                string.Equals(node.ShapeType, type, StringComparison.OrdinalIgnoreCase));
            return n != null ? n.ShapeId : null;
        }

        private static string FindParentId(PptHtmlReadResult tree, string childId)
        {
            if (string.IsNullOrEmpty(childId))
            {
                return null;
            }

            PptHtmlShapeNode parent = Find(tree != null ? tree.Shapes : null, node =>
                IsGroup(node) && FindChild(node, childId) != null);
            return parent != null ? parent.ShapeId : null;
        }

        private static PptHtmlShapeNode FindChild(PptHtmlShapeNode parent, string id)
        {
            if (parent == null || parent.Children == null)
            {
                return null;
            }

            for (int i = 0; i < parent.Children.Count; i++)
            {
                if (parent.Children[i] != null
                    && string.Equals(parent.Children[i].ShapeId, id, StringComparison.Ordinal))
                {
                    return parent.Children[i];
                }
            }

            return null;
        }

        private static PptHtmlShapeNode FindNode(IList<PptHtmlShapeNode> nodes, string id)
        {
            return Find(nodes, n => string.Equals(n.ShapeId, id, StringComparison.Ordinal));
        }

        private static PptHtmlShapeNode Find(
            IList<PptHtmlShapeNode> nodes,
            Func<PptHtmlShapeNode, bool> pred)
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

                if (pred(n))
                {
                    return n;
                }

                PptHtmlShapeNode kid = Find(n.Children, pred);
                if (kid != null)
                {
                    return kid;
                }
            }

            return null;
        }

        private static bool IsGroup(PptHtmlShapeNode n)
        {
            return n != null && string.Equals(n.ShapeType, "group", StringComparison.OrdinalIgnoreCase);
        }

        private static bool HasGroup(IList<PptHtmlShapeNode> nodes)
        {
            return Find(nodes, IsGroup) != null;
        }

        private static List<PptHtmlShapeNode> CollectPics(IList<PptHtmlShapeNode> forest)
        {
            var list = new List<PptHtmlShapeNode>();
            Walk(forest, n =>
            {
                if (string.Equals(n.ShapeType, "picture", StringComparison.OrdinalIgnoreCase))
                {
                    list.Add(n);
                }
            });
            return list;
        }

        private static void Walk(IList<PptHtmlShapeNode> nodes, Action<PptHtmlShapeNode> visit)
        {
            if (nodes == null)
            {
                return;
            }

            for (int i = 0; i < nodes.Count; i++)
            {
                if (nodes[i] == null)
                {
                    continue;
                }

                visit(nodes[i]);
                Walk(nodes[i].Children, visit);
            }
        }

        private static bool LooksPagePct(string style)
        {
            return !string.IsNullOrEmpty(style)
                && style.IndexOf("left:", StringComparison.OrdinalIgnoreCase) >= 0
                && style.IndexOf('%') >= 0
                && ContentAssert.TryParseGeo(style, out _, out _, out _, out _);
        }

        private static bool HasTag(string display)
        {
            string html = ExtractHtml(display);
            return html.IndexOf("<div", StringComparison.OrdinalIgnoreCase) >= 0
                || html.IndexOf("<img", StringComparison.OrdinalIgnoreCase) >= 0
                || html.IndexOf("<table", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private static HashSet<string> LeafIds(IList<PptHtmlShapeNode> forest)
        {
            var set = new HashSet<string>(StringComparer.Ordinal);
            Walk(forest, n =>
            {
                if (!IsGroup(n) && !string.IsNullOrEmpty(n.ShapeId))
                {
                    set.Add(n.ShapeId);
                }
            });
            return set;
        }

        private static HashSet<string> HitIds(string display)
        {
            var set = new HashSet<string>(StringComparer.Ordinal);
            string html = ExtractHtml(display);
            foreach (Match m in Regex.Matches(html, "ShapeId=\"([^\"]+)\""))
            {
                if (m.Success)
                {
                    set.Add(m.Groups[1].Value);
                }
            }

            return set;
        }

        private static string ExtractHtml(string display)
        {
            if (string.IsNullOrEmpty(display))
            {
                return "";
            }

            int i = display.IndexOf('<');
            return i >= 0 ? display.Substring(i) : "";
        }

        private static string ExtractAttr(string html, string shapeId, string attr)
        {
            if (string.IsNullOrEmpty(html) || string.IsNullOrEmpty(shapeId))
            {
                return null;
            }

            int at = html.IndexOf("ShapeId=\"" + shapeId + "\"", StringComparison.Ordinal);
            if (at < 0)
            {
                return null;
            }

            int start = html.LastIndexOf('<', at);
            int end = html.IndexOf('>', at);
            if (start < 0 || end < start)
            {
                return null;
            }

            string tag = html.Substring(start, end - start);
            Match m = Regex.Match(tag, attr + "=\"([^\"]+)\"");
            return m.Success ? m.Groups[1].Value : null;
        }

        private sealed class DummyPptChannel : IOperationChannel
        {
            public string ChannelId
            {
                get { return "srch-test"; }
            }

            public ChannelKind Kind
            {
                get { return ChannelKind.Ppt; }
            }
        }

        private sealed class SearchHit
        {
            public string Display;
            public int MatchCount;
            public int MatchTotal;
            public bool PictureOriented;
            public List<PptHtmlShapeNode> Forest;
        }

        private sealed class Session
        {
            public Session(object pres, ContentAssets assets, bool wpp)
            {
                Pres = pres;
                Assets = assets;
                Wpp = wpp;
            }

            public bool Wpp { get; private set; }
            public object Pres { get; private set; }
            public object Slide { get; private set; }
            public string SlideId { get; private set; }
            public ContentAssets Assets { get; private set; }

            public bool TryNewSlide(out string error)
            {
                error = null;
                try
                {
                    if (Wpp)
                    {
                        object slides = WppCom.GetProperty(Pres, "Slides");
                        int count = Convert.ToInt32(WppCom.GetProperty(slides, "Count") ?? 0);
                        try
                        {
                            Slide = WppCom.Invoke(slides, "Add", count + 1, 12);
                        }
                        catch
                        {
                            Slide = WppCom.Invoke(slides, "Add", count + 1, 1);
                        }

                        SlideId = Convert.ToInt32(
                                WppCom.GetProperty(Slide, "SlideID") ?? WppCom.GetProperty(Slide, "SlideId"))
                            .ToString(CultureInfo.InvariantCulture);
                        return true;
                    }

                    PowerPoint.Presentation ppt = (PowerPoint.Presentation)Pres;
                    PowerPoint.Slide slide = ppt.Slides.Add(
                        ppt.Slides.Count + 1,
                        PowerPoint.PpSlideLayout.ppLayoutBlank);
                    Slide = slide;
                    SlideId = slide.SlideID.ToString(CultureInfo.InvariantCulture);
                    return true;
                }
                catch (Exception ex)
                {
                    error = ex.Message;
                    return false;
                }
            }

            public bool TryApply(string html, out List<string> ids, out string error)
            {
                ids = new List<string>();
                if (!ContentHtml.TryParse(html, SlideId, out PptHtmlApplyPlan plan, out error))
                {
                    return false;
                }

                plan.AllowCreate = true;
                ContentHtml.ResolvePicturePaths(plan);
                PptHtmlApplyResult result;
                bool ok = Wpp
                    ? PptHtmlWppApplier.TryApply(Pres, plan, "srch-synth-wpp", out result, out error)
                    : PptHtmlPowerPointApplier.TryApply(
                        (PowerPoint.Presentation)Pres, plan, "srch-synth", out result, out error);
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

            public bool TryReadSearch(out PptHtmlReadResult tree, out string error)
            {
                if (Wpp)
                {
                    return PptHtmlWppReader.TryRead(
                        Pres,
                        SlideId,
                        "srch-synth-wpp",
                        "wpp",
                        null,
                        false,
                        true,
                        out tree,
                        out error);
                }

                return PptHtmlPowerPointReader.TryRead(
                    (PowerPoint.Presentation)Pres,
                    SlideId,
                    "srch-synth",
                    "ppt",
                    null,
                    false,
                    true,
                    out tree,
                    out error);
            }

            public bool TryAddFreeform(out string error)
            {
                error = null;
                try
                {
                    if (Wpp)
                    {
                        object shapes = WppCom.GetProperty(Slide, "Shapes");
                        object builder = WppCom.Invoke(shapes, "BuildFreeform", 1, 320f, 40f);
                        WppCom.Invoke(builder, "AddNodes", 0, 0, 400f, 40f);
                        WppCom.Invoke(builder, "AddNodes", 0, 0, 360f, 110f);
                        WppCom.Invoke(builder, "AddNodes", 0, 0, 320f, 40f);
                        WppCom.Invoke(builder, "ConvertToShape");
                        return true;
                    }

                    PowerPoint.Slide slide = (PowerPoint.Slide)Slide;
                    PowerPoint.FreeformBuilder b = slide.Shapes.BuildFreeform(
                        Office.MsoEditingType.msoEditingCorner,
                        320f,
                        40f);
                    b.AddNodes(Office.MsoSegmentType.msoSegmentLine, Office.MsoEditingType.msoEditingAuto, 400f, 40f);
                    b.AddNodes(Office.MsoSegmentType.msoSegmentLine, Office.MsoEditingType.msoEditingAuto, 360f, 110f);
                    b.AddNodes(Office.MsoSegmentType.msoSegmentLine, Office.MsoEditingType.msoEditingAuto, 320f, 40f);
                    b.ConvertToShape();
                    return true;
                }
                catch (Exception ex)
                {
                    error = ex.Message;
                    return false;
                }
            }

            public bool TrySearch(Dictionary<string, object> args, out SearchHit hit, out string error)
            {
                hit = null;
                if (!TryReadSearch(out PptHtmlReadResult tree, out error))
                {
                    return false;
                }

                if (!PptHtmlNodeSearch.TryParseRequest(args, out PptHtmlSearchRequest request, out error))
                {
                    return false;
                }

                if (!PptHtmlNodeSearch.TryMatch(
                        tree.Shapes,
                        request,
                        out List<PptHtmlShapeNode> forest,
                        out int leafCount,
                        out int matchTotal,
                        out bool pictureOriented,
                        out error))
                {
                    return false;
                }

                if (request.Fields != null && request.Fields.Count > 0 && leafCount > 0)
                {
                    var cache = new Dictionary<string, PptHtmlShapeNode>(StringComparer.Ordinal);
                    if (!PptHtmlNodeSearch.TryProjectForest(
                            forest,
                            request.Fields,
                            id => LoadDetail(id, cache),
                            out error))
                    {
                        return false;
                    }
                }

                hit = new SearchHit
                {
                    Forest = forest,
                    MatchCount = leafCount,
                    MatchTotal = matchTotal,
                    PictureOriented = pictureOriented,
                    Display = PptHtmlNodeSearch.BuildDisplayContents(
                        forest,
                        leafCount,
                        matchTotal,
                        pictureOriented,
                        request.Fields,
                        request.MatchAny,
                        request.Clauses != null ? request.Clauses.Count : 0)
                };
                return true;
            }

            private PptHtmlShapeNode LoadDetail(string shapeId, Dictionary<string, PptHtmlShapeNode> cache)
            {
                if (cache.TryGetValue(shapeId, out PptHtmlShapeNode hit))
                {
                    return hit;
                }

                bool ok = Wpp
                    ? PptHtmlWppReader.TryRead(
                        Pres,
                        SlideId,
                        "srch-synth-wpp",
                        "wpp",
                        shapeId,
                        false,
                        false,
                        out PptHtmlReadResult detail,
                        out _)
                    : PptHtmlPowerPointReader.TryRead(
                        (PowerPoint.Presentation)Pres,
                        SlideId,
                        "srch-synth",
                        "ppt",
                        shapeId,
                        false,
                        false,
                        out detail,
                        out _);
                if (!ok)
                {
                    return null;
                }

                PptHtmlShapeNode node = Find(
                    detail != null ? detail.Shapes : null,
                    n => string.Equals(n.ShapeId, shapeId, StringComparison.Ordinal));
                if (node != null)
                {
                    cache[shapeId] = node;
                }

                return node;
            }
        }
    }
}
