using System;
using System.Collections.Generic;
using System.Globalization;
using System.Reflection;
using System.Threading.Tasks;
using WordAddIn1;
using WordAddIn1.OpenFiles;
using WordAddIn1.PresentationHost;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PptHtmlContentRoundtripTest
{
    /// <summary>空白页 apply 再 F_capture_ppt_shape。前缀 cap-。</summary>
    internal static class CaptureShapeRunner
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
                if (f.IndexOf("cap-", StringComparison.OrdinalIgnoreCase) >= 0
                    || string.Equals(f, "cap", StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
            }

            return false;
        }

        public static void Run(TestRun run, IList<string> nameFilters, bool useWpp)
        {
            Console.WriteLine(useWpp
                ? "形状截图 · WPP · apply 再 cap"
                : "形状截图 · PPT · apply 再 cap");
            RunParse(run, nameFilters);
            if (NeedCom(nameFilters))
            {
                RunCom(run, nameFilters, useWpp);
            }

            Console.WriteLine("形状截图 通过 " + run.CasesPassed
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
                "cap-single-", "cap-grid-", "cap-neg-missing", "cap-neg-hidden",
                "cap-neg-cross", "cap-neg-grid-one"
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
                if (string.Equals(f, "cap", StringComparison.OrdinalIgnoreCase)
                    || string.Equals(f, "cap-", StringComparison.OrdinalIgnoreCase))
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
            Console.WriteLine("--- A 解析 / 互斥（不启 COM）---");
            ExpectFail(run, filters, "cap-neg-neither-fail", new Dictionary<string, object>(), "其一");
            ExpectFail(
                run,
                filters,
                "cap-neg-both-fail",
                Args(shapeId: "sid1-s1", shapeIds: new[] { "sid1-s2" }),
                "其一");
            ExpectFail(run, filters, "cap-neg-slide-id-fail", Args(shapeId: "sid1-s1", extra: "slide_id", extraVal: "1"), "slide_id");
            ExpectFail(run, filters, "cap-neg-page-fail", Args(shapeId: "sid1-s1", extra: "page_number", extraVal: 1), "F_capture");
            ExpectFail(run, filters, "cap-neg-kb-fail", Args(shapeId: "sid1-s1", extra: "storage_doc_uuid", extraVal: "u"), "ppt/wpp");
            ExpectFail(run, filters, "cap-neg-sheet-fail", Args(shapeId: "sid1-s1", extra: "sheet", extraVal: "Sheet1"), "sheet");
            ExpectFail(
                run,
                filters,
                "cap-neg-source-channel-fail",
                Args(shapeId: "sid1-s1", extra: "source_channel_id", extraVal: "p1"),
                "channel_id");
            ExpectFail(
                run,
                filters,
                "cap-neg-ids-string-fail",
                new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase) { ["shape_ids"] = "sid1-s1" },
                "数组");
            ExpectFail(
                run,
                filters,
                "cap-neg-ids-empty-item-fail",
                new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase)
                {
                    ["shape_ids"] = new List<object> { "sid1-s1", "" }
                },
                "空项");

            if (Want(filters, "cap-neg-over9-fail"))
            {
                var ids = new List<object>();
                for (int i = 1; i <= 10; i++)
                {
                    ids.Add("sid1-s" + i.ToString(CultureInfo.InvariantCulture));
                }

                ExpectFail(
                    run,
                    filters,
                    "cap-neg-over9-fail",
                    new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase) { ["shape_ids"] = ids },
                    "9");
            }

            if (Want(filters, "cap-old-capture-shape-fail"))
            {
                var fails = new List<string>();
                ToolResult old = InvokeOldCapture(new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase)
                {
                    ["shape_id"] = "sid1-s1"
                });
                if (old == null || old.Success)
                {
                    fails.Add("旧入口应失败");
                }
                else if ((old.Error ?? "").IndexOf("F_capture_ppt_shape", StringComparison.OrdinalIgnoreCase) < 0)
                {
                    fails.Add("文案应含 F_capture_ppt_shape，实际 " + (old.Error ?? ""));
                }

                Finish(run, "cap-old-capture-shape-fail", fails);
            }
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
            ToolResult r = InvokeShapeTool(args);
            if (r == null)
            {
                fails.Add("工具不可调");
            }
            else if (r.Success)
            {
                fails.Add("应失败却成功");
            }
            else if ((r.Error ?? "").IndexOf(contains, StringComparison.OrdinalIgnoreCase) < 0)
            {
                fails.Add("文案应含「" + contains + "」实际 " + (r.Error ?? ""));
            }

            if (HasImage(r))
            {
                fails.Add("失败不应带图");
            }

            Finish(run, id, fails);
        }

        private static void RunCom(TestRun run, IList<string> filters, bool useWpp)
        {
            Console.WriteLine(useWpp ? "--- B–D COM（WPP）---" : "--- B–D COM ---");
            ContentAssets assets = ContentAssetsIo.Ensure(
                System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "cap-synth-assets"));
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
                run.CaseFail("cap-", "COM 失败: " + ex.Message);
            }
            finally
            {
                ClosePpt(ref app, ref pres);
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
                    run.CaseFail("cap-", startErr);
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
                    run.CaseFail("cap-", "WPP Add 失败");
                    return;
                }

                var session = new Session(pres, assets, wpp: true);
                RunPages(run, filters, session);
            }
            catch (Exception ex)
            {
                run.CaseFail("cap-", "COM 失败: " + ex.Message);
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

        private static void ClosePpt(ref PowerPoint.Application app, ref PowerPoint.Presentation pres)
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

        private static void RunPages(TestRun run, IList<string> filters, Session s)
        {
            RunSinglePage(run, filters, s);
            RunGridPage(run, filters, s);
            RunCrossPage(run, filters, s);
        }

        private static void RunSinglePage(TestRun run, IList<string> filters, Session s)
        {
            bool need = Want(filters, "cap-single-")
                || Want(filters, "cap-neg-missing")
                || Want(filters, "cap-neg-hidden")
                || Want(filters, "cap-neg-grid-one");
            if (!need)
            {
                return;
            }

            string pie = ContentHtml.ChartFrag("pie2d", 52, 8, 36, 36);
            pie = pie.Replace("data-chart-type=\"pie2d\"", "data-chart-type=\"pie2d\" data-title=\"电力占比\"");
            string html = ContentHtml.Section(
                "  <div data-shape-type=\"textbox\" style=\"" + ContentHtml.Geo(6, 8, 28, 12) + "\">电力</div>\n"
                + ContentHtml.MixPicMarkup(s.Assets.RedPng, 6, 24, 40, 30) + "\n"
                + pie + "\n"
                + ContentHtml.MixPicMarkup(s.Assets.BluePng, 6, 72, 3, 3) + "\n"
                + "  <div data-shape-type=\"textbox\" style=\"" + ContentHtml.Geo(50, 52, 18, 10) + "\">甲</div>\n"
                + "  <div data-shape-type=\"textbox\" style=\"" + ContentHtml.Geo(70, 68, 18, 10) + "\">乙</div>");
            if (!s.TryNewSlide(out string err)
                || !s.TryApply(html, out _, out err)
                || !s.TryReadSearch(out PptHtmlReadResult tree, out err))
            {
                FailNeeded(run, filters, err, "cap-single-textbox", "cap-single-picture", "cap-single-chart",
                    "cap-single-group", "cap-single-small-scale", "cap-single-ids-one",
                    "cap-neg-missing-id-fail", "cap-neg-hidden-fail", "cap-neg-grid-one-bad-fail");
                return;
            }

            string idBox = FindIdByText(tree, "电力");
            string idPic = FindPicNear(tree, 40, 30);
            string idChart = FindIdByType(tree, "chart");
            string idSmall = FindPicNear(tree, 3, 3);
            string idJia = FindIdByText(tree, "甲");
            string idYi = FindIdByText(tree, "乙");
            var groupSess = new ContentGroupSession
            {
                Wpp = s.Wpp,
                Presentation = s.Pres,
                Slide = s.Slide,
                SlideId = s.SlideId
            };
            if (string.IsNullOrEmpty(idJia) || string.IsNullOrEmpty(idYi)
                || !groupSess.TryGroup(new[] { idJia, idYi }, out _, out err)
                || !s.TryReadSearch(out tree, out err))
            {
                FailNeeded(run, filters, err ?? "编组失败", "cap-single-group", "cap-neg-grid-one-bad-fail");
            }

            string idGroup = FindParentId(tree, idJia);
            IOperationChannel ch = s.MakeChannel();

            if (Want(filters, "cap-single-textbox"))
            {
                AssertSingle(run, ch, "cap-single-textbox", idBox, 6, 8, 28, 12, expectScale: false);
            }

            if (Want(filters, "cap-single-picture"))
            {
                AssertSingle(run, ch, "cap-single-picture", idPic, 6, 24, 40, 30, expectScale: false);
            }

            if (Want(filters, "cap-single-chart"))
            {
                AssertSingle(run, ch, "cap-single-chart", idChart, 52, 8, 36, 36, expectScale: false, notFullPage: true);
            }

            if (Want(filters, "cap-single-group"))
            {
                AssertGroup(run, ch, idGroup, idJia, idYi);
            }

            if (Want(filters, "cap-single-small-scale"))
            {
                AssertSingle(run, ch, "cap-single-small-scale", idSmall, 6, 72, 3, 3, expectScale: true);
            }

            if (Want(filters, "cap-single-ids-one"))
            {
                var fails = new List<string>();
                ToolResult r = InvokeGrid(ch, new List<string> { idBox });
                CheckOk(fails, r, idBox, "shape");
                if (GetInt(r, "grid_cols") != 0)
                {
                    fails.Add("一条 shape_ids 不应写 grid_cols");
                }

                Finish(run, "cap-single-ids-one", fails);
            }

            if (Want(filters, "cap-neg-missing-id-fail"))
            {
                AssertFail(run, "cap-neg-missing-id-fail", InvokeSingle(ch, "sid" + s.SlideId + "-s999999"), "未找到");
            }

            if (Want(filters, "cap-neg-grid-one-bad-fail"))
            {
                var bad = new List<string> { idBox, idPic, idChart, "sid" + s.SlideId + "-s999998" };
                AssertFail(run, "cap-neg-grid-one-bad-fail", InvokeGrid(ch, bad), "未找到");
            }

            if (Want(filters, "cap-neg-hidden-fail"))
            {
                if (!s.TrySetVisible(idBox, false, out err))
                {
                    run.CaseFail("cap-neg-hidden-fail", err ?? "隐藏失败");
                }
                else
                {
                    AssertFail(run, "cap-neg-hidden-fail", InvokeSingle(ch, idBox), "隐藏");
                    s.TrySetVisible(idBox, true, out _);
                }
            }
        }

        private static void RunGridPage(TestRun run, IList<string> filters, Session s)
        {
            if (!Want(filters, "cap-grid-"))
            {
                return;
            }

            string body = ContentHtml.MixPicMarkup(s.Assets.RedPng, 4, 6, 8, 8) + "\n"
                + ContentHtml.MixPicMarkup(s.Assets.BluePng, 30, 6, 40, 30) + "\n"
                + ContentHtml.MixPicMarkup(s.Assets.GreenPng, 4, 42, 20, 16) + "\n"
                + ContentHtml.MixPicMarkup(s.Assets.RedPng, 50, 42, 16, 16) + "\n"
                + ContentHtml.MixPicMarkup(s.Assets.BluePng, 72, 42, 12, 12) + "\n"
                + "  <div data-shape-type=\"textbox\" style=\"" + ContentHtml.Geo(4, 72, 16, 10) + "\">丙</div>\n"
                + "  <div data-shape-type=\"textbox\" style=\"" + ContentHtml.Geo(24, 80, 16, 10) + "\">丁</div>";
            if (!s.TryNewSlide(out string err)
                || !s.TryApply(ContentHtml.Section(body), out _, out err)
                || !s.TryReadSearch(out PptHtmlReadResult tree, out err))
            {
                FailNeeded(run, filters, err, "cap-grid-2x2", "cap-grid-3x3", "cap-grid-group-mix", "cap-grid-order");
                return;
            }

            string idSmall = FindPicNear(tree, 8, 8);
            string idBig = FindPicNear(tree, 40, 30);
            string idMid = FindPicNear(tree, 20, 16);
            string idD = FindPicNear(tree, 16, 16);
            string idE = FindPicNear(tree, 12, 12);
            string idBing = FindIdByText(tree, "丙");
            string idDing = FindIdByText(tree, "丁");
            var groupSess = new ContentGroupSession
            {
                Wpp = s.Wpp,
                Presentation = s.Pres,
                Slide = s.Slide,
                SlideId = s.SlideId
            };
            if (!groupSess.TryGroup(new[] { idBing, idDing }, out _, out err)
                || !s.TryReadSearch(out tree, out err))
            {
                FailNeeded(run, filters, err ?? "宫格页编组失败", "cap-grid-group-mix");
            }

            string idGroup = FindParentId(tree, idBing);
            IOperationChannel ch = s.MakeChannel();
            var four = new List<string> { idSmall, idBig, idMid, idD };
            var five = new List<string> { idSmall, idBig, idMid, idD, idE };

            if (Want(filters, "cap-grid-2x2"))
            {
                AssertGrid(run, "cap-grid-2x2", InvokeGrid(ch, four), four, 2, 2);
            }

            if (Want(filters, "cap-grid-3x3"))
            {
                AssertGrid(run, "cap-grid-3x3", InvokeGrid(ch, five), five, 3, 3);
            }

            if (Want(filters, "cap-grid-group-mix"))
            {
                var mix = new List<string> { idGroup, idBig, idMid };
                AssertGrid(run, "cap-grid-group-mix", InvokeGrid(ch, mix), mix, 2, 2);
            }

            if (Want(filters, "cap-grid-order"))
            {
                var order = new List<string> { idSmall, idBig, idMid };
                AssertGrid(run, "cap-grid-order", InvokeGrid(ch, order), order, 2, 2);
            }
        }

        private static void RunCrossPage(TestRun run, IList<string> filters, Session s)
        {
            if (!Want(filters, "cap-neg-cross-page-fail"))
            {
                return;
            }

            if (!s.TryNewSlide(out string err)
                || !s.TryApply(ContentHtml.PictureCreate(s.Assets.RedPng, 10, 10, 20, 16), out _, out err)
                || !s.TryReadSearch(out PptHtmlReadResult a, out err))
            {
                run.CaseFail("cap-neg-cross-page-fail", err ?? "第一页失败");
                return;
            }

            string idA = FindIdByType(a, "picture");
            if (!s.TryNewSlide(out err)
                || !s.TryApply(ContentHtml.PictureCreate(s.Assets.BluePng, 10, 10, 20, 16), out _, out err)
                || !s.TryReadSearch(out PptHtmlReadResult b, out err))
            {
                run.CaseFail("cap-neg-cross-page-fail", err ?? "第二页失败");
                return;
            }

            string idB = FindIdByType(b, "picture");
            AssertFail(
                run,
                "cap-neg-cross-page-fail",
                InvokeGrid(s.MakeChannel(), new List<string> { idA, idB }),
                "同一页");
        }

        private static void AssertSingle(
            TestRun run,
            IOperationChannel ch,
            string id,
            string shapeId,
            double left,
            double top,
            double width,
            double height,
            bool expectScale,
            bool notFullPage = false)
        {
            var fails = new List<string>();
            if (string.IsNullOrEmpty(shapeId))
            {
                fails.Add("未找到目标 ShapeId");
                Finish(run, id, fails);
                return;
            }

            ToolResult r = InvokeSingle(ch, shapeId);
            CheckOk(fails, r, shapeId, "shape");
            CheckBounds(fails, r, left, top, width, height);
            if (notFullPage)
            {
                double bw = GetBound(r, "width");
                double bh = GetBound(r, "height");
                if (bw > 90 || bh > 90)
                {
                    fails.Add("应是形状外接矩形，不像整页");
                }
            }

            if (expectScale)
            {
                double scale = GetDouble(r, "scale");
                int w = GetInt(r, "width_px");
                int h = GetInt(r, "height_px");
                if (scale <= 1.01)
                {
                    fails.Add("过小应上采样 scale>1，实际 " + scale.ToString("0.###", CultureInfo.InvariantCulture));
                }

                if (Math.Max(w, h) < 200)
                {
                    fails.Add("最长边应约≥256，实际 " + Math.Max(w, h));
                }
            }

            Finish(run, id, fails);
        }

        private static void AssertGroup(TestRun run, IOperationChannel ch, string idGroup, string idJia, string idYi)
        {
            var fails = new List<string>();
            if (string.IsNullOrEmpty(idGroup))
            {
                fails.Add("未找到组");
                Finish(run, "cap-single-group", fails);
                return;
            }

            ToolResult r = InvokeSingle(ch, idGroup);
            CheckOk(fails, r, idGroup, "shape");
            double gl = GetBound(r, "left");
            double gt = GetBound(r, "top");
            double gw = GetBound(r, "width");
            double gh = GetBound(r, "height");
            if (gw < 20 || gh < 16)
            {
                fails.Add("组框应覆盖两框，实际 " + gw.ToString("0.#", CultureInfo.InvariantCulture)
                    + "x" + gh.ToString("0.#", CultureInfo.InvariantCulture));
            }

            Finish(run, "cap-single-group", fails);
        }

        private static void AssertGrid(
            TestRun run,
            string id,
            ToolResult r,
            IList<string> expectIds,
            int cols,
            int rows)
        {
            var fails = new List<string>();
            if (r == null || !r.Success)
            {
                fails.Add(r == null ? "无回包" : (r.Error ?? "失败"));
                Finish(run, id, fails);
                return;
            }

            if (HasNoImage(r))
            {
                fails.Add("成功应有图");
            }

            if (!string.Equals(GetString(r, "capture_mode"), "shape_contact_sheet", StringComparison.Ordinal))
            {
                fails.Add("capture_mode 应为 shape_contact_sheet");
            }

            if (GetInt(r, "grid_cols") != cols || GetInt(r, "grid_rows") != rows)
            {
                fails.Add("网格 " + GetInt(r, "grid_cols") + "x" + GetInt(r, "grid_rows")
                    + " 期望 " + cols + "x" + rows);
            }

            if (GetInt(r, "cell_count") != expectIds.Count)
            {
                fails.Add("cell_count=" + GetInt(r, "cell_count") + " 期望 " + expectIds.Count);
            }

            IList<string> got = GetIdList(r);
            if (got.Count != expectIds.Count)
            {
                fails.Add("回显条数 " + got.Count);
            }
            else
            {
                for (int i = 0; i < expectIds.Count; i++)
                {
                    if (!string.Equals(got[i], expectIds[i], StringComparison.Ordinal))
                    {
                        fails.Add("顺序被重排，第 " + (i + 1) + " 期望 " + expectIds[i] + " 实际 " + got[i]);
                        break;
                    }
                }
            }

            Finish(run, id, fails);
        }

        private static void AssertFail(TestRun run, string id, ToolResult r, string contains)
        {
            var fails = new List<string>();
            if (r == null)
            {
                fails.Add("无回包");
            }
            else if (r.Success)
            {
                fails.Add("应失败却成功");
            }
            else if ((r.Error ?? "").IndexOf(contains, StringComparison.OrdinalIgnoreCase) < 0)
            {
                fails.Add("文案应含「" + contains + "」实际 " + (r.Error ?? ""));
            }

            if (HasImage(r))
            {
                fails.Add("失败不应带图");
            }

            Finish(run, id, fails);
        }

        private static void CheckOk(List<string> fails, ToolResult r, string shapeId, string mode)
        {
            if (r == null || !r.Success)
            {
                fails.Add(r == null ? "无回包" : (r.Error ?? "失败"));
                return;
            }

            if (HasNoImage(r))
            {
                fails.Add("成功应有 image_base64");
            }

            if (GetInt(r, "width_px") <= 0 || GetInt(r, "height_px") <= 0)
            {
                fails.Add("宽高应 > 0");
            }

            if (!string.Equals(GetString(r, "capture_mode"), mode, StringComparison.Ordinal))
            {
                fails.Add("capture_mode=" + GetString(r, "capture_mode") + " 期望 " + mode);
            }

            if (!string.IsNullOrEmpty(shapeId)
                && !string.Equals(GetString(r, "shape_id"), shapeId, StringComparison.Ordinal)
                && mode == "shape")
            {
                fails.Add("回显 id 应为 " + shapeId);
            }
        }

        private static void CheckBounds(
            List<string> fails,
            ToolResult r,
            double left,
            double top,
            double width,
            double height)
        {
            if (r == null || !r.Success)
            {
                return;
            }

            if (Math.Abs(GetBound(r, "left") - left) > 2.5
                || Math.Abs(GetBound(r, "top") - top) > 2.5
                || Math.Abs(GetBound(r, "width") - width) > 3.5
                || Math.Abs(GetBound(r, "height") - height) > 3.5)
            {
                fails.Add("bounds_pct 与铺页差过大 "
                    + GetBound(r, "left").ToString("0.#", CultureInfo.InvariantCulture) + ","
                    + GetBound(r, "top").ToString("0.#", CultureInfo.InvariantCulture) + " "
                    + GetBound(r, "width").ToString("0.#", CultureInfo.InvariantCulture) + "x"
                    + GetBound(r, "height").ToString("0.#", CultureInfo.InvariantCulture));
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

        private static Dictionary<string, object> Args(
            string shapeId = null,
            string[] shapeIds = null,
            string extra = null,
            object extraVal = null)
        {
            var d = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);
            if (shapeId != null)
            {
                d["shape_id"] = shapeId;
            }

            if (shapeIds != null)
            {
                d["shape_ids"] = new List<object>(shapeIds);
            }

            if (extra != null)
            {
                d[extra] = extraVal;
            }

            return d;
        }

        private static ToolResult InvokeShapeTool(Dictionary<string, object> args)
        {
            var tools = new Dictionary<string, Func<Dictionary<string, object>, Task<ToolResult>>>(
                StringComparer.OrdinalIgnoreCase);
            F_CapturePptShapeTool.Register(tools, null);
            if (!tools.TryGetValue("F_capture_ppt_shape", out Func<Dictionary<string, object>, Task<ToolResult>> fn))
            {
                return null;
            }

            return fn(args).GetAwaiter().GetResult();
        }

        private static ToolResult InvokeOldCapture(Dictionary<string, object> args)
        {
            var tools = new Dictionary<string, Func<Dictionary<string, object>, Task<ToolResult>>>(
                StringComparer.OrdinalIgnoreCase);
            F_CaptureDocumentPageImageTool.Register(tools, null);
            if (!tools.TryGetValue("F_capture_document_page_image", out Func<Dictionary<string, object>, Task<ToolResult>> fn))
            {
                return null;
            }

            return fn(args).GetAwaiter().GetResult();
        }

        private static ToolResult InvokeSingle(IOperationChannel channel, string shapeId)
        {
            return InvokePrivate("CaptureSingle", new object[] { channel, shapeId });
        }

        private static ToolResult InvokeGrid(IOperationChannel channel, List<string> ids)
        {
            return InvokePrivate("CaptureGrid", new object[] { channel, ids });
        }

        private static ToolResult InvokePrivate(string name, object[] args)
        {
            MethodInfo mi = typeof(F_CapturePptShapeTool).GetMethod(
                name,
                BindingFlags.NonPublic | BindingFlags.Static);
            if (mi == null)
            {
                return new ToolResult { Success = false, Error = "找不到 " + name };
            }

            try
            {
                return (ToolResult)mi.Invoke(null, args);
            }
            catch (TargetInvocationException ex)
            {
                return new ToolResult
                {
                    Success = false,
                    Error = (ex.InnerException ?? ex).Message
                };
            }
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

        private static bool HasNoImage(ToolResult r)
        {
            return !HasImage(r);
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

        private static double GetDouble(ToolResult r, string key)
        {
            var d = r != null ? r.Data as Dictionary<string, object> : null;
            if (d == null || !d.TryGetValue(key, out object v) || v == null)
            {
                return 0;
            }

            try
            {
                return Convert.ToDouble(v, CultureInfo.InvariantCulture);
            }
            catch
            {
                return 0;
            }
        }

        private static double GetBound(ToolResult r, string key)
        {
            var d = r != null ? r.Data as Dictionary<string, object> : null;
            if (d == null || !d.TryGetValue("bounds_pct", out object raw))
            {
                return 0;
            }

            var b = raw as Dictionary<string, object>;
            if (b == null || !b.TryGetValue(key, out object v) || v == null)
            {
                return 0;
            }

            try
            {
                return Convert.ToDouble(v, CultureInfo.InvariantCulture);
            }
            catch
            {
                return 0;
            }
        }

        private static IList<string> GetIdList(ToolResult r)
        {
            var list = new List<string>();
            var d = r != null ? r.Data as Dictionary<string, object> : null;
            if (d == null || !d.TryGetValue("shape_ids", out object raw) || raw == null)
            {
                return list;
            }

            if (raw is IEnumerable<object> en)
            {
                foreach (object one in en)
                {
                    list.Add(Convert.ToString(one) ?? "");
                }
            }
            else if (raw is System.Collections.IEnumerable any)
            {
                foreach (object one in any)
                {
                    if (one is string || one is int)
                    {
                        list.Add(Convert.ToString(one) ?? "");
                    }
                }
            }

            return list;
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

        private static string FindPicNear(PptHtmlReadResult tree, double width, double height)
        {
            PptHtmlShapeNode n = Find(tree != null ? tree.Shapes : null, node =>
            {
                if (!string.Equals(node.ShapeType, "picture", StringComparison.OrdinalIgnoreCase)
                    || !ContentAssert.TryParseGeo(node.Style, out _, out _, out double w, out double h))
                {
                    return false;
                }

                return Math.Abs(w - width) < 3 && Math.Abs(h - height) < 3;
            });
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

        private static bool IsGroup(PptHtmlShapeNode n)
        {
            return n != null && string.Equals(n.ShapeType, "group", StringComparison.OrdinalIgnoreCase);
        }

        private static PptHtmlShapeNode Find(IList<PptHtmlShapeNode> nodes, Func<PptHtmlShapeNode, bool> pred)
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

            public IOperationChannel MakeChannel()
            {
                string uuid = Guid.NewGuid().ToString("N");
                return Wpp
                    ? (IOperationChannel)new WppChannel("cap-wpp", uuid, Pres, "")
                    : new PptChannel("cap-ppt", uuid, (PowerPoint.Presentation)Pres, "");
            }

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
                    ? PptHtmlWppApplier.TryApply(Pres, plan, "cap-wpp", out result, out error)
                    : PptHtmlPowerPointApplier.TryApply(
                        (PowerPoint.Presentation)Pres, plan, "cap-ppt", out result, out error);
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
                        Pres, SlideId, "cap-wpp", "wpp", null, false, true, out tree, out error);
                }

                return PptHtmlPowerPointReader.TryRead(
                    (PowerPoint.Presentation)Pres, SlideId, "cap-ppt", "ppt", null, false, true, out tree, out error);
            }

            public bool TrySetVisible(string shapeId, bool visible, out string error)
            {
                error = null;
                if (!PptShapeId.TryParseShape(shapeId, out _, out int comId))
                {
                    error = "id 无法解析";
                    return false;
                }

                try
                {
                    if (Wpp)
                    {
                        object shapes = WppCom.GetProperty(Slide, "Shapes");
                        int count = Convert.ToInt32(WppCom.GetProperty(shapes, "Count") ?? 0);
                        for (int i = 1; i <= count; i++)
                        {
                            object shape = WppCom.GetIndexed(shapes, i);
                            int sid = Convert.ToInt32(WppCom.GetProperty(shape, "Id") ?? 0);
                            if (sid == comId)
                            {
                                WppCom.TrySetProperty(shape, "Visible", visible ? -1 : 0);
                                return true;
                            }
                        }

                        error = "找不到形状";
                        return false;
                    }

                    PowerPoint.Shapes shapesPpt = ((PowerPoint.Slide)Slide).Shapes;
                    for (int i = 1; i <= shapesPpt.Count; i++)
                    {
                        PowerPoint.Shape shape = shapesPpt[i];
                        if (shape.Id == comId)
                        {
                            shape.Visible = visible
                                ? Office.MsoTriState.msoTrue
                                : Office.MsoTriState.msoFalse;
                            return true;
                        }
                    }

                    error = "找不到形状";
                    return false;
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
