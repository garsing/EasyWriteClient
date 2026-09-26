using System;
using System.Collections.Generic;
using System.Globalization;
using System.Threading.Tasks;
using Newtonsoft.Json;
using WordAddIn1;
using WordAddIn1.OpenFiles;
using WordAddIn1.PresentationHost;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PptHtmlContentRoundtripTest
{
    /// <summary>F_get_presentation_content。前缀 ov-。</summary>
    internal static class OverviewRunner
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
                if (f.IndexOf("ov-", StringComparison.OrdinalIgnoreCase) >= 0
                    || string.Equals(f, "ov", StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
            }

            return false;
        }

        public static void Run(TestRun run, IList<string> nameFilters, bool useWpp)
        {
            Console.WriteLine(useWpp
                ? "演示文稿概览 · WPP"
                : "演示文稿概览 · PPT");
            RunChannel(run, nameFilters);
            if (NeedCom(nameFilters))
            {
                RunCom(run, nameFilters, useWpp);
            }

            Console.WriteLine("概览 通过 " + run.CasesPassed
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
                "ov-count-", "ov-id-stable", "ov-hidden", "ov-slim",
                "ov-unsaved", "ov-display", "ov-no-html"
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
                if (string.Equals(f, "ov", StringComparison.OrdinalIgnoreCase)
                    || string.Equals(f, "ov-", StringComparison.OrdinalIgnoreCase))
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

        private static void RunChannel(TestRun run, IList<string> filters)
        {
            Console.WriteLine("--- A 渠道（不启演示 COM）---");
            if (Want(filters, "ov-neg-no-channel"))
            {
                var fails = new List<string>();
                ToolResult r = InvokeTool(new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase));
                if (r == null)
                {
                    fails.Add("工具不可调");
                }
                else if (r.Success)
                {
                    fails.Add("无渠道应失败");
                }

                Finish(run, "ov-neg-no-channel", fails);
            }

            if (Want(filters, "ov-neg-wrong-kind"))
            {
                var fails = new List<string>();
                var stub = new KindStub("ov-wrong-kind", ChannelKind.Word);
                if (!PresentationHostAdapter.TryGetPresentationContent(
                        stub,
                        out _,
                        out ToolResult err)
                    && err != null
                    && !err.Success
                    && (err.Error ?? "").IndexOf("ppt", StringComparison.OrdinalIgnoreCase) >= 0
                    && (err.Error ?? "").IndexOf("wpp", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    Finish(run, "ov-neg-wrong-kind", fails);
                }
                else
                {
                    fails.Add(err == null || err.Success
                        ? "文字渠道应失败"
                        : "文案应含 ppt/wpp，实际 " + (err.Error ?? ""));
                    Finish(run, "ov-neg-wrong-kind", fails);
                }
            }
        }

        private static void RunCom(TestRun run, IList<string> filters, bool useWpp)
        {
            Console.WriteLine(useWpp ? "--- B 页目录（WPP）---" : "--- B 页目录 ---");
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
                PptChannel channel = ChannelRegistry.CreateOrGetPpt(pres, "", claimDefaultIfEmpty: false);
                if (!channel.TryGetLivePresentation(out PowerPoint.Presentation live) || live == null)
                {
                    run.CaseFail("ov-", "渠道拿不到活稿");
                    return;
                }

                RunPages(run, filters, new Session(channel, live, wpp: false));
            }
            catch (Exception ex)
            {
                run.CaseFail("ov-", "COM 失败: " + ex.Message);
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
                    run.CaseFail("ov-", startErr);
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
                    run.CaseFail("ov-", "WPP Add 失败");
                    return;
                }

                WppChannel channel = ChannelRegistry.CreateOrGetWpp(pres, "", claimDefaultIfEmpty: false);
                if (!channel.TryGetLivePresentation(out object live) || live == null)
                {
                    run.CaseFail("ov-", "渠道拿不到活稿");
                    return;
                }

                RunPages(run, filters, new Session(channel, live, wpp: true));
            }
            catch (Exception ex)
            {
                run.CaseFail("ov-", "COM 失败: " + ex.Message);
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

        private static void RunPages(TestRun run, IList<string> filters, Session s)
        {
            if (!s.TryEnsureOneSlide(out string err)
                || !s.TryReadComIds(out List<string> one, out err)
                || one.Count < 1)
            {
                FailNeeded(run, filters, err ?? "读不到首页",
                    "ov-count-1", "ov-count-3", "ov-id-stable-after-move",
                    "ov-hidden-listed", "ov-slim-fields", "ov-unsaved-path",
                    "ov-display-table", "ov-no-html");
                return;
            }

            ToolResult r1 = InvokeTool(s.Args());
            string expectKind = s.Wpp ? "wpp" : "ppt";

            if (Want(filters, "ov-count-1"))
            {
                AssertCatalog(run, "ov-count-1", r1, one, expectKind);
            }

            if (Want(filters, "ov-unsaved-path"))
            {
                AssertUnsaved(run, r1);
            }

            if (Want(filters, "ov-slim-fields"))
            {
                AssertSlim(run, r1);
            }

            if (Want(filters, "ov-no-html"))
            {
                AssertNoHtml(run, s, r1);
            }

            bool needThree = Want(filters, "ov-count-3")
                || Want(filters, "ov-id-stable-after-move")
                || Want(filters, "ov-hidden-listed")
                || Want(filters, "ov-display-table");
            if (!needThree)
            {
                return;
            }

            if (!s.TryAddBlank(out err) || !s.TryAddBlank(out err) || !s.TryReadComIds(out List<string> three, out err)
                || three.Count != 3)
            {
                FailNeeded(run, filters, err ?? "加到 3 页失败",
                    "ov-count-3", "ov-id-stable-after-move", "ov-hidden-listed", "ov-display-table");
                return;
            }

            ToolResult r3 = InvokeTool(s.Args());
            if (Want(filters, "ov-count-3"))
            {
                AssertCatalog(run, "ov-count-3", r3, three, expectKind);
            }

            if (Want(filters, "ov-display-table"))
            {
                AssertDisplay(run, r3, three);
            }

            if (Want(filters, "ov-hidden-listed"))
            {
                if (!s.TrySetHidden(2, true, out err))
                {
                    run.CaseFail("ov-hidden-listed", err ?? "隐藏失败");
                }
                else
                {
                    AssertHidden(run, InvokeTool(s.Args()), three[1]);
                }
            }

            if (Want(filters, "ov-id-stable-after-move"))
            {
                if (!s.TryMoveFirstToEnd(out err))
                {
                    run.CaseFail("ov-id-stable-after-move", err ?? "MoveTo 失败");
                }
                else
                {
                    AssertMoved(run, InvokeTool(s.Args()), three);
                }
            }
        }

        private static void AssertCatalog(
            TestRun run,
            string id,
            ToolResult r,
            IList<string> comIds,
            string expectKind)
        {
            var fails = new List<string>();
            if (!IsOk(r, fails))
            {
                Finish(run, id, fails);
                return;
            }

            if (!string.Equals(GetString(r, "kind"), expectKind, StringComparison.OrdinalIgnoreCase))
            {
                fails.Add("kind=" + GetString(r, "kind") + " 期望 " + expectKind);
            }

            int count = GetInt(r, "slide_count");
            if (count != comIds.Count)
            {
                fails.Add("slide_count=" + count + " 期望 " + comIds.Count);
            }

            IList<SlideRow> rows = GetRows(r, fails);
            if (rows.Count != comIds.Count)
            {
                fails.Add("slides 行数 " + rows.Count);
            }
            else
            {
                for (int i = 0; i < comIds.Count; i++)
                {
                    if (rows[i].Index != i + 1)
                    {
                        fails.Add("第 " + (i + 1) + " 行 index=" + rows[i].Index);
                    }

                    if (!string.Equals(rows[i].SlideId, comIds[i], StringComparison.Ordinal))
                    {
                        fails.Add("第 " + (i + 1) + " 行 id=" + rows[i].SlideId + " 期望 " + comIds[i]);
                    }
                }
            }

            Finish(run, id, fails);
        }

        private static void AssertUnsaved(TestRun run, ToolResult r)
        {
            var fails = new List<string>();
            if (!IsOk(r, fails))
            {
                Finish(run, "ov-unsaved-path", fails);
                return;
            }

            if (!string.IsNullOrEmpty(GetString(r, "path")))
            {
                fails.Add("未存盘 path 应空，实际 " + GetString(r, "path"));
            }

            string display = GetString(r, "display_contents");
            if (display.IndexOf("未保存", StringComparison.Ordinal) < 0)
            {
                fails.Add("display_contents 应含「未保存」");
            }

            Finish(run, "ov-unsaved-path", fails);
        }

        private static void AssertSlim(TestRun run, ToolResult r)
        {
            var fails = new List<string>();
            if (!IsOk(r, fails))
            {
                Finish(run, "ov-slim-fields", fails);
                return;
            }

            if (HasKey(r, "palette") || HasKey(r, "palette_sampled"))
            {
                fails.Add("顶层不应有 palette");
            }

            GetRows(r, fails);
            Finish(run, "ov-slim-fields", fails);
        }

        private static void AssertNoHtml(TestRun run, Session s, ToolResult before)
        {
            var fails = new List<string>();
            if (!s.TryApplyPower(out string err))
            {
                run.CaseFail("ov-no-html", err ?? "apply 失败");
                return;
            }

            ToolResult r = InvokeTool(s.Args());
            if (!IsOk(r, fails))
            {
                Finish(run, "ov-no-html", fails);
                return;
            }

            string blob = (GetString(r, "display_contents") ?? "") + "\n" + JsonConvert.SerializeObject(r.Data);
            if (blob.IndexOf("电力", StringComparison.Ordinal) >= 0)
            {
                fails.Add("概览倾倒了正文");
            }

            if (blob.IndexOf("<section", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                fails.Add("概览含 <section>");
            }

            if (blob.IndexOf("sid", StringComparison.OrdinalIgnoreCase) >= 0
                && blob.IndexOf("-s", StringComparison.Ordinal) >= 0)
            {
                fails.Add("概览含 ShapeId");
            }

            Finish(run, "ov-no-html", fails);
        }

        private static void AssertDisplay(TestRun run, ToolResult r, IList<string> comIds)
        {
            var fails = new List<string>();
            if (!IsOk(r, fails))
            {
                Finish(run, "ov-display-table", fails);
                return;
            }

            string d = GetString(r, "display_contents") ?? "";
            if (d.IndexOf("页数：3", StringComparison.Ordinal) < 0
                && d.IndexOf("页数:3", StringComparison.Ordinal) < 0)
            {
                fails.Add("应含「页数：3」");
            }

            if (d.IndexOf("| index | slide_id |", StringComparison.Ordinal) < 0)
            {
                fails.Add("应含 index / slide_id 表头");
            }

            if (d.IndexOf("slide_id", StringComparison.Ordinal) < 0
                || d.IndexOf("主键", StringComparison.Ordinal) < 0)
            {
                fails.Add("应说明主键是 slide_id");
            }

            for (int i = 0; i < comIds.Count; i++)
            {
                if (d.IndexOf(comIds[i], StringComparison.Ordinal) < 0)
                {
                    fails.Add("表中抄不到 " + comIds[i]);
                }
            }

            Finish(run, "ov-display-table", fails);
        }

        private static void AssertHidden(TestRun run, ToolResult r, string hiddenId)
        {
            var fails = new List<string>();
            if (!IsOk(r, fails))
            {
                Finish(run, "ov-hidden-listed", fails);
                return;
            }

            if (GetInt(r, "slide_count") != 3)
            {
                fails.Add("隐藏后仍应 3 页");
            }

            IList<SlideRow> rows = GetRows(r, fails);
            bool found = false;
            for (int i = 0; i < rows.Count; i++)
            {
                if (string.Equals(rows[i].SlideId, hiddenId, StringComparison.Ordinal))
                {
                    found = true;
                    break;
                }
            }

            if (!found)
            {
                fails.Add("隐藏页 " + hiddenId + " 应仍在列表");
            }

            Finish(run, "ov-hidden-listed", fails);
        }

        private static void AssertMoved(TestRun run, ToolResult r, IList<string> original)
        {
            var fails = new List<string>();
            if (!IsOk(r, fails))
            {
                Finish(run, "ov-id-stable-after-move", fails);
                return;
            }

            IList<SlideRow> rows = GetRows(r, fails);
            if (rows.Count != 3)
            {
                fails.Add("挪页后应仍 3 行");
                Finish(run, "ov-id-stable-after-move", fails);
                return;
            }

            SlideRow firstNowLast = null;
            for (int i = 0; i < rows.Count; i++)
            {
                if (string.Equals(rows[i].SlideId, original[0], StringComparison.Ordinal))
                {
                    firstNowLast = rows[i];
                    break;
                }
            }

            if (firstNowLast == null)
            {
                fails.Add("原首页 id 丢了");
            }
            else if (firstNowLast.Index != 3)
            {
                fails.Add("原首页 id 的 index=" + firstNowLast.Index + " 应为 3");
            }

            bool has2 = false;
            bool has3 = false;
            for (int i = 0; i < rows.Count; i++)
            {
                if (string.Equals(rows[i].SlideId, original[1], StringComparison.Ordinal))
                {
                    has2 = true;
                }

                if (string.Equals(rows[i].SlideId, original[2], StringComparison.Ordinal))
                {
                    has3 = true;
                }
            }

            if (!has2 || !has3)
            {
                fails.Add("另两页 id 丢了");
            }

            Finish(run, "ov-id-stable-after-move", fails);
        }

        private static bool IsOk(ToolResult r, List<string> fails)
        {
            if (r == null || !r.Success)
            {
                fails.Add(r == null ? "无回包" : (r.Error ?? "失败"));
                return false;
            }

            return true;
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

        private static ToolResult InvokeTool(Dictionary<string, object> args)
        {
            var tools = new Dictionary<string, Func<Dictionary<string, object>, Task<ToolResult>>>(
                StringComparer.OrdinalIgnoreCase);
            F_GetPresentationContentTool.Register(tools, null);
            if (!tools.TryGetValue("F_get_presentation_content", out Func<Dictionary<string, object>, Task<ToolResult>> fn))
            {
                return null;
            }

            return fn(args ?? new Dictionary<string, object>()).GetAwaiter().GetResult();
        }

        private static bool HasKey(ToolResult r, string key)
        {
            var d = r != null ? r.Data as Dictionary<string, object> : null;
            return d != null && d.ContainsKey(key);
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

        private static IList<SlideRow> GetRows(ToolResult r, List<string> fails)
        {
            var list = new List<SlideRow>();
            var d = r != null ? r.Data as Dictionary<string, object> : null;
            if (d == null || !d.TryGetValue("slides", out object raw) || raw == null)
            {
                fails.Add("无 slides");
                return list;
            }

            var arr = raw as System.Collections.IEnumerable;
            if (arr == null)
            {
                fails.Add("slides 不是列表");
                return list;
            }

            foreach (object item in arr)
            {
                var one = item as Dictionary<string, object>;
                if (one == null)
                {
                    fails.Add("slides 项不是对象");
                    continue;
                }

                foreach (string key in one.Keys)
                {
                    if (!string.Equals(key, "index", StringComparison.OrdinalIgnoreCase)
                        && !string.Equals(key, "slide_id", StringComparison.OrdinalIgnoreCase))
                    {
                        fails.Add("slides 多了键 " + key);
                        break;
                    }
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

                string sid = "";
                if (one.TryGetValue("slide_id", out object sv) && sv != null)
                {
                    sid = Convert.ToString(sv) ?? "";
                }

                list.Add(new SlideRow { Index = index, SlideId = sid });
            }

            return list;
        }

        private sealed class SlideRow
        {
            public int Index;
            public string SlideId;
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

        private sealed class Session
        {
            public Session(IOperationChannel channel, object pres, bool wpp)
            {
                Channel = channel;
                Pres = pres;
                Wpp = wpp;
            }

            public IOperationChannel Channel { get; private set; }

            public object Pres { get; private set; }

            public bool Wpp { get; private set; }

            public Dictionary<string, object> Args()
            {
                return new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase)
                {
                    ["channel_id"] = Channel.ChannelId
                };
            }

            public bool TryEnsureOneSlide(out string error)
            {
                error = null;
                try
                {
                    if (Wpp)
                    {
                        object slides = WppCom.GetProperty(Pres, "Slides");
                        int n = Convert.ToInt32(WppCom.GetProperty(slides, "Count") ?? 0);
                        if (n >= 1)
                        {
                            return true;
                        }

                        return TryAddBlank(out error);
                    }

                    PowerPoint.Presentation ppt = (PowerPoint.Presentation)Pres;
                    if (ppt.Slides.Count >= 1)
                    {
                        return true;
                    }

                    return TryAddBlank(out error);
                }
                catch (Exception ex)
                {
                    error = ex.Message;
                    return false;
                }
            }

            public bool TryReadComIds(out List<string> ids, out string error)
            {
                ids = new List<string>();
                error = null;
                try
                {
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

                        return ids.Count > 0;
                    }

                    PowerPoint.Presentation ppt = (PowerPoint.Presentation)Pres;
                    for (int i = 1; i <= ppt.Slides.Count; i++)
                    {
                        ids.Add(ppt.Slides[i].SlideID.ToString(CultureInfo.InvariantCulture));
                    }

                    return ids.Count > 0;
                }
                catch (Exception ex)
                {
                    error = ex.Message;
                    return false;
                }
            }

            public bool TryAddBlank(out string error)
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
                            WppCom.Invoke(slides, "Add", count + 1, 12);
                        }
                        catch
                        {
                            WppCom.Invoke(slides, "Add", count + 1, 1);
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

            public bool TrySetHidden(int index1, bool hidden, out string error)
            {
                error = null;
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

                    PowerPoint.Slide slidePpt = ((PowerPoint.Presentation)Pres).Slides[index1];
                    slidePpt.SlideShowTransition.Hidden = hidden
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

            public bool TryMoveFirstToEnd(out string error)
            {
                error = null;
                try
                {
                    if (Wpp)
                    {
                        object slides = WppCom.GetProperty(Pres, "Slides");
                        int n = Convert.ToInt32(WppCom.GetProperty(slides, "Count") ?? 0);
                        object first = WppCom.GetIndexed(slides, 1);
                        try
                        {
                            WppCom.Invoke(first, "MoveTo", n + 1);
                        }
                        catch
                        {
                            WppCom.Invoke(first, "MoveTo", n);
                        }

                        return true;
                    }

                    PowerPoint.Presentation ppt = (PowerPoint.Presentation)Pres;
                    ppt.Slides[1].MoveTo(ppt.Slides.Count);
                    return true;
                }
                catch (Exception ex)
                {
                    error = ex.Message;
                    return false;
                }
            }

            public bool TryApplyPower(out string error)
            {
                string html = ContentHtml.TextboxCreate("电力", 8, 10, 30, 12);
                string slideId;
                if (Wpp)
                {
                    object slides = WppCom.GetProperty(Pres, "Slides");
                    object slide = WppCom.GetIndexed(slides, 1);
                    slideId = Convert.ToString(
                        WppCom.GetProperty(slide, "SlideID") ?? WppCom.GetProperty(slide, "SlideId")) ?? "";
                }
                else
                {
                    slideId = ((PowerPoint.Presentation)Pres).Slides[1]
                        .SlideID.ToString(CultureInfo.InvariantCulture);
                }

                if (!ContentHtml.TryParse(html, slideId, out PptHtmlApplyPlan plan, out error))
                {
                    return false;
                }

                plan.AllowCreate = true;
                return Wpp
                    ? PptHtmlWppApplier.TryApply(Pres, plan, "ov-wpp", out _, out error)
                    : PptHtmlPowerPointApplier.TryApply(
                        (PowerPoint.Presentation)Pres, plan, "ov-ppt", out _, out error);
            }
        }
    }
}
