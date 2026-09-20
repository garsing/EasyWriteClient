using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using WordAddIn1.OpenFiles;
using WordAddIn1.PresentationHost;

namespace PptHtmlContentRoundtripTest
{
    /// <summary>WPS 演示（Kwpp）正式往返：与 ContentBatchRunner 同目录，走 Wpp Applier/Reader。</summary>
    internal static class ContentBatchRunnerWpp
    {
        private const int RecycleEvery = 6;
        private const int PpLayoutBlank = 12;

        public static void Run(TestRun run, int? onlyBatch, string outDir, IList<string> nameFilters)
        {
            Directory.CreateDirectory(outDir);
            ContentAssets assets = ContentAssetsIo.Ensure(Path.Combine(outDir, "assets"));
            List<ContentCase> all = ContentCatalog.Build();
            if (nameFilters != null && nameFilters.Count > 0)
            {
                Console.WriteLine("过滤用例名含: " + string.Join(", ", nameFilters));
            }

            Console.WriteLine("宿主: WPP (Kwpp.Application 晚绑定)");

            for (int batch = 1; batch <= ContentCatalog.BatchCount; batch++)
            {
                if (onlyBatch.HasValue && onlyBatch.Value != batch)
                {
                    continue;
                }

                if (!BatchHasCases(all, batch, nameFilters))
                {
                    continue;
                }

                RunBatch(run, all, batch, outDir, assets, nameFilters);
            }
        }

        private static bool BatchHasCases(List<ContentCase> all, int batch, IList<string> nameFilters)
        {
            for (int i = 0; i < all.Count; i++)
            {
                if (all[i].Batch == batch && NameMatches(all[i].Name, nameFilters))
                {
                    return true;
                }
            }

            return false;
        }

        private static bool NameMatches(string name, IList<string> nameFilters)
        {
            if (nameFilters == null || nameFilters.Count == 0)
            {
                return true;
            }

            string n = name ?? "";
            for (int i = 0; i < nameFilters.Count; i++)
            {
                string f = nameFilters[i];
                if (!string.IsNullOrEmpty(f)
                    && n.IndexOf(f, StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    return true;
                }
            }

            return false;
        }

        private static void RunBatch(
            TestRun run,
            List<ContentCase> all,
            int batch,
            string outDir,
            ContentAssets assets,
            IList<string> nameFilters)
        {
            string path = Path.Combine(outDir, "content-batch-wpp-" + batch + ".pptx");
            Console.WriteLine();
            Console.WriteLine("--- WPP 内容往返 批次 " + batch + " → " + path + " ---");

            var cases = new List<ContentCase>();
            for (int i = 0; i < all.Count; i++)
            {
                if (all[i].Batch == batch && NameMatches(all[i].Name, nameFilters))
                {
                    cases.Add(all[i]);
                }
            }

            if (File.Exists(path))
            {
                try
                {
                    File.Delete(path);
                }
                catch (Exception ex)
                {
                    Console.WriteLine("  无法覆盖旧文件: " + ex.Message);
                }
            }

            object app = null;
            object pres = null;
            int sinceRecycle = 0;
            try
            {
                if (!TryStartApp(out app, out string startError))
                {
                    run.CaseFail("batch-wpp-" + batch + "-open", startError);
                    return;
                }

                if (!TryNewDeck(app, out pres, out startError))
                {
                    run.CaseFail("batch-wpp-" + batch + "-open", startError);
                    return;
                }

                for (int i = 0; i < cases.Count; i++)
                {
                    ContentCase one = cases[i];
                    string tag = "WppB" + batch + "P" + one.Page.ToString("00", CultureInfo.InvariantCulture)
                        + " #" + one.Index + " " + one.Name;

                    if (!IsAppAlive(app) || pres == null)
                    {
                        Console.WriteLine("  WPP 已断开，重启后续跑…");
                        if (!TryRecycle(ref app, ref pres, path, out string recErr))
                        {
                            run.CaseSkip(tag, "重启 WPP 失败: " + recErr);
                            SkipRest(run, cases, i + 1, "WPP 无法重启");
                            break;
                        }

                        sinceRecycle = 0;
                    }
                    else if (sinceRecycle >= RecycleEvery)
                    {
                        SaveQuiet(pres, path);
                        if (!TryRecycle(ref app, ref pres, path, out string recErr))
                        {
                            run.CaseSkip(tag, "定期重启失败: " + recErr);
                            SkipRest(run, cases, i + 1, "WPP 无法重启");
                            break;
                        }

                        sinceRecycle = 0;
                    }

                    bool retry = false;
                    try
                    {
                        object slide = AddBlankSlide(pres);
                        TryAddCaseLabel(slide, tag);
                        string rpc = RunOneSlide(run, pres, slide, one, assets, outDir, tag);
                        if (rpc != null)
                        {
                            retry = true;
                        }
                        else
                        {
                            SaveQuiet(pres, path);
                            sinceRecycle++;
                            int slideCount = Convert.ToInt32(
                                WppCom.GetProperty(WppCom.GetProperty(pres, "Slides"), "Count") ?? 0);
                            Console.WriteLine("  … " + (i + 1) + "/" + cases.Count
                                + " 页=" + slideCount);
                        }
                    }
                    catch (Exception ex)
                    {
                        if (IsRpcText(ex.Message))
                        {
                            retry = true;
                        }
                        else
                        {
                            run.CaseFail(tag, ex.GetType().Name + ": " + ex.Message);
                            sinceRecycle++;
                        }
                    }

                    if (!retry)
                    {
                        continue;
                    }

                    Console.WriteLine("  COM 断开，重启后重试 " + tag);
                    if (!TryRecycle(ref app, ref pres, path, out string recErr2))
                    {
                        run.CaseFail(tag, "WPP 断开且重启失败: " + recErr2);
                        SkipRest(run, cases, i + 1, "WPP 无法重启");
                        break;
                    }

                    sinceRecycle = 0;
                    try
                    {
                        object slide = AddBlankSlide(pres);
                        TryAddCaseLabel(slide, tag);
                        string rpc = RunOneSlide(run, pres, slide, one, assets, outDir, tag);
                        if (rpc != null)
                        {
                            run.CaseFail(tag, "重启后仍断开: " + rpc);
                            SkipRest(run, cases, i + 1, "WPP 再次断开");
                            break;
                        }

                        SaveQuiet(pres, path);
                        sinceRecycle++;
                        Console.WriteLine("  … " + (i + 1) + "/" + cases.Count
                            + " （重启后重试成功）");
                    }
                    catch (Exception ex)
                    {
                        run.CaseFail(tag, "重启后仍失败: " + ex.Message);
                        SkipRest(run, cases, i + 1, "WPP 再次失败");
                        break;
                    }
                }

                SaveQuiet(pres, path);
                if (File.Exists(path))
                {
                    Console.WriteLine("  已保存 " + path);
                }
            }
            finally
            {
                CloseQuiet(pres);
                SoftQuit(ref app);
            }
        }

        private static string RunOneSlide(
            TestRun run,
            object presentation,
            object slide,
            ContentCase one,
            ContentAssets assets,
            string outDir,
            string tag)
        {
            string slideId;
            try
            {
                object idObj = WppCom.GetProperty(slide, "SlideID");
                if (idObj == null)
                {
                    idObj = WppCom.GetProperty(slide, "SlideId");
                }

                slideId = Convert.ToInt32(idObj).ToString(CultureInfo.InvariantCulture);
            }
            catch (Exception ex)
            {
                if (IsRpcText(ex.Message))
                {
                    return ex.Message;
                }

                run.CaseFail(tag, "取 SlideID 失败: " + ex.Message);
                return null;
            }

            string createHtml;
            try
            {
                createHtml = one.CreateHtml(assets);
            }
            catch (Exception ex)
            {
                run.CaseFail(tag, "生成新建稿失败: " + ex.Message);
                return null;
            }

            if (!ContentHtml.TryParse(createHtml, slideId, out PptHtmlApplyPlan createPlan, out string error))
            {
                if (MatchesExpectError(one.ExpectApplyErrorContains, error))
                {
                    run.CaseOk(tag);
                    return null;
                }

                run.CaseFail(tag, "新建稿解析失败: " + error);
                return null;
            }

            ContentHtml.ResolvePicturePaths(createPlan);
            createPlan.AllowCreate = true;

            if (!PptHtmlWppApplier.TryApply(
                    presentation,
                    createPlan,
                    "content-test-wpp",
                    out PptHtmlApplyResult createResult,
                    out error))
            {
                if (IsRpcText(error))
                {
                    return error;
                }

                if (MatchesExpectError(one.ExpectApplyErrorContains, error))
                {
                    run.CaseOk(tag);
                    return null;
                }

                run.CaseFail(tag, "新建 apply 失败: " + error + FormatWarnings(createResult));
                return null;
            }

            if (!string.IsNullOrEmpty(one.ExpectApplyErrorContains) && one.ReplaceHtml == null)
            {
                run.CaseFail(tag, "期望 apply 失败含「" + one.ExpectApplyErrorContains + "」却成功了");
                return null;
            }

            List<string> createdIds = CollectCreatedIds(createResult);
            if (createdIds.Count == 0 && one.ReplaceHtml != null)
            {
                run.CaseFail(tag, "新建未返回 shape_id，无法换数");
                return null;
            }

            if (!TryReadPage(presentation, slideId, one.ContentOnly, out PptHtmlReadResult afterCreate, out error))
            {
                if (IsRpcText(error))
                {
                    return error;
                }

                run.CaseFail(tag, "新建后读回失败: " + error);
                return null;
            }

            // 换图前先导出 data-src：替换后旧 COM Id 可能已失效，事后再导出会漏 before
            string exportRoot = Path.Combine(outDir, "export-wpp", one.Name);
            ContentAssert.AttachPictureSrc(
                presentation, afterCreate, Path.Combine(exportRoot, "after-create"));

            PptHtmlReadResult afterReplace = null;
            PptHtmlApplyResult replaceResult = null;
            if (one.ReplaceHtml != null)
            {
                string replaceHtml;
                try
                {
                    replaceHtml = one.ReplaceHtml(assets, createdIds);
                }
                catch (Exception ex)
                {
                    run.CaseFail(tag, "生成换数稿失败: " + ex.Message);
                    return null;
                }

                if (!ContentHtml.TryParse(replaceHtml, slideId, out PptHtmlApplyPlan replacePlan, out error))
                {
                    if (MatchesExpectError(one.ExpectApplyErrorContains, error))
                    {
                        run.CaseOk(tag);
                        return null;
                    }

                    run.CaseFail(tag, "换数稿解析失败: " + error);
                    return null;
                }

                ContentHtml.ResolvePicturePaths(replacePlan);
                replacePlan.AllowCreate = false;
                if (!PptHtmlWppApplier.TryApply(
                        presentation,
                        replacePlan,
                        "content-test-wpp",
                        out replaceResult,
                        out error))
                {
                    if (IsRpcText(error))
                    {
                        return error;
                    }

                    if (MatchesExpectError(one.ExpectApplyErrorContains, error))
                    {
                        run.CaseOk(tag);
                        return null;
                    }

                    run.CaseFail(tag, "换数 apply 失败: " + error + FormatWarnings(replaceResult));
                    return null;
                }

                if (!string.IsNullOrEmpty(one.ExpectApplyErrorContains))
                {
                    run.CaseFail(tag, "期望换数失败含「" + one.ExpectApplyErrorContains + "」却成功了");
                    return null;
                }

                if (!TryReadPage(presentation, slideId, one.ContentOnly, out afterReplace, out error))
                {
                    if (IsRpcText(error))
                    {
                        return error;
                    }

                    run.CaseFail(tag, "换数后读回失败: " + error);
                    return null;
                }
            }

            if (one.Match == null)
            {
                run.CaseOk(tag);
                return null;
            }

            var ctx = new ContentAssertContext
            {
                Case = one,
                CreatedShapeIds = createdIds,
                AfterCreate = afterCreate,
                AfterReplace = afterReplace,
                CreateResult = createResult,
                ReplaceResult = replaceResult,
                ExportDir = Path.Combine(outDir, "export-wpp", one.Name),
                Presentation = presentation
            };

            string mismatch;
            try
            {
                mismatch = one.Match(ctx);
            }
            catch (Exception ex)
            {
                run.CaseFail(tag, "断言异常: " + ex.Message);
                return null;
            }

            if (mismatch == null)
            {
                run.CaseOk(tag);
            }
            else
            {
                run.CaseFail(tag, mismatch);
            }

            return null;
        }

        private static bool TryReadPage(
            object presentation,
            string slideId,
            bool contentOnly,
            out PptHtmlReadResult result,
            out string error)
        {
            return PptHtmlWppReader.TryRead(
                presentation,
                slideId,
                "content-test-wpp",
                "wpp",
                null,
                true,
                false,
                false,
                contentOnly,
                PptConventionHtml.DefaultContentMinArea,
                out result,
                out error);
        }

        private static List<string> CollectCreatedIds(PptHtmlApplyResult result)
        {
            var ids = new List<string>();
            if (result?.CreatedShapes == null)
            {
                return ids;
            }

            for (int i = 0; i < result.CreatedShapes.Count; i++)
            {
                Dictionary<string, object> d = result.CreatedShapes[i];
                if (d != null && d.TryGetValue("shape_id", out object v) && v != null)
                {
                    string id = v.ToString();
                    if (!string.IsNullOrEmpty(id))
                    {
                        ids.Add(id);
                    }
                }
            }

            return ids;
        }

        private static string FormatWarnings(PptHtmlApplyResult result)
        {
            if (result?.Warnings == null || result.Warnings.Count == 0)
            {
                return "";
            }

            return " warnings=[" + string.Join("; ", result.Warnings) + "]";
        }

        private static void SkipRest(TestRun run, List<ContentCase> cases, int from, string reason)
        {
            for (int i = from; i < cases.Count; i++)
            {
                string tag = "#" + cases[i].Index + " " + cases[i].Name;
                run.CaseSkip(tag, reason);
            }
        }

        private static bool MatchesExpectError(string expectContains, string error)
        {
            return !string.IsNullOrEmpty(expectContains)
                && !string.IsNullOrEmpty(error)
                && error.IndexOf(expectContains, StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private static bool IsRpcText(string message)
        {
            if (string.IsNullOrEmpty(message))
            {
                return false;
            }

            return message.IndexOf("RPC", StringComparison.OrdinalIgnoreCase) >= 0
                || message.IndexOf("0x800706BA", StringComparison.OrdinalIgnoreCase) >= 0
                || message.IndexOf("被呼叫方", StringComparison.OrdinalIgnoreCase) >= 0
                || message.IndexOf("服务器忙", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private static object AddBlankSlide(object presentation)
        {
            object slides = WppCom.GetProperty(presentation, "Slides");
            int count = Convert.ToInt32(WppCom.GetProperty(slides, "Count") ?? 0);
            try
            {
                return WppCom.Invoke(slides, "Add", count + 1, PpLayoutBlank);
            }
            catch
            {
                return WppCom.Invoke(slides, "Add", count + 1, 1);
            }
        }

        private static void TryAddCaseLabel(object slide, string tag)
        {
            try
            {
                object shapes = WppCom.GetProperty(slide, "Shapes");
                // AddTextbox(Orientation, Left, Top, Width, Height)
                object box = WppCom.Invoke(shapes, "AddTextbox", 1, 6.0, 2.0, 700.0, 18.0);
                object tf = WppCom.GetProperty(box, "TextFrame");
                object tr = WppCom.GetProperty(tf, "TextRange");
                WppCom.TrySetProperty(tr, "Text", tag);
            }
            catch
            {
            }
        }

        private static bool TryStartApp(out object app, out string error)
        {
            return WppCom.TryCreateApplication(out app, out error);
        }

        private static bool TryNewDeck(object app, out object presentation, out string error)
        {
            presentation = null;
            error = null;
            try
            {
                object presentations = WppCom.GetProperty(app, "Presentations");
                try
                {
                    presentation = WppCom.Invoke(presentations, "Add", true);
                }
                catch
                {
                    presentation = WppCom.Invoke(presentations, "Add");
                }

                if (presentation == null)
                {
                    error = "Presentations.Add 无返回";
                    return false;
                }

                ClearToBlankDeck(presentation);
                return true;
            }
            catch (Exception ex)
            {
                error = ex.Message;
                return false;
            }
        }

        private static void ClearToBlankDeck(object presentation)
        {
            try
            {
                object slides = WppCom.GetProperty(presentation, "Slides");
                int count = Convert.ToInt32(WppCom.GetProperty(slides, "Count") ?? 0);
                if (count == 0)
                {
                    return;
                }

                while (Convert.ToInt32(WppCom.GetProperty(slides, "Count") ?? 0) > 0)
                {
                    object first = WppCom.GetIndexed(slides, 1);
                    WppCom.Invoke(first, "Delete");
                }
            }
            catch
            {
            }
        }

        private static bool TryRecycle(
            ref object app,
            ref object presentation,
            string path,
            out string error)
        {
            error = null;
            SaveQuiet(presentation, path);
            CloseQuiet(presentation);
            presentation = null;
            SoftQuit(ref app);
            GC.Collect();
            GC.WaitForPendingFinalizers();
            Thread.Sleep(900);

            if (!TryStartApp(out app, out error))
            {
                return false;
            }

            try
            {
                object presentations = WppCom.GetProperty(app, "Presentations");
                if (File.Exists(path))
                {
                    try
                    {
                        presentation = WppCom.Invoke(presentations, "Open", path, false, false, true);
                    }
                    catch
                    {
                        presentation = WppCom.Invoke(presentations, "Open", path);
                    }
                }
                else if (!TryNewDeck(app, out presentation, out error))
                {
                    return false;
                }

                return presentation != null;
            }
            catch (Exception ex)
            {
                error = ex.Message;
                return false;
            }
        }

        private static void SaveQuiet(object presentation, string path)
        {
            if (presentation == null || string.IsNullOrEmpty(path))
            {
                return;
            }

            try
            {
                WppCom.Invoke(presentation, "SaveAs", path);
            }
            catch
            {
                try
                {
                    WppCom.Invoke(presentation, "SaveAs", path, 24);
                }
                catch
                {
                }
            }
        }

        private static void CloseQuiet(object presentation)
        {
            if (presentation == null)
            {
                return;
            }

            try
            {
                WppCom.Invoke(presentation, "Close");
            }
            catch
            {
            }
        }

        private static void SoftQuit(ref object app)
        {
            if (app == null)
            {
                return;
            }

            try
            {
                WppCom.Invoke(app, "Quit");
            }
            catch
            {
            }

            try
            {
                Marshal.FinalReleaseComObject(app);
            }
            catch
            {
            }

            app = null;
        }

        private static bool IsAppAlive(object app)
        {
            if (app == null)
            {
                return false;
            }

            try
            {
                return WppCom.GetProperty(app, "Name") != null;
            }
            catch
            {
                return false;
            }
        }
    }
}
