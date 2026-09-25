using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using WordAddIn1.PresentationHost;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PptHtmlContentRoundtripTest
{
    /// <summary>一页一场景；走完整 Applier → Reader 管道。</summary>
    internal static class ContentBatchRunner
    {
        private const int RecycleEveryDefault = 6;

        public static void Run(
            TestRun run,
            int? onlyBatch,
            string outDir,
            IList<string> nameFilters,
            bool failFast = false,
            int recycleEvery = RecycleEveryDefault)
        {
            Directory.CreateDirectory(outDir);
            ContentAssets assets = ContentAssetsIo.Ensure(Path.Combine(outDir, "assets"));
            List<ContentCase> all = ContentCatalog.Build();
            if (nameFilters != null && nameFilters.Count > 0)
            {
                Console.WriteLine("过滤用例名含: " + string.Join(", ", nameFilters));
            }

            if (failFast)
            {
                Console.WriteLine("失败即停");
            }

            if (recycleEvery <= 0)
            {
                recycleEvery = int.MaxValue;
                Console.WriteLine("不定期 Quit 回收");
            }
            else
            {
                Console.WriteLine("每 " + recycleEvery + " 条 Quit + Kill POWERPNT 后重启");
            }

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

                RunBatch(run, all, batch, outDir, assets, nameFilters, failFast, recycleEvery);
                if (failFast && run.CasesFailed > 0)
                {
                    break;
                }
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
            IList<string> nameFilters,
            bool failFast,
            int recycleEvery)
        {
            string path = Path.Combine(outDir, "content-batch-" + batch + ".pptx");
            Console.WriteLine();
            Console.WriteLine("--- 内容往返 批次 " + batch + " → " + path + " ---");

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

            PowerPoint.Application app = null;
            PowerPoint.Presentation pres = null;
            int sinceRecycle = 0;
            try
            {
                if (!TryStartApp(out app, out string startError))
                {
                    run.CaseFail("batch-" + batch + "-open", startError);
                    return;
                }

                if (!TryNewDeck(app, out pres, out startError))
                {
                    run.CaseFail("batch-" + batch + "-open", startError);
                    return;
                }

                for (int i = 0; i < cases.Count; i++)
                {
                    ContentCase one = cases[i];
                    string tag = "B" + batch + "P" + one.Page.ToString("00", CultureInfo.InvariantCulture)
                        + " #" + one.Index + " " + one.Name;

                    if (!IsAppAlive(app) || pres == null)
                    {
                        Console.WriteLine("  PowerPoint 已断开，重启后续跑…");
                        if (!TryRecycle(ref app, ref pres, path, out string recErr))
                        {
                            run.CaseSkip(tag, "重启 PowerPoint 失败: " + recErr);
                            SkipRest(run, cases, i + 1, "PowerPoint 无法重启");
                            break;
                        }

                        sinceRecycle = 0;
                    }
                    else if (sinceRecycle >= recycleEvery)
                    {
                        SaveQuiet(pres, path);
                        if (!TryRecycle(ref app, ref pres, path, out string recErr))
                        {
                            run.CaseSkip(tag, "定期重启失败: " + recErr);
                            SkipRest(run, cases, i + 1, "PowerPoint 无法重启");
                            break;
                        }

                        sinceRecycle = 0;
                    }

                    bool retry = false;
                    bool ran = false;
                    try
                    {
                        PowerPoint.Slide slide = pres.Slides.Add(
                            pres.Slides.Count + 1,
                            PowerPoint.PpSlideLayout.ppLayoutBlank);
                        TryAddCaseLabel(slide, tag);
                        string rpc = RunOneSlide(run, app, pres, slide, one, assets, outDir, tag);
                        ran = true;
                        if (rpc != null)
                        {
                            retry = true;
                        }
                        else
                        {
                            SaveQuiet(pres, path);
                            sinceRecycle++;
                            int pages = TrySlideCount(pres);
                            Console.WriteLine("  … " + (i + 1) + "/" + cases.Count
                                + " 页=" + (pages >= 0 ? pages.ToString(CultureInfo.InvariantCulture) : "?")
                                + " " + DescribePptSession(app));
                            if (pages < 0)
                            {
                                CloseQuiet(pres);
                                pres = null;
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        if (IsRpcText(ex.Message))
                        {
                            retry = true;
                        }
                        else if (IsDeadPres(ex.Message))
                        {
                            if (!ran)
                            {
                                retry = true;
                            }
                            else
                            {
                                CloseQuiet(pres);
                                pres = null;
                            }
                        }
                        else
                        {
                            run.CaseFail(tag, ex.GetType().Name + ": " + ex.Message);
                            sinceRecycle++;
                        }
                    }

                    if (failFast && run.CasesFailed > 0)
                    {
                        Console.WriteLine("  失败即停，余下 " + (cases.Count - i - 1) + " 条不跑");
                        break;
                    }

                    if (!retry)
                    {
                        continue;
                    }

                    Console.WriteLine("  COM 断开，重启后重试 " + tag);
                    if (!TryRecycle(ref app, ref pres, path, out string recErr2))
                    {
                        run.CaseFail(tag, "PowerPoint 断开且重启失败: " + recErr2);
                        SkipRest(run, cases, i + 1, "PowerPoint 无法重启");
                        break;
                    }

                    sinceRecycle = 0;
                    try
                    {
                        PowerPoint.Slide slide = pres.Slides.Add(
                            pres.Slides.Count + 1,
                            PowerPoint.PpSlideLayout.ppLayoutBlank);
                        TryAddCaseLabel(slide, tag);
                        string rpc = RunOneSlide(run, app, pres, slide, one, assets, outDir, tag);
                        if (rpc != null)
                        {
                            run.CaseFail(tag, "重启后仍断开: " + rpc);
                            SkipRest(run, cases, i + 1, "PowerPoint 再次断开");
                            break;
                        }

                        SaveQuiet(pres, path);
                        sinceRecycle++;
                        Console.WriteLine("  … " + (i + 1) + "/" + cases.Count
                            + " 页=" + pres.Slides.Count + "（重启后重试成功）");
                    }
                    catch (Exception ex)
                    {
                        run.CaseFail(tag, "重启后仍失败: " + ex.Message);
                        SkipRest(run, cases, i + 1, "PowerPoint 再次失败");
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
                ThoroughClose(ref app);
            }
        }

        /// <returns>RPC 死信，否则 null。</returns>
        private static string RunOneSlide(
            TestRun run,
            PowerPoint.Application app,
            PowerPoint.Presentation pres,
            PowerPoint.Slide slide,
            ContentCase one,
            ContentAssets assets,
            string outDir,
            string tag)
        {
            string slideId;
            try
            {
                slideId = slide.SlideID.ToString(CultureInfo.InvariantCulture);
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

            if (one.GroupRun != null)
            {
                return ContentGroupHarness.Run(run, pres, slide, slideId, one, tag, wpp: false, assets: assets);
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

            if (!PptHtmlPowerPointApplier.TryApply(
                    pres,
                    createPlan,
                    "content-test",
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

            if (!TryReadPage(pres, slideId, one.ContentOnly, out PptHtmlReadResult afterCreate, out error))
            {
                if (IsRpcText(error))
                {
                    return error;
                }

                run.CaseFail(tag, "新建后读回失败: " + error + " " + DescribePptSession(app));
                return null;
            }

            // 换图前先导出 data-src：替换后旧 COM Id 可能已失效，事后再导出会漏 before
            string exportRoot = Path.Combine(outDir, "export", one.Name);
            ContentAssert.AttachPictureSrc(pres, afterCreate, Path.Combine(exportRoot, "after-create"));

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
                if (!PptHtmlPowerPointApplier.TryApply(
                        pres,
                        replacePlan,
                        "content-test",
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

                if (!TryReadPage(pres, slideId, one.ContentOnly, out afterReplace, out error))
                {
                    if (IsRpcText(error))
                    {
                        return error;
                    }

                    run.CaseFail(tag, "换数后读回失败: " + error + " " + DescribePptSession(app));
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
                ExportDir = Path.Combine(outDir, "export", one.Name),
                Presentation = pres
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

        private static bool MatchesExpectError(string expectContains, string error)
        {
            return !string.IsNullOrEmpty(expectContains)
                && !string.IsNullOrEmpty(error)
                && error.IndexOf(expectContains, StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private static bool TryReadPage(
            PowerPoint.Presentation presentation,
            string slideId,
            bool contentOnly,
            out PptHtmlReadResult result,
            out string error)
        {
            return PptHtmlPowerPointReader.TryRead(
                presentation,
                slideId,
                "content-test",
                "ppt",
                null,
                true,
                false,
                false,
                contentOnly,
                0,
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

            List<string> w = result.Warnings;
            int take = w.Count < 24 ? w.Count : 24;
            var tail = new List<string>();
            for (int i = w.Count - take; i < w.Count; i++)
            {
                tail.Add(w[i]);
            }

            return " warnings尾" + take + "/" + w.Count + "=[" + string.Join(" | ", tail) + "]";
        }

        private static void SkipRest(TestRun run, List<ContentCase> cases, int from, string reason)
        {
            for (int i = from; i < cases.Count; i++)
            {
                string tag = "#" + cases[i].Index + " " + cases[i].Name;
                run.CaseSkip(tag, reason);
            }
        }

        private static void TryAddCaseLabel(PowerPoint.Slide slide, string tag)
        {
            try
            {
                PowerPoint.Shape box = slide.Shapes.AddTextbox(
                    Office.MsoTextOrientation.msoTextOrientationHorizontal,
                    6f,
                    2f,
                    700f,
                    18f);
                box.TextFrame.TextRange.Text = tag;
                box.TextFrame.TextRange.Font.Size = 9f;
                box.TextFrame.TextRange.Font.Color.RGB = 0x666666;
            }
            catch
            {
            }
        }

        private static bool TryStartApp(out PowerPoint.Application app, out string error)
        {
            app = null;
            error = null;
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

                return true;
            }
            catch (Exception ex)
            {
                error = "无法启动 PowerPoint：" + ex.Message;
                return false;
            }
        }

        private static bool TryNewDeck(PowerPoint.Application app, out PowerPoint.Presentation pres, out string error)
        {
            pres = null;
            error = null;
            try
            {
                pres = app.Presentations.Add(Office.MsoTriState.msoTrue);
                ClearToBlankDeck(pres);
                return true;
            }
            catch (Exception ex)
            {
                error = ex.Message;
                return false;
            }
        }

        private static void ClearToBlankDeck(PowerPoint.Presentation pres)
        {
            if (pres.Slides.Count == 0)
            {
                return;
            }

            try
            {
                pres.Slides.Add(pres.Slides.Count + 1, PowerPoint.PpSlideLayout.ppLayoutBlank);
                while (pres.Slides.Count > 1)
                {
                    pres.Slides[1].Delete();
                }

                pres.Slides[1].Delete();
            }
            catch
            {
            }
        }

        private static bool TryRecycle(
            ref PowerPoint.Application app,
            ref PowerPoint.Presentation pres,
            string path,
            out string error)
        {
            error = null;
            SaveQuiet(pres, path);
            CloseQuiet(pres);
            pres = null;
            int[] oldPids = PowerpntPids();
            Console.WriteLine("  回收：Quit + Kill POWERPNT 后重启 oldPid="
                + string.Join(",", oldPids ?? new int[0]));
            ThoroughClose(ref app);
            if (!WaitPowerpntGone(oldPids, 15000))
            {
                error = "Quit 后 POWERPNT 未退尽 pid=" + string.Join(",", oldPids ?? new int[0]);
                return false;
            }

            if (!TryStartApp(out app, out error))
            {
                return false;
            }

            Console.WriteLine("  回收完成 newPid=" + string.Join(",", PowerpntPids()));

            try
            {
                if (File.Exists(path))
                {
                    pres = app.Presentations.Open(
                        path,
                        Office.MsoTriState.msoFalse,
                        Office.MsoTriState.msoFalse,
                        Office.MsoTriState.msoTrue);
                }
                else if (!TryNewDeck(app, out pres, out error))
                {
                    return false;
                }

                return true;
            }
            catch (Exception ex)
            {
                error = ex.Message;
                return false;
            }
        }

        private static void SaveQuiet(PowerPoint.Presentation pres, string path)
        {
            if (pres == null || string.IsNullOrEmpty(path))
            {
                return;
            }

            try
            {
                string full = null;
                try
                {
                    full = pres.FullName;
                }
                catch
                {
                }

                if (string.IsNullOrEmpty(full) || !File.Exists(path))
                {
                    pres.SaveAs(path, PowerPoint.PpSaveAsFileType.ppSaveAsOpenXMLPresentation);
                }
                else
                {
                    pres.Save();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("  保存失败: " + ex.Message);
            }
        }

        private static void CloseQuiet(PowerPoint.Presentation pres)
        {
            if (pres == null)
            {
                return;
            }

            try
            {
                pres.Saved = Office.MsoTriState.msoTrue;
                pres.Close();
            }
            catch
            {
            }

            TryRelease(pres);
        }

        private static void ThoroughClose(ref PowerPoint.Application app)
        {
            int[] pids = PowerpntPids();
            SoftQuit(ref app);
            GC.Collect();
            GC.WaitForPendingFinalizers();
            KillPowerpnt(pids);
        }

        private static void SoftQuit(ref PowerPoint.Application app)
        {
            if (app == null)
            {
                return;
            }

            try
            {
                if (IsAppAlive(app))
                {
                    app.Quit();
                }
            }
            catch
            {
            }

            TryRelease(app);
            app = null;
        }

        private static int[] PowerpntPids()
        {
            try
            {
                Process[] ps = Process.GetProcessesByName("POWERPNT");
                var ids = new int[ps.Length];
                for (int i = 0; i < ps.Length; i++)
                {
                    ids[i] = ps[i].Id;
                }

                return ids;
            }
            catch
            {
                return new int[0];
            }
        }

        private static void KillPowerpnt(int[] pids)
        {
            if (pids == null)
            {
                return;
            }

            for (int i = 0; i < pids.Length; i++)
            {
                try
                {
                    Process.GetProcessById(pids[i]).Kill();
                    Console.WriteLine("  回收已 Kill POWERPNT pid=" + pids[i]);
                }
                catch
                {
                }
            }
        }

        private static bool WaitPowerpntGone(int[] pids, int timeoutMs)
        {
            if (pids == null || pids.Length == 0)
            {
                return true;
            }

            var sw = Stopwatch.StartNew();
            while (sw.ElapsedMilliseconds < timeoutMs)
            {
                bool any = false;
                for (int i = 0; i < pids.Length; i++)
                {
                    try
                    {
                        using (Process p = Process.GetProcessById(pids[i]))
                        {
                            if (!p.HasExited)
                            {
                                any = true;
                                break;
                            }
                        }
                    }
                    catch
                    {
                    }
                }

                if (!any)
                {
                    return true;
                }

                Thread.Sleep(200);
            }

            return false;
        }

        private static bool IsAppAlive(PowerPoint.Application app)
        {
            if (app == null)
            {
                return false;
            }

            try
            {
                int _ = app.Presentations.Count;
                return true;
            }
            catch
            {
                return false;
            }
        }

        private static int TrySlideCount(PowerPoint.Presentation pres)
        {
            if (pres == null)
            {
                return -1;
            }

            try
            {
                return pres.Slides.Count;
            }
            catch
            {
                return -1;
            }
        }

        private static string DescribePptSession(PowerPoint.Application app)
        {
            int excel = 0;
            try
            {
                excel = Process.GetProcessesByName("EXCEL").Length;
            }
            catch
            {
            }

            if (app == null)
            {
                return "会话 app=null excel=" + excel;
            }

            try
            {
                int n = app.Presentations.Count;
                var names = new List<string>();
                for (int i = 1; i <= n; i++)
                {
                    try
                    {
                        names.Add(app.Presentations[i].Name);
                    }
                    catch (Exception ex)
                    {
                        names.Add("[" + i + "死:" + ex.Message + "]");
                    }
                }

                return "会话 稿=" + n + " [" + string.Join(",", names) + "] excel=" + excel;
            }
            catch (Exception ex)
            {
                return "会话 读失败:" + ex.Message + " excel=" + excel;
            }
        }

        private static bool IsDeadPres(string message)
        {
            if (string.IsNullOrEmpty(message))
            {
                return false;
            }

            return message.IndexOf("Object does not exist", StringComparison.OrdinalIgnoreCase) >= 0
                || message.IndexOf("对象不存在", StringComparison.OrdinalIgnoreCase) >= 0
                || message.IndexOf("unknown member", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private static bool IsRpcText(string message)
        {
            if (string.IsNullOrEmpty(message))
            {
                return false;
            }

            return message.IndexOf("RPC", StringComparison.OrdinalIgnoreCase) >= 0
                || message.IndexOf("0x800706BA", StringComparison.OrdinalIgnoreCase) >= 0
                || message.IndexOf("被调用的对象已与其客户端断开", StringComparison.OrdinalIgnoreCase) >= 0
                || message.IndexOf("RPC server", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private static void TryRelease(object com)
        {
            if (com == null)
            {
                return;
            }

            try
            {
                Marshal.FinalReleaseComObject(com);
            }
            catch
            {
            }
        }
    }
}
