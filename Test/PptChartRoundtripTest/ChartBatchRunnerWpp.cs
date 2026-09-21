using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using WordAddIn1.OpenFiles;
using WordAddIn1.PresentationHost;

namespace PptChartRoundtripTest
{
    /// <summary>WPS 演示晚绑定图表往返；建图/换数/读回复用 ChartBatchRunner.RunOneSlide。</summary>
    internal static class ChartBatchRunnerWpp
    {
        /// <summary>灌完 Close 内嵌簿后同进程可再开表；仍隔几张重启防 COM 老化。</summary>
        private const int RecycleEvery = 4;
        private const int PpLayoutBlank = 12;

        public static void Run(TestRun run, int? onlyBatch, string outDir, IList<string> nameFilters)
        {
            RunSuite(
                run,
                onlyBatch,
                outDir,
                ChartCaseCatalog.Build(),
                ChartCaseCatalog.BatchCount,
                "chart-batch-wpp-",
                nameFilters);
        }

        public static void RunAttrs(TestRun run, int? onlyBatch, string outDir, IList<string> nameFilters)
        {
            RunSuite(
                run,
                onlyBatch,
                outDir,
                ChartAttrCatalog.Build(),
                ChartAttrCatalog.BatchCount,
                "chart-attr-batch-wpp-",
                nameFilters);
        }

        private static void RunSuite(
            TestRun run,
            int? onlyBatch,
            string outDir,
            List<ChartCase> all,
            int batchCount,
            string filePrefix,
            IList<string> nameFilters)
        {
            Directory.CreateDirectory(outDir);
            if (nameFilters != null && nameFilters.Count > 0)
            {
                Console.WriteLine("过滤用例名含: " + string.Join(", ", nameFilters));
            }

            Console.WriteLine("宿主: WPP (Kwpp.Application 晚绑定)");

            for (int batch = 1; batch <= batchCount; batch++)
            {
                if (onlyBatch.HasValue && onlyBatch.Value != batch)
                {
                    continue;
                }

                if (!BatchHasCases(all, batch, nameFilters))
                {
                    continue;
                }

                RunBatch(run, all, batch, batchCount, outDir, filePrefix, nameFilters);
            }
        }

        private static bool BatchHasCases(List<ChartCase> all, int batch, IList<string> nameFilters)
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
            List<ChartCase> all,
            int batch,
            int batchCount,
            string outDir,
            string filePrefix,
            IList<string> nameFilters)
        {
            string path = Path.Combine(outDir, filePrefix + batch + ".pptx");
            Console.WriteLine();
            Console.WriteLine("--- WPP 图表往返 批次 " + batch + "/" + batchCount + " → " + path + " ---");

            var cases = new List<ChartCase>();
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
                Thread.Sleep(2500);
                if (!TryStartWithRetry(out app, out string startError))
                {
                    run.CaseFail("batch-wpp-" + batch + "-open", startError);
                    SkipRest(run, cases, 0, "WPP 无法启动: " + startError);
                    return;
                }

                if (!TryNewDeckWithRetry(app, out pres, out startError))
                {
                    SoftQuit(ref app);
                    Thread.Sleep(2500);
                    if (!TryStartWithRetry(out app, out startError)
                        || !TryNewDeckWithRetry(app, out pres, out startError))
                    {
                        run.CaseFail("batch-wpp-" + batch + "-open", startError);
                        SkipRest(run, cases, 0, "WPP 无法建空白稿: " + startError);
                        return;
                    }
                }

                for (int i = 0; i < cases.Count; i++)
                {
                    ChartCase one = cases[i];
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
                        string rpc = RunWppCase(run, ref app, ref pres, path, one, tag);
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

                    Console.WriteLine("  内嵌表未就绪，重启后重试 " + tag);
                    if (!TryRecycle(ref app, ref pres, path, out string recErr2))
                    {
                        run.CaseFail(tag, "WPP 断开且重启失败: " + recErr2);
                        SkipRest(run, cases, i + 1, "WPP 无法重启");
                        break;
                    }

                    sinceRecycle = 0;
                    try
                    {
                        string rpc = RunWppCase(run, ref app, ref pres, path, one, tag);
                        if (rpc != null)
                        {
                            run.CaseFail(tag, "重启后仍失败: " + rpc);
                            continue;
                        }

                        SaveQuiet(pres, path);
                        sinceRecycle++;
                        Console.WriteLine("  … " + (i + 1) + "/" + cases.Count
                            + " （重启后重试成功）");
                    }
                    catch (Exception ex)
                    {
                        run.CaseFail(tag, "重启后仍失败: " + ex.Message);
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

        /// <summary>换数必须先建底图再重启，否则同一进程第二次打不开内嵌表。</summary>
        private static string RunWppCase(
            TestRun run,
            ref object app,
            ref object presentation,
            string path,
            ChartCase one,
            string tag)
        {
            if (one != null && one.IsReplace)
            {
                object probeSlide = AddBlankSlide(presentation);
                TryAddCaseLabel(probeSlide, tag);
                object probeShapes = WppCom.GetProperty(probeSlide, "Shapes");
                string sameProcess = ChartBatchRunner.RunOneSlide(run, probeShapes, one, tag);
                if (sameProcess == null)
                {
                    return null;
                }

                if (ChartBatchRunner.IsRetryableCom(sameProcess))
                {
                    Console.WriteLine("  同进程换数未释放内嵌表，改为建底图后重启再灌 " + tag);
                    return RunReplaceCase(run, ref app, ref presentation, path, one, tag);
                }

                return sameProcess;
            }

            object slide = AddBlankSlide(presentation);
            TryAddCaseLabel(slide, tag);
            object shapes = WppCom.GetProperty(slide, "Shapes");
            return ChartBatchRunner.RunOneSlide(run, shapes, one, tag);
        }

        private static string RunReplaceCase(
            TestRun run,
            ref object app,
            ref object presentation,
            string path,
            ChartCase one,
            string tag)
        {
            object slide = AddBlankSlide(presentation);
            TryAddCaseLabel(slide, tag);
            object shapes = WppCom.GetProperty(slide, "Shapes");
            if (!ChartBatchRunner.TryCreate(
                one.CreateHtml,
                shapes,
                out object shape,
                out string error,
                out List<string> warnings))
            {
                if (ChartBatchRunner.IsRetryableCom(error))
                {
                    return error;
                }

                run.CaseFail(tag, "建底图失败: " + error + ChartBatchRunner.FormatWarnings(warnings));
                return null;
            }

            if (one.AfterCreateMutate != null)
            {
                try
                {
                    one.AfterCreateMutate(shape);
                }
                catch (Exception ex)
                {
                    run.CaseFail(tag, "底图微调失败: " + ex.Message);
                    return null;
                }
            }

            SaveQuiet(presentation, path);
            if (!TryRecycle(ref app, ref presentation, path, out error))
            {
                return "换数前重启失败: " + error;
            }

            if (!TryFindLastChart(presentation, out shapes, out shape, out error))
            {
                run.CaseFail(tag, "重开后找不到底图: " + error);
                return null;
            }

            if (!ChartHtml.TryParseFirstChart(one.ReplaceHtml, out PptHtmlApplyNode node, out error))
            {
                run.CaseFail(tag, "换数稿解析失败: " + error);
                return null;
            }

            if (!ChartBatchRunner.TryReplace(shapes, shape, node, out shape, out error, out warnings))
            {
                if (ChartBatchRunner.IsRetryableCom(error))
                {
                    SaveQuiet(presentation, path);
                    if (!TryRecycle(ref app, ref presentation, path, out string recErr))
                    {
                        return "换数后重启失败: " + recErr;
                    }

                    if (!TryFindLastChart(presentation, out shapes, out shape, out recErr))
                    {
                        return recErr;
                    }

                    if (!ChartBatchRunner.TryReplace(shapes, shape, node, out shape, out error, out warnings))
                    {
                        if (ChartBatchRunner.IsRetryableCom(error))
                        {
                            return error;
                        }

                        run.CaseFail(tag, "换数失败: " + error + ChartBatchRunner.FormatWarnings(warnings));
                        return null;
                    }
                }
                else
                {
                    run.CaseFail(tag, "换数失败: " + error + ChartBatchRunner.FormatWarnings(warnings));
                    return null;
                }
            }

            if (!PptHtmlChartIo.TryRead(shape, out PptHtmlChartReadModel model, out error))
            {
                if (ChartBatchRunner.IsRetryableCom(error))
                {
                    return error;
                }

                run.CaseFail(tag, "读回失败: " + error);
                return null;
            }

            string mismatch = ChartBatchRunner.MatchExpect(one, model);
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

        private static bool TryFindLastChart(
            object presentation,
            out object shapes,
            out object shape,
            out string error)
        {
            shapes = null;
            shape = null;
            error = null;
            try
            {
                object slides = WppCom.GetProperty(presentation, "Slides");
                int slideCount = Convert.ToInt32(WppCom.GetProperty(slides, "Count") ?? 0);
                if (slideCount < 1)
                {
                    error = "重开后无幻灯片";
                    return false;
                }

                object slide = WppCom.GetIndexed(slides, slideCount);
                shapes = WppCom.GetProperty(slide, "Shapes");
                int n = Convert.ToInt32(WppCom.GetProperty(shapes, "Count") ?? 0);
                for (int i = n; i >= 1; i--)
                {
                    object one = WppCom.GetIndexed(shapes, i);
                    if (PptHtmlChartIo.LooksLikeChart(one))
                    {
                        shape = one;
                        return true;
                    }
                }

                error = "末页无图表";
                return false;
            }
            catch (Exception ex)
            {
                error = ex.Message;
                return false;
            }
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
                object box = WppCom.Invoke(shapes, "AddTextbox", 1, 18.0, 12.0, 680.0, 36.0);
                object tf = WppCom.GetProperty(box, "TextFrame");
                object tr = WppCom.GetProperty(tf, "TextRange");
                WppCom.TrySetProperty(tr, "Text", tag);
            }
            catch
            {
            }
        }

        private static bool TryStartWithRetry(out object app, out string error)
        {
            app = null;
            error = null;
            for (int i = 0; i < 4; i++)
            {
                if (i > 0)
                {
                    Console.WriteLine("  重试启动 WPP #" + (i + 1));
                    Thread.Sleep(1500);
                }

                if (WppCom.TryCreateApplication(out app, out error))
                {
                    return true;
                }

                SoftQuit(ref app);
            }

            return false;
        }

        private static bool TryNewDeckWithRetry(object app, out object presentation, out string error)
        {
            presentation = null;
            error = null;
            for (int i = 0; i < 3; i++)
            {
                if (TryNewDeck(app, out presentation, out error))
                {
                    return true;
                }

                Thread.Sleep(800);
            }

            return false;
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
            Thread.Sleep(1500);

            if (!TryStartWithRetry(out app, out error))
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

        private static void SkipRest(TestRun run, List<ChartCase> cases, int fromIndex, string reason)
        {
            for (int i = fromIndex; i < cases.Count; i++)
            {
                run.CaseSkip(cases[i].Name, reason);
            }
        }

        private static bool IsRpcText(string text)
        {
            if (string.IsNullOrEmpty(text))
            {
                return false;
            }

            return ChartBatchRunner.IsRetryableCom(text)
                || text.IndexOf("RPC", StringComparison.OrdinalIgnoreCase) >= 0
                || text.IndexOf("被呼叫方", StringComparison.OrdinalIgnoreCase) >= 0
                || text.IndexOf("服务器忙", StringComparison.OrdinalIgnoreCase) >= 0;
        }
    }
}
