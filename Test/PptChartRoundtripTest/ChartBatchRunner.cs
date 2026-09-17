using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using WordAddIn1.PresentationHost;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PptChartRoundtripTest
{
    /// <summary>一页一图；约 50 页一份 PPT，共 6 份。COM 挂了就存盘、重启后续跑。</summary>
    internal static class ChartBatchRunner
    {
        private const int RecycleEvery = 8;

        public static void Run(TestRun run, int? onlyBatch, string outDir, IList<string> nameFilters)
        {
            RunSuite(
                run,
                onlyBatch,
                outDir,
                ChartCaseCatalog.Build(),
                ChartCaseCatalog.BatchCount,
                "chart-batch-",
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
                "chart-attr-batch-",
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
            Console.WriteLine("--- 正式用例 批次 " + batch + "/" + batchCount
                + " （约 50 页）→ " + path + " ---");

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
                    ChartCase one = cases[i];
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
                    else if (sinceRecycle >= RecycleEvery)
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
                    try
                    {
                        PowerPoint.Slide slide = pres.Slides.Add(
                            pres.Slides.Count + 1,
                            PowerPoint.PpSlideLayout.ppLayoutBlank);
                        TryAddCaseLabel(slide, tag);
                        string rpc = RunOneSlide(run, slide.Shapes, one, tag);
                        if (rpc != null)
                        {
                            retry = true;
                        }
                        else
                        {
                            SaveQuiet(pres, path);
                            sinceRecycle++;
                            Console.WriteLine("  … " + (i + 1) + "/" + cases.Count
                                + " 页=" + pres.Slides.Count);
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
                        string rpc = RunOneSlide(run, slide.Shapes, one, tag);
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
                    int pages = 0;
                    try
                    {
                        pages = pres != null && IsAppAlive(app) ? pres.Slides.Count : -1;
                    }
                    catch
                    {
                        pages = -1;
                    }

                    Console.WriteLine("  已保存 " + path + (pages >= 0 ? " 页数=" + pages : ""));
                }
            }
            finally
            {
                CloseQuiet(pres);
                SoftQuit(ref app);
            }
        }

        /// <returns>RPC 死信，否则 null（已记 CaseOk/Fail）。</returns>
        private static string RunOneSlide(TestRun run, object shapes, ChartCase one, string tag)
        {
            if (!TryCreate(one.CreateHtml, shapes, out object shape, out string error, out List<string> warnings))
            {
                if (IsRpcText(error))
                {
                    return error;
                }

                run.CaseFail(tag, "建图失败: " + error + FormatWarnings(warnings));
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

            if (one.IsReplace)
            {
                if (!ChartHtml.TryParseFirstChart(one.ReplaceHtml, out PptHtmlApplyNode node, out error))
                {
                    run.CaseFail(tag, "换数稿解析失败: " + error);
                    return null;
                }

                if (!TryReplace(shapes, shape, node, out shape, out error, out warnings))
                {
                    if (IsRpcText(error))
                    {
                        return error;
                    }

                    run.CaseFail(tag, "换数失败: " + error + FormatWarnings(warnings));
                    return null;
                }
            }

            if (!PptHtmlChartIo.TryRead(shape, out PptHtmlChartReadModel model, out error))
            {
                if (IsRpcText(error))
                {
                    return error;
                }

                run.CaseFail(tag, "读回失败: " + error);
                return null;
            }

            string mismatch = MatchExpect(one, model);
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

        private static string MatchExpect(ChartCase one, PptHtmlChartReadModel model)
        {
            if (model == null || model.Grid == null)
            {
                return "读回无表";
            }

            string actualType = NormalizeType(model.Format != null ? model.Format.ChartType : null);
            if (!string.Equals(NormalizeType(one.ExpectType), actualType, StringComparison.Ordinal))
            {
                return "类型期望 " + one.ExpectType + " 实际 " + actualType;
            }

            int rows = model.Grid.Rows == null ? 0 : model.Grid.Rows.Count;
            if (rows < one.ExpectRows)
            {
                return "行数期望 ≥" + one.ExpectRows + " 实际 " + rows;
            }

            int series = 0;
            if (model.Grid.Columns != null)
            {
                for (int i = 0; i < model.Grid.Columns.Count; i++)
                {
                    if (model.Grid.Columns[i] != null && model.Grid.Columns[i].Role != "category")
                    {
                        series++;
                    }
                }
            }

            if (series < one.ExpectSeries)
            {
                return "系列数期望 ≥" + one.ExpectSeries + " 实际 " + series;
            }

            if (!string.IsNullOrEmpty(one.ExpectLineAxis))
            {
                PptHtmlChartColumn line = ChartHtml.ValueCol(model.Grid, 1);
                string axis = line == null ? null : line.AxisY;
                if (string.IsNullOrEmpty(axis))
                {
                    axis = "y";
                }

                if (!string.Equals(one.ExpectLineAxis, axis, StringComparison.Ordinal))
                {
                    return "折线轴期望 " + one.ExpectLineAxis + " 实际 " + axis;
                }
            }

            string cat = Cell(model, 0, 0);
            if (!string.Equals(one.FirstCategory, cat, StringComparison.Ordinal))
            {
                return "首类期望 " + one.FirstCategory + " 实际 " + (cat ?? "(null)");
            }

            if (!CloseTo(one.FirstValue, Cell(model, 0, 1)))
            {
                return "首值期望 " + one.FirstValue.ToString(CultureInfo.InvariantCulture)
                    + " 实际 " + (Cell(model, 0, 1) ?? "(null)");
            }

            if (!CloseTo(one.LastValue, Cell(model, one.ExpectRows - 1, 1)))
            {
                return "末值期望 " + one.LastValue.ToString(CultureInfo.InvariantCulture)
                    + " 实际 " + (Cell(model, one.ExpectRows - 1, 1) ?? "(null)");
            }

            return ChartAttrMatch.Check(one.Attrs, model);
        }

        private static bool TryCreate(
            string html,
            object shapes,
            out object shape,
            out string error,
            out List<string> warnings)
        {
            shape = null;
            warnings = new List<string>();
            if (!ChartHtml.TryParseFirstChart(html, out PptHtmlApplyNode node, out error))
            {
                return false;
            }

            string typeRaw = node.ChartType ?? (node.ChartFormat != null ? node.ChartFormat.ChartType : null) ?? "column";
            if (!PptHtmlChartIo.TryParseType(typeRaw, out int xlType, out _, out error))
            {
                return false;
            }

            bool ok = PptHtmlChartIo.TryCreateOnSlide(
                shapes,
                72f,
                80f,
                480f,
                280f,
                xlType,
                node.ChartGrid,
                node.ChartFormat ?? new PptHtmlChartFormat(),
                warnings,
                out shape,
                out error,
                -1,
                true,
                true);
            if (!ok && IsChartDataBusy(error) && !IsRpcText(error))
            {
                Thread.Sleep(500);
                warnings.Add("retry after ChartData busy");
                ok = PptHtmlChartIo.TryCreateOnSlide(
                    shapes,
                    72f,
                    80f,
                    480f,
                    280f,
                    xlType,
                    node.ChartGrid,
                    node.ChartFormat ?? new PptHtmlChartFormat(),
                    warnings,
                    out shape,
                    out error,
                    -1,
                    true,
                    true);
            }

            return ok;
        }

        private static bool TryReplace(
            object shapes,
            object oldShape,
            PptHtmlApplyNode node,
            out object newShape,
            out string error,
            out List<string> warnings)
        {
            warnings = new List<string>();
            bool ok = PptHtmlChartIo.TryReplaceOnSlide(
                shapes,
                oldShape,
                node.ChartGrid,
                node.ChartFormat,
                null,
                null,
                null,
                null,
                warnings,
                out newShape,
                out error);
            if (!ok && IsChartDataBusy(error) && !IsRpcText(error))
            {
                Thread.Sleep(500);
                warnings.Add("retry after ChartData busy");
                ok = PptHtmlChartIo.TryReplaceOnSlide(
                    shapes,
                    oldShape,
                    node.ChartGrid,
                    node.ChartFormat,
                    null,
                    null,
                    null,
                    null,
                    warnings,
                    out newShape,
                    out error);
            }

            return ok;
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

        private static void SoftQuit(ref PowerPoint.Application app)
        {
            if (app == null)
            {
                return;
            }

            try
            {
                if (IsAppAlive(app) && app.Presentations.Count == 0)
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

        private static void SkipRest(TestRun run, List<ChartCase> cases, int fromIndex, string reason)
        {
            for (int i = fromIndex; i < cases.Count; i++)
            {
                run.CaseSkip(cases[i].Name, reason);
            }
        }

        private static bool IsChartDataBusy(string error)
        {
            return !string.IsNullOrEmpty(error)
                && error.IndexOf("ChartData", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private static void TryAddCaseLabel(PowerPoint.Slide slide, string text)
        {
            try
            {
                PowerPoint.Shape box = slide.Shapes.AddTextbox(
                    Office.MsoTextOrientation.msoTextOrientationHorizontal,
                    18f,
                    12f,
                    680f,
                    36f);
                box.TextFrame.TextRange.Text = text;
                box.TextFrame.TextRange.Font.Size = 11f;
                box.TextFrame.TextRange.Font.Color.RGB = 0x333333;
            }
            catch
            {
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

        private static string Cell(PptHtmlChartReadModel model, int row, int col)
        {
            if (model.Grid.Rows == null
                || row < 0
                || row >= model.Grid.Rows.Count
                || model.Grid.Rows[row] == null
                || col < 0
                || col >= model.Grid.Rows[row].Count)
            {
                return null;
            }

            return model.Grid.Rows[row][col];
        }

        private static bool CloseTo(double expected, string actualRaw)
        {
            return double.TryParse(actualRaw, NumberStyles.Float, CultureInfo.InvariantCulture, out double actual)
                && Math.Abs(expected - actual) <= 1e-6;
        }

        private static string NormalizeType(string raw)
        {
            if (string.IsNullOrWhiteSpace(raw))
            {
                return raw;
            }

            string t = raw.Trim().ToLowerInvariant();
            if (t.IndexOf("pie3d", StringComparison.Ordinal) >= 0)
            {
                return "pie3d";
            }

            if (t.IndexOf("pie", StringComparison.Ordinal) >= 0)
            {
                return "pie2d";
            }

            if (t.IndexOf("line", StringComparison.Ordinal) >= 0)
            {
                return "line";
            }

            if (t.IndexOf("bar", StringComparison.Ordinal) >= 0)
            {
                return "bar";
            }

            return "column";
        }

        private static string FormatWarnings(List<string> warnings)
        {
            if (warnings == null || warnings.Count == 0)
            {
                return "";
            }

            int n = Math.Min(3, warnings.Count);
            return " | " + string.Join(" / ", warnings.GetRange(0, n));
        }

        private static bool IsAppAlive(PowerPoint.Application app)
        {
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

        private static bool IsRpcText(string text)
        {
            if (string.IsNullOrEmpty(text))
            {
                return false;
            }

            return text.IndexOf("0x800706BA", StringComparison.OrdinalIgnoreCase) >= 0
                || text.IndexOf("RPC 服务器不可用", StringComparison.OrdinalIgnoreCase) >= 0
                || text.IndexOf("RPC server", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private static void TryRelease(object com)
        {
            if (com == null)
            {
                return;
            }

            try
            {
                Marshal.ReleaseComObject(com);
            }
            catch
            {
            }
        }
    }
}
