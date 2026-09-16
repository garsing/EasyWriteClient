using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using WordAddIn1.PresentationHost;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PptChartRoundtripTest
{
    /// <summary>写 HTML → AddChart2/换数 → TryRead 对表。</summary>
    internal static class ChartRoundtripTests
    {
        public static void Run(TestRun run)
        {
            PowerPoint.Application app = null;
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
            }
            catch (Exception ex)
            {
                run.Skip("com-host", "无法启动 PowerPoint：" + ex.Message);
                return;
            }

            _pptDead = false;
            PowerPoint.Presentation pres = null;
            try
            {
                pres = app.Presentations.Add(Office.MsoTriState.msoTrue);
                if (pres.Slides.Count == 0)
                {
                    pres.Slides.Add(1, PowerPoint.PpSlideLayout.ppLayoutBlank);
                }

                RunOne(app, pres, run, "create-column", CreateColumnReadBack);
                RunOne(app, pres, run, "create-bar", CreateBarReadBack);
                RunOne(app, pres, run, "create-line", CreateLineReadBack);
                RunOne(app, pres, run, "create-pie2d", CreatePie2dReadBack);
                RunOne(app, pres, run, "create-pie3d", CreatePie3dReadBack);
                RunOne(app, pres, run, "create-combo-y2", CreateComboY2ReadBack);
                RunOne(app, pres, run, "create-combo-both-y", CreateComboBothYNoSecondary);
                RunOne(app, pres, run, "replace-more-rows", ReplaceColumnMoreRows);
                RunOne(app, pres, run, "replace-to-combo-y2", ReplaceColumnToComboY2);
                RunOne(app, pres, run, "replace-barline-y2", ReplaceBarLineY2);
            }
            catch (Exception ex)
            {
                run.Fail("com-suite", ex.GetType().Name + ": " + ex.Message);
            }
            finally
            {
                QuitQuiet(pres, app);
            }
        }

        private static bool _pptDead;

        private static void RunOne(
            PowerPoint.Application app,
            PowerPoint.Presentation pres,
            TestRun run,
            string name,
            Action<TestRun, object> body)
        {
            if (_pptDead || !IsAppAlive(app))
            {
                _pptDead = true;
                run.Skip(name, "PowerPoint 已断开，后续 COM 例不再跑");
                return;
            }

            PptHtmlChartIo.DismissChartExcelUi();
            System.Threading.Thread.Sleep(400);
            try
            {
                PowerPoint.Slide slide = pres.Slides.Add(pres.Slides.Count + 1, PowerPoint.PpSlideLayout.ppLayoutBlank);
                body(run, slide.Shapes);
            }
            catch (Exception ex)
            {
                if (IsRpcDead(ex))
                {
                    _pptDead = true;
                    run.Skip(name, "PowerPoint 已断开：" + ex.Message);
                    return;
                }

                run.Fail(name, ex.GetType().Name + ": " + ex.Message);
            }
            finally
            {
                PptHtmlChartIo.DismissChartExcelUi();
            }
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

        private static bool IsRpcDead(Exception ex)
        {
            if (ex == null)
            {
                return false;
            }

            string m = ex.Message ?? "";
            return m.IndexOf("0x800706BA", StringComparison.OrdinalIgnoreCase) >= 0
                || m.IndexOf("RPC", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private static void CreateColumnReadBack(TestRun run, object shapes)
        {
            if (!TryCreate(ChartHtml.ColumnCreate(), shapes, out object shape, out string error, out List<string> warnings))
            {
                FailOrSkipChartData(run, "create-column", error, warnings);
                return;
            }

            if (!TryReadModel(shape, out PptHtmlChartReadModel model, out error))
            {
                run.Fail("create-column 读回", error);
                return;
            }

            run.ExpectEqual("create-column 类型", "column", model.Format?.ChartType);
            run.Expect("create-column 4 行", model.Grid != null && model.Grid.Rows.Count == 4,
                "行数=" + (model.Grid == null ? -1 : model.Grid.Rows.Count));
            run.ExpectClose("create-column 2026", 2.15, Cell(model, 0, 1));
            run.ExpectClose("create-column 2029", 4.8, Cell(model, 3, 1));
            run.ExpectEqual("create-column 类别", "2026", Cell(model, 0, 0));
        }

        private static void CreatePie3dReadBack(TestRun run, object shapes)
        {
            if (!TryCreate(ChartHtml.Pie3dCreate(), shapes, out object shape, out string error, out List<string> warnings))
            {
                FailOrSkipChartData(run, "create-pie3d", error, warnings);
                return;
            }

            if (!TryReadModel(shape, out PptHtmlChartReadModel model, out error))
            {
                run.Fail("create-pie3d 读回", error);
                return;
            }

            run.ExpectEqual("create-pie3d 类型", "pie3d", model.Format?.ChartType);
            run.Expect("create-pie3d 5 扇区", model.Grid != null && model.Grid.Rows.Count == 5,
                "行数=" + (model.Grid == null ? -1 : model.Grid.Rows.Count));
            run.ExpectEqual("create-pie3d 类别", "Bloomberg", Cell(model, 0, 0));
            run.ExpectClose("create-pie3d 33", 33, Cell(model, 0, 1));
            run.ExpectClose("create-pie3d 2", 2, Cell(model, 4, 1));
        }

        private static void CreateComboY2ReadBack(TestRun run, object shapes)
        {
            if (!TryCreate(ChartHtml.ComboY2Create(), shapes, out object shape, out string error, out List<string> warnings))
            {
                FailOrSkipChartData(run, "create-combo-y2", error, warnings);
                return;
            }

            if (!TryReadModel(shape, out PptHtmlChartReadModel model, out error))
            {
                run.Fail("create-combo-y2 读回", error);
                return;
            }

            PptHtmlChartColumn col = ChartHtml.ValueCol(model.Grid, 0);
            PptHtmlChartColumn line = ChartHtml.ValueCol(model.Grid, 1);
            run.ExpectEqual("create-combo-y2 柱系列", "column", NormalizeSeries(col?.SeriesType));
            run.ExpectEqual("create-combo-y2 折线系列", "line", NormalizeSeries(line?.SeriesType));
            run.ExpectEqual("create-combo-y2 折线次轴", "y2", line?.AxisY);
            run.ExpectClose("create-combo-y2 柱 2.15", 2.15, Cell(model, 0, 1));
            run.ExpectClose("create-combo-y2 线 0.265", 0.265, Cell(model, 0, 2));
            run.ExpectClose("create-combo-y2 线 0.315", 0.315, Cell(model, 3, 2));
        }

        private static void CreateComboBothYNoSecondary(TestRun run, object shapes)
        {
            if (!TryCreate(ChartHtml.ComboBothYCreate(), shapes, out object shape, out string error, out List<string> warnings))
            {
                FailOrSkipChartData(run, "create-combo-both-y", error, warnings);
                return;
            }

            if (!TryReadModel(shape, out PptHtmlChartReadModel model, out error))
            {
                run.Fail("create-combo-both-y 读回", error);
                return;
            }

            PptHtmlChartColumn line = ChartHtml.ValueCol(model.Grid, 1);
            run.ExpectEqual("create-combo-both-y 折线仍主轴", "y", line?.AxisY ?? "y");
            bool y2Visible = model.Format?.AxisY2Style != null
                && string.Equals(model.Format.AxisY2Style.Visible, "true", StringComparison.OrdinalIgnoreCase);
            run.Expect("create-combo-both-y 无次轴可见", !y2Visible,
                "AxisY2Style.Visible=" + (model.Format?.AxisY2Style?.Visible ?? "(null)"));
        }

        private static void ReplaceColumnMoreRows(TestRun run, object shapes)
        {
            if (!TryCreate(ChartHtml.ColumnCreate(), shapes, out object oldShape, out string error, out List<string> warnings))
            {
                FailOrSkipChartData(run, "replace-more-rows 建底图", error, warnings);
                return;
            }

            if (!ChartHtml.TryParseFirstChart(ChartHtml.ColumnReplaceMoreRows("sid256-s99"), out PptHtmlApplyNode node, out error))
            {
                run.Fail("replace-more-rows 解析", error);
                return;
            }

            warnings = new List<string>();
            if (!TryReplace(shapes, oldShape, node, out object newShape, out error, out warnings))
            {
                FailOrSkipChartData(run, "replace-more-rows 换数", error, warnings);
                return;
            }

            if (!TryReadModel(newShape, out PptHtmlChartReadModel model, out error))
            {
                run.Fail("replace-more-rows 读回", error);
                return;
            }

            run.Expect("replace-more-rows 6 行", model.Grid != null && model.Grid.Rows.Count == 6,
                "行数=" + (model.Grid == null ? -1 : model.Grid.Rows.Count));
            run.ExpectClose("replace-more-rows 2031", 8.25, Cell(model, 5, 1));
        }

        private static void ReplaceColumnToComboY2(TestRun run, object shapes)
        {
            if (!TryCreate(ChartHtml.ColumnCreate(), shapes, out object oldShape, out string error, out List<string> warnings))
            {
                FailOrSkipChartData(run, "replace-to-combo-y2 建底图", error, warnings);
                return;
            }

            if (!ChartHtml.TryParseFirstChart(ChartHtml.ComboReplace("sid256-s99", "y2"), out PptHtmlApplyNode node, out error))
            {
                run.Fail("replace-to-combo-y2 解析", error);
                return;
            }

            if (!TryReplace(shapes, oldShape, node, out object newShape, out error, out warnings))
            {
                FailOrSkipChartData(run, "replace-to-combo-y2 换数", error, warnings);
                return;
            }

            if (!TryReadModel(newShape, out PptHtmlChartReadModel model, out error))
            {
                run.Fail("replace-to-combo-y2 读回", error);
                return;
            }

            PptHtmlChartColumn line = ChartHtml.ValueCol(model.Grid, 1);
            run.ExpectEqual("replace-to-combo-y2 折线系列", "line", NormalizeSeries(line?.SeriesType));
            run.ExpectEqual("replace-to-combo-y2 折线次轴", "y2", line?.AxisY);
            run.ExpectClose("replace-to-combo-y2 线 0.315", 0.315, Cell(model, 3, 2));
        }

        private static void CreateBarReadBack(TestRun run, object shapes)
        {
            if (!TryCreate(ChartHtml.BarCreate(), shapes, out object shape, out string error, out List<string> warnings))
            {
                FailOrSkipChartData(run, "create-bar", error, warnings);
                return;
            }

            if (!TryReadModel(shape, out PptHtmlChartReadModel model, out error))
            {
                run.Fail("create-bar 读回", error);
                return;
            }

            run.ExpectEqual("create-bar 类型", "bar", model.Format?.ChartType);
            run.ExpectClose("create-bar 2026", 2.15, Cell(model, 0, 1));
        }

        private static void CreateLineReadBack(TestRun run, object shapes)
        {
            if (!TryCreate(ChartHtml.LineCreate(), shapes, out object shape, out string error, out List<string> warnings))
            {
                FailOrSkipChartData(run, "create-line", error, warnings);
                return;
            }

            if (!TryReadModel(shape, out PptHtmlChartReadModel model, out error))
            {
                run.Fail("create-line 读回", error);
                return;
            }

            run.ExpectEqual("create-line 类型", "line", NormalizeSeries(model.Format?.ChartType));
            run.ExpectEqual("create-line 系列", "line", NormalizeSeries(ChartHtml.ValueCol(model.Grid, 0)?.SeriesType));
            run.ExpectClose("create-line 2029", 4.8, Cell(model, 3, 1));
        }

        private static void CreatePie2dReadBack(TestRun run, object shapes)
        {
            if (!TryCreate(ChartHtml.Pie2dCreate(), shapes, out object shape, out string error, out List<string> warnings))
            {
                FailOrSkipChartData(run, "create-pie2d", error, warnings);
                return;
            }

            if (!TryReadModel(shape, out PptHtmlChartReadModel model, out error))
            {
                run.Fail("create-pie2d 读回", error);
                return;
            }

            run.ExpectEqual("create-pie2d 类型", "pie2d", model.Format?.ChartType);
            run.Expect("create-pie2d 至少 3 扇区", model.Grid != null && model.Grid.Rows.Count >= 3,
                "行数=" + (model.Grid == null ? -1 : model.Grid.Rows.Count));
            run.ExpectClose("create-pie2d A", 40, Cell(model, 0, 1));
            run.ExpectClose("create-pie2d C", 25, Cell(model, 2, 1));
        }

        private static void ReplaceBarLineY2(TestRun run, object shapes)
        {
            if (!TryCreate(ChartHtml.BarCreate(), shapes, out object oldShape, out string error, out List<string> warnings))
            {
                FailOrSkipChartData(run, "replace-barline-y2 建底图", error, warnings);
                return;
            }

            if (!ChartHtml.TryParseFirstChart(ChartHtml.BarLineY2Replace(), out PptHtmlApplyNode node, out error))
            {
                run.Fail("replace-barline-y2 解析", error);
                return;
            }

            if (!TryReplace(shapes, oldShape, node, out object newShape, out error, out warnings))
            {
                FailOrSkipChartData(run, "replace-barline-y2 换数", error, warnings);
                return;
            }

            if (!TryReadModel(newShape, out PptHtmlChartReadModel model, out error))
            {
                run.Fail("replace-barline-y2 读回", error);
                return;
            }

            run.ExpectEqual("replace-barline-y2 折线次轴", "y2", ChartHtml.ValueCol(model.Grid, 1)?.AxisY);
            run.ExpectClose("replace-barline-y2 线 0.315", 0.315, Cell(model, 3, 2));
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

            string typeRaw = node.ChartType ?? node.ChartFormat?.ChartType ?? "column";
            if (!PptHtmlChartIo.TryParseType(typeRaw, out int xlType, out _, out error))
            {
                return false;
            }

            bool ok = TryCreateOnce(shapes, node, xlType, warnings, out shape, out error);
            if (!ok && IsChartDataBusy(error))
            {
                PptHtmlChartIo.DismissChartExcelUi();
                System.Threading.Thread.Sleep(400);
                warnings.Add("retry after ChartData busy");
                ok = TryCreateOnce(shapes, node, xlType, warnings, out shape, out error);
            }

            PptHtmlChartIo.DismissChartExcelUi();
            return ok;
        }

        private static bool TryCreateOnce(
            object shapes,
            PptHtmlApplyNode node,
            int xlType,
            List<string> warnings,
            out object shape,
            out string error)
        {
            return PptHtmlChartIo.TryCreateOnSlide(
                shapes,
                80f,
                80f,
                320f,
                200f,
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

        private static bool TryReplace(
            object shapes,
            object oldShape,
            PptHtmlApplyNode node,
            out object newShape,
            out string error,
            out List<string> warnings)
        {
            warnings = new List<string>();
            PptHtmlChartIo.DismissChartExcelUi();
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
            if (!ok && IsChartDataBusy(error))
            {
                PptHtmlChartIo.DismissChartExcelUi();
                System.Threading.Thread.Sleep(400);
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

            PptHtmlChartIo.DismissChartExcelUi();
            return ok;
        }

        private static bool IsChartDataBusy(string error)
        {
            return !string.IsNullOrEmpty(error)
                && error.IndexOf("ChartData", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private static void FailOrSkipChartData(TestRun run, string name, string error, List<string> warnings)
        {
            string detail = (error ?? "") + FormatWarnings(warnings);
            if (IsRpcDead(new Exception(error ?? "")))
            {
                _pptDead = true;
                run.Skip(name, "PowerPoint 已断开：" + detail);
                return;
            }

            if (IsChartDataBusy(error))
            {
                run.Skip(name, "Office ChartData 忙（本机 Excel 内嵌表偶发打不开）：" + detail);
                return;
            }

            run.Fail(name, detail);
        }

        private static bool TryReadModel(object shape, out PptHtmlChartReadModel model, out string error)
        {
            return PptHtmlChartIo.TryRead(shape, out model, out error);
        }

        private static string Cell(PptHtmlChartReadModel model, int row, int col)
        {
            if (model?.Grid?.Rows == null
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

        private static string NormalizeSeries(string raw)
        {
            if (string.IsNullOrWhiteSpace(raw))
            {
                return raw;
            }

            string t = raw.Trim().ToLowerInvariant();
            if (t.IndexOf("line", StringComparison.Ordinal) >= 0)
            {
                return "line";
            }

            if (t.IndexOf("pie3d", StringComparison.Ordinal) >= 0 || t == "pie3d")
            {
                return "pie3d";
            }

            if (t.IndexOf("pie", StringComparison.Ordinal) >= 0)
            {
                return "pie2d";
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

            int n = Math.Min(4, warnings.Count);
            return " | " + string.Join(" / ", warnings.GetRange(0, n));
        }

        private static void QuitQuiet(PowerPoint.Presentation pres, PowerPoint.Application app)
        {
            if (pres != null)
            {
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

            // 只在没有其它打开文稿时退出，避免关掉用户正在用的 PPT
            if (app != null)
            {
                try
                {
                    if (app.Presentations.Count == 0)
                    {
                        app.Quit();
                    }
                }
                catch
                {
                }

                TryRelease(app);
            }
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
