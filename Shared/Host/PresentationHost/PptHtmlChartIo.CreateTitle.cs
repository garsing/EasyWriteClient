using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using System.Xml.Linq;
using WordAddIn1.OpenFiles;


namespace WordAddIn1.PresentationHost
{
    internal static partial class PptHtmlChartIo
    {
        public static bool TryCreateOnSlide(
            object shapes,
            float left,
            float top,
            float width,
            float height,
            int xlType,
            PptHtmlChartGrid grid,
            PptHtmlChartFormat format,
            List<string> warnings,
            out object shape,
            out string error,
            int chartStyle = -1,
            bool newLayout = true,
            bool applyHtmlChrome = true)
        {
            shape = null;
            error = null;
            if (grid == null || !grid.IsPourable)
            {
                error = "新建 chart 必须在节点内嵌 <table> 灌数，不能建空图";
                return false;
            }

            if (chartStyle < 0
                && format != null
                && int.TryParse(format.ChartStyle, NumberStyles.Integer, CultureInfo.InvariantCulture, out int htmlStyle)
                && htmlStyle > 0)
            {
                chartStyle = htmlStyle;
            }

            PrepareChartExcelUiSuppression();
            bool avoidLive = HostAvoidsFromShapes(shapes);
            PptHtmlApplyHangLog.Mark("建图进入 avoidLive=" + avoidLive);
            if (avoidLive)
            {
                PourLog(warnings, "PPT 磁盘改内嵌表后插入，不在本进程打开副本");
                if (!TryCreateChartOffLiveDeck(
                    shapes,
                    left,
                    top,
                    width,
                    height,
                    xlType,
                    grid,
                    format,
                    warnings,
                    chartStyle,
                    newLayout,
                    out shape,
                    out error))
                {
                    return false;
                }

                object pastedChart = TryGetChart(shape);
                if (pastedChart == null)
                {
                    TryDelete(shape);
                    shape = null;
                    error = "PPT 贴回后无法访问 Chart";
                    return false;
                }

                PptHtmlApplyHangLog.Mark("建图贴回完成，开始收尾 chrome=" + applyHtmlChrome);
                bool finished = TryFinishNewChart(
                    ref shape,
                    ref pastedChart,
                    grid,
                    format,
                    xlType,
                    applyHtmlChrome,
                    warnings,
                    out error);
                PptHtmlApplyHangLog.Mark("建图收尾返回 ok=" + finished);
                return finished;
            }

            PourLog(warnings, "即将 AddChart2 style=" + chartStyle
                + " xl=" + xlType
                + " newLayout=" + newLayout
                + " box=" + left.ToString("0.#", CultureInfo.InvariantCulture) + ","
                + top.ToString("0.#", CultureInfo.InvariantCulture) + " "
                + width.ToString("0.#", CultureInfo.InvariantCulture) + "x"
                + height.ToString("0.#", CultureInfo.InvariantCulture));
            shape = TryAddChartOnShapes(
                shapes,
                left,
                top,
                width,
                height,
                xlType,
                chartStyle,
                newLayout,
                warnings,
                out error);
            if (shape == null)
            {
                if (string.IsNullOrEmpty(error))
                {
                    error = "创建图表失败";
                }

                return false;
            }

            object chart = TryGetChart(shape);
            if (chart == null)
            {
                TryDelete(shape);
                error = "创建图表后无法访问 Chart";
                return false;
            }

            PourLog(warnings, "建图后取 Chart " + (chart == null ? "null" : "ok")
                + " | " + DescribeLiveSeries(chart));
            DismissChartExcelUiForChart(chart);
            if (!TryPourGrid(chart, grid, out error, warnings))
            {
                if (TryFixSeriesViaOoxml(shape, grid, warnings, out shape, out _))
                {
                    chart = TryGetChart(shape);
                    PourLog(warnings, "OOXML 修点后 " + DescribeLiveSeries(chart)
                        + " | " + DescribeSeriesExtra(chart));
                }
                else
                {
                    TryDelete(shape);
                    shape = null;
                    return false;
                }
            }

            return TryFinishNewChart(
                ref shape,
                ref chart,
                grid,
                format,
                xlType,
                applyHtmlChrome,
                warnings,
                out error);
        }

        private static bool TryFinishNewChart(
            ref object shape,
            ref object chart,
            PptHtmlChartGrid grid,
            PptHtmlChartFormat format,
            int xlType,
            bool applyHtmlChrome,
            List<string> warnings,
            out string error)
        {
            error = null;
            PptHtmlApplyHangLog.Mark("收尾进入 chrome=" + applyHtmlChrome);
            try
            {
                if (applyHtmlChrome)
                {
                    PptHtmlApplyHangLog.Mark("收尾套格式前");
                    if (!TryApplyFormat(chart, format, warnings, out error))
                    {
                        TryDelete(shape);
                        shape = null;
                        return false;
                    }

                    PptHtmlApplyHangLog.Mark("收尾套格式返回");
                    ChartStyleSnap htmlSnap = SnapFromFormat(format, grid);
                    // 新建无旧图：展示开关按「未提默认关」，避免 AddChart2 默认 ChartTitle=系列名
                    EnsureDisplaySwitches(oldSnap: null, format, htmlSnap, warnings);
                    FinishLineChartLayout(chart, xlType, format, warnings);
                    TryApplyStyleSnap(chart, htmlSnap, warnings, grid, format);
                    PptHtmlApplyHangLog.Mark("收尾套样式返回");
                    TryInheritAxisChrome(chart, htmlSnap, warnings);
                    EnsureNewPieVariesByCategory(chart, grid, warnings);
                    TryApplyPieExplosionFromFormat(chart, format, warnings);
                    TryRestorePieChartType(chart, htmlSnap);
                    TryPinOoxmlChrome(ref shape, ref chart, htmlSnap, warnings, xlType);
                }
                else
                {
                    EnsureCategoryAxisLabels(chart, grid);
                }

                return true;
            }
            finally
            {
                DismissChartExcelUiForChart(chart);
            }
        }

        private static bool TryCreateChartOffLiveDeck(
            object destShapes,
            float left,
            float top,
            float width,
            float height,
            int xlType,
            PptHtmlChartGrid grid,
            PptHtmlChartFormat format,
            List<string> warnings,
            int chartStyle,
            bool newLayout,
            out object shape,
            out string error)
        {
            shape = null;
            error = null;
            object destSlide = null;
            object destPres = null;
            object app = null;
            try
            {
                destSlide = destShapes == null ? null : WppCom.GetProperty(destShapes, "Parent");
                destPres = destSlide == null ? null : WppCom.GetProperty(destSlide, "Parent");
                app = destPres == null ? null : WppCom.GetProperty(destPres, "Application");
            }
            catch (Exception)
            {
                destSlide = null;
                destPres = null;
                app = null;
            }
            if ((destPres == null || !IsDeckAlive(destPres)) && IsDeckAlive(_chartPres))
            {
                if (!TryAdoptReopenedDeck(ref destSlide, ref destPres, ref app, ref destShapes))
                {
                    error = "PPT 旧进程换稿：接不回新进程里的稿";
                    return false;
                }
            }

            if (destPres == null || app == null)
            {
                error = "PPT 旁路建图：没有 Presentation";
                return false;
            }

            PptHtmlApplyHangLog.Mark("旁路即将检查是否换进程重开");
            if (!TryRestartPowerPointWhenDeckChanged(
                ref destSlide,
                ref destPres,
                ref app,
                ref destShapes,
                warnings,
                out error))
            {
                PptHtmlApplyHangLog.Mark("旁路重开失败 " + (error ?? ""));
                return false;
            }

            PptHtmlApplyHangLog.Mark("旁路重开检查结束");

            string tempPath = WordAddIn1.ChartOoxmlHost.PptChartPackageHost.NewTempPath("旁路建图");
            try
            {
                int destIndex = 1;
                try
                {
                    destIndex = Convert.ToInt32(WppCom.GetProperty(destSlide, "SlideIndex") ?? 1);
                }
                catch (Exception)
                {
                }

                if (destIndex < 1)
                {
                    destIndex = 1;
                }

                PptHtmlApplyHangLog.Mark("旁路 SaveCopy 前 " + tempPath);
                if (!WordAddIn1.ChartOoxmlHost.PptChartPackageHost.TrySaveCopy(destPres, tempPath, out error))
                {
                    if (string.IsNullOrEmpty(error))
                    {
                        error = "PPT 旁路建图：SaveCopy 未落盘";
                    }

                    return false;
                }

                PptHtmlApplyHangLog.Mark("旁路 SaveCopy 后，写内嵌图前");
                if (!TryEnsureEmbeddedChart(
                    tempPath,
                    destIndex,
                    left,
                    top,
                    width,
                    height,
                    grid,
                    warnings,
                    out int shapeId,
                    out error))
                {
                    return false;
                }

                PptHtmlApplyHangLog.Mark("旁路写内嵌图后，改包前");
                if (!PptHtmlOoxmlIo.TryRewriteSavedFile(
                    tempPath,
                    (zip, parts) => TryRewriteChartXmlInPackage(
                        zip,
                        parts,
                        destIndex,
                        shapeId,
                        (doc, w) =>
                        {
                            if (!RewriteChartSeries(doc, grid, w))
                            {
                                return false;
                            }

                            RewriteOoxmlChrome(doc, SnapFromFormat(format, grid), w, xlType);
                            return true;
                        },
                        warnings,
                        "灌数",
                        grid),
                    out error,
                    "灌数",
                    warnings))
                {
                    return false;
                }

                PptHtmlApplyHangLog.Mark("旁路改包后，插入前");
                shape = TryInsertChartSlide(
                    destPres,
                    destShapes,
                    destIndex,
                    tempPath,
                    left,
                    top,
                    width,
                    height,
                    warnings,
                    out error);
                if (shape == null)
                {
                    if (string.IsNullOrEmpty(error))
                    {
                        error = "PPT 插入改包页失败";
                    }

                    return false;
                }

                PourLog(warnings, "改包页已插入 " + DescribeLiveSeries(TryGetChart(shape)));
                PptHtmlApplyHangLog.Mark("旁路插入返回，记住本进程已建图");
                RememberChartDeck(app, destPres);
                return true;
            }
            catch (Exception ex)
            {
                error = "PPT 旁路建图失败: " + ex.Message;
                PourLog(warnings, error);
                shape = null;
                return false;
            }
            finally
            {
                WordAddIn1.ChartOoxmlHost.PptChartPackageHost.TryDeleteFile(tempPath);
                if (shape == null)
                {
                    DismissChartExcelUi();
                }
            }
        }

        private static int _chartPid;
        private static int _chartSlideIndex = 1;
        private static object _chartPres;
        private static bool _chartUsed;

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(int hWnd, out uint processId);

        private static void RememberChartDeck(object app, object pres)
        {
            int pid = TryAppPid(app);
            if (pid > 0 && pres != null)
            {
                _chartUsed = true;
                _chartPid = pid;
                _chartPres = pres;
                try
                {
                    object slide = WppCom.GetProperty(pres, "Slides");
                    int count = Convert.ToInt32(WppCom.GetProperty(slide, "Count") ?? 1);
                    if (count > 0)
                    {
                        _chartSlideIndex = count;
                    }
                }
                catch (Exception)
                {
                }
            }
        }

        /// <summary>
        /// 这个进程已经建过图，又打开了另一份稿：保存后结束进程，用新进程打开，再继续。
        /// 同一份还开着的稿不重开。
        /// </summary>
        private static bool TryRestartPowerPointWhenDeckChanged(
            ref object destSlide,
            ref object destPres,
            ref object app,
            ref object destShapes,
            List<string> warnings,
            out string error)
        {
            error = null;
            int pid = TryAppPid(app);
            bool same = SameOpenDeck(_chartPres, destPres);
            bool skip = !_chartUsed || pid <= 0 || pid != _chartPid || same;
            PptHtmlApplyHangLog.Mark(
                "重开判断 used=" + _chartUsed
                + " pid=" + pid
                + " chartPid=" + _chartPid
                + " same=" + same
                + " skip=" + skip);
            if (skip)
            {
                return true;
            }

            // 暂时停用关旧进程。原进程不断，新稿也不关，用来看换稿建图会不会再断。恢复时删掉下面两行和 #if false / #endif。
            PptHtmlApplyHangLog.Mark("重开机制已停用，原进程和新稿都保持打开");
            return true;
#if false
            if (pid == Process.GetCurrentProcess().Id)
            {
                error = "PPT 旧进程换稿：代码跑在 PowerPoint 内部，无法结束自身后再开";
                return false;
            }

            int slideIndex = 1;
            try
            {
                slideIndex = Convert.ToInt32(WppCom.GetProperty(destSlide, "SlideIndex") ?? 1);
            }
            catch (Exception)
            {
            }

            PptHtmlApplyHangLog.Mark("重开保存前");
            string fullPath;
            try
            {
                string dir = Convert.ToString(WppCom.GetProperty(destPres, "Path") ?? "");
                if (string.IsNullOrEmpty(dir))
                {
                    fullPath = PptHtmlOoxmlIo.NewTempPath("重开");
                    try
                    {
                        WppCom.Invoke(destPres, "SaveAs", fullPath, 24);
                    }
                    catch (Exception)
                    {
                        WppCom.Invoke(destPres, "SaveAs", fullPath);
                    }
                }
                else
                {
                    fullPath = Convert.ToString(WppCom.GetProperty(destPres, "FullName") ?? "");
                    WppCom.Invoke(destPres, "Save");
                }
            }
            catch (Exception ex)
            {
                error = "PPT 旧进程换稿：保存失败: " + ex.Message;
                return false;
            }

            if (string.IsNullOrEmpty(fullPath) || !File.Exists(fullPath))
            {
                error = "PPT 旧进程换稿：保存后文件不在";
                return false;
            }

            try
            {
                object presentations = WppCom.GetProperty(app, "Presentations");
                int count = presentations == null
                    ? 0
                    : Convert.ToInt32(WppCom.GetProperty(presentations, "Count") ?? 0);
                for (int i = 1; i <= count; i++)
                {
                    object one = WppCom.GetIndexed(presentations, i);
                    if (one == null || ReferenceEquals(one, destPres))
                    {
                        continue;
                    }

                    string oneDir = Convert.ToString(WppCom.GetProperty(one, "Path") ?? "");
                    if (!string.IsNullOrEmpty(oneDir))
                    {
                        try
                        {
                            WppCom.Invoke(one, "Save");
                        }
                        catch (Exception)
                        {
                        }
                    }
                }
            }
            catch (Exception)
            {
            }

            PptHtmlApplyHangLog.Mark("重开保存后，关闭当前稿前 " + fullPath);
            try
            {
                WppCom.TrySetProperty(destPres, "Saved", true);
                WppCom.Invoke(destPres, "Close");
            }
            catch (Exception)
            {
            }

            PourLog(warnings, "旧进程已建过图，结束 pid=" + pid + " 后用新进程打开 " + fullPath);
            PptHtmlApplyHangLog.Mark("重开杀进程前 pid=" + pid);
            if (!WaitUntilPidGone(pid))
            {
                error = "PPT 旧进程换稿：进程没有退出 pid=" + pid;
                return false;
            }

            PptHtmlApplyHangLog.Mark("重开进程已退出，新进程打开前");
            try
            {
                var restarted = new Microsoft.Office.Interop.PowerPoint.Application();
                restarted.Visible = Microsoft.Office.Core.MsoTriState.msoTrue;
                Microsoft.Office.Interop.PowerPoint.Presentation opened = restarted.Presentations.Open(
                    fullPath,
                    Microsoft.Office.Core.MsoTriState.msoFalse,
                    Microsoft.Office.Core.MsoTriState.msoFalse,
                    Microsoft.Office.Core.MsoTriState.msoTrue);
                Microsoft.Office.Interop.PowerPoint.Slide slide = opened.Slides[slideIndex];
                destPres = opened;
                destSlide = slide;
                destShapes = slide.Shapes;
                app = restarted;
                _chartPres = opened;
                _chartSlideIndex = slideIndex;
                _chartPid = TryAppPid(restarted);
                PptHtmlApplyHangLog.Mark("重开打开后 pid=" + _chartPid);
                return true;
            }
            catch (Exception ex)
            {
                error = "PPT 旧进程换稿：新进程打开失败: " + ex.Message;
                return false;
            }
#endif
        }

        private static bool IsDeckAlive(object pres)
        {
            if (pres == null)
            {
                return false;
            }

            try
            {
                WppCom.GetProperty(pres, "Name");
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        private static bool TryAdoptReopenedDeck(
            ref object destSlide,
            ref object destPres,
            ref object app,
            ref object destShapes)
        {
            try
            {
                object slides = WppCom.GetProperty(_chartPres, "Slides");
                object slide = WppCom.GetIndexed(slides, _chartSlideIndex < 1 ? 1 : _chartSlideIndex);
                object shapes = slide == null ? null : WppCom.GetProperty(slide, "Shapes");
                object application = WppCom.GetProperty(_chartPres, "Application");
                if (slide == null || shapes == null || application == null)
                {
                    return false;
                }

                destPres = _chartPres;
                destSlide = slide;
                destShapes = shapes;
                app = application;
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        private static bool SameOpenDeck(object left, object right)
        {
            if (left == null || right == null)
            {
                return false;
            }

            if (ReferenceEquals(left, right))
            {
                return true;
            }

            int leftHwnd = TryPresWindowHwnd(left);
            int rightHwnd = TryPresWindowHwnd(right);
            if (leftHwnd > 0 && rightHwnd > 0)
            {
                return leftHwnd == rightHwnd;
            }

            try
            {
                string leftName = Convert.ToString(WppCom.GetProperty(left, "FullName") ?? "");
                string rightName = Convert.ToString(WppCom.GetProperty(right, "FullName") ?? "");
                return leftName.Length > 0 && string.Equals(leftName, rightName, StringComparison.OrdinalIgnoreCase);
            }
            catch (Exception)
            {
                return false;
            }
        }

        private static int TryAppPid(object app)
        {
            try
            {
                int hwnd = 0;
                Microsoft.Office.Interop.PowerPoint.Application typed =
                    app as Microsoft.Office.Interop.PowerPoint.Application;
                if (typed != null)
                {
                    hwnd = typed.HWND;
                }
                else
                {
                    hwnd = Convert.ToInt32(WppCom.GetProperty(app, "HWND") ?? 0);
                }

                if (hwnd != 0)
                {
                    uint pid;
                    GetWindowThreadProcessId(hwnd, out pid);
                    if (pid > 0)
                    {
                        return (int)pid;
                    }
                }
            }
            catch (Exception)
            {
            }

            try
            {
                Process[] found = Process.GetProcessesByName("POWERPNT");
                if (found != null && found.Length == 1)
                {
                    return found[0].Id;
                }
            }
            catch (Exception)
            {
            }

            return 0;
        }

        private static int TryPresWindowHwnd(object pres)
        {
            try
            {
                object windows = WppCom.GetProperty(pres, "Windows");
                int count = windows == null
                    ? 0
                    : Convert.ToInt32(WppCom.GetProperty(windows, "Count") ?? 0);
                if (count < 1)
                {
                    return 0;
                }

                object window = WppCom.GetIndexed(windows, 1);
                return Convert.ToInt32(WppCom.GetProperty(window, "HWND") ?? 0);
            }
            catch (Exception)
            {
                return 0;
            }
        }

        private static bool WaitUntilPidGone(int pid)
        {
            Process process = null;
            try
            {
                process = Process.GetProcessById(pid);
            }
            catch (Exception)
            {
                return true;
            }

            try
            {
                if (!process.HasExited)
                {
                    process.Kill();
                }

                return process.HasExited || process.WaitForExit(15000);
            }
            catch (Exception)
            {
                try
                {
                    return Process.GetProcessById(pid).HasExited;
                }
                catch (Exception)
                {
                    return true;
                }
            }
        }

        private static bool TryEnsureEmbeddedChart(
            string path,
            int slideIndex,
            float left,
            float top,
            float width,
            float height,
            PptHtmlChartGrid grid,
            List<string> warnings,
            out int shapeId,
            out string error)
        {
            shapeId = 0;
            error = null;
            string slidePart = "ppt/slides/slide" + slideIndex.ToString(CultureInfo.InvariantCulture) + ".xml";
            string relsPart = "ppt/slides/_rels/slide" + slideIndex.ToString(CultureInfo.InvariantCulture) + ".xml.rels";
            try
            {
                using (ZipArchive zip = ZipFile.Open(path, ZipArchiveMode.Update))
                {
                    ZipArchiveEntry slideEntry = FindZipEntry(zip, slidePart);
                    ZipArchiveEntry relsEntry = FindZipEntry(zip, relsPart);
                    if (slideEntry == null || relsEntry == null)
                    {
                        error = "PPT 旁路建图：副本没有 " + slidePart;
                        return false;
                    }

                    XDocument slideDoc = LoadZipXml(slideEntry);
                    XDocument relsDoc = LoadZipXml(relsEntry);
                    string relId = FindChartRelId(slideDoc, out shapeId);
                    if (shapeId < 1)
                    {
                        shapeId = NextShapeId(slideDoc);
                        relId = NextRelId(relsDoc);
                        int chartNo = NextChartNumber(zip);
                        string chartPart = "ppt/charts/chart" + chartNo.ToString(CultureInfo.InvariantCulture) + ".xml";
                        string xlsxPart = "ppt/embeddings/ew-chart" + chartNo.ToString(CultureInfo.InvariantCulture) + ".xlsx";
                        AddChartFrame(slideDoc, shapeId, relId, left, top, width, height);
                        AddRelationship(relsDoc, relId, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart", "../charts/chart" + chartNo.ToString(CultureInfo.InvariantCulture) + ".xml");
                        WriteZipBytes(zip, chartPart, XmlUtf8(MinimalColumnChartXml()));
                        WriteZipBytes(zip, ChartRelsPart(chartPart), XmlUtf8(EmbeddingRelsXml("rId1", "../embeddings/ew-chart" + chartNo.ToString(CultureInfo.InvariantCulture) + ".xlsx")));
                        WriteZipBytes(zip, xlsxPart, BuildMinimalChartWorkbook(grid));
                        EnsureContentTypes(zip, chartPart, xlsxPart);
                        ReplaceZipXml(zip, slidePart, slideDoc);
                        ReplaceZipXml(zip, relsPart, relsDoc);
                        PourLog(warnings, "副本无图，已写入内嵌图 shapeId=" + shapeId);
                        return true;
                    }

                    string chartTarget = FindRelTarget(relsDoc, relId);
                    string chartPartExisting = NormalizeChartPartPath(chartTarget);
                    if (!ChartHasEmbedding(zip, chartPartExisting))
                    {
                        int chartNo = NextChartNumber(zip);
                        string xlsxPart = "ppt/embeddings/ew-chart" + chartNo.ToString(CultureInfo.InvariantCulture) + ".xlsx";
                        string embedRel = EnsurePackageRel(zip, chartPartExisting, xlsxPart);
                        PointChartAtEmbedding(zip, chartPartExisting, embedRel);
                        WriteZipBytes(zip, xlsxPart, BuildMinimalChartWorkbook(grid));
                        EnsureContentTypes(zip, null, xlsxPart);
                        PourLog(warnings, "副本图无内嵌表，已补上并去掉外链 shapeId=" + shapeId);
                    }

                    return true;
                }
            }
            catch (Exception ex)
            {
                error = "PPT 旁路建图：写内嵌图失败: " + ex.Message;
                return false;
            }
        }

        private static object TryInsertChartSlide(
            object destPres,
            object destShapes,
            int destIndex,
            string path,
            float left,
            float top,
            float width,
            float height,
            List<string> warnings,
            out string error)
        {
            error = null;
            object slides = WppCom.GetProperty(destPres, "Slides");
            int before = Convert.ToInt32(WppCom.GetProperty(slides, "Count") ?? 0);
            PptHtmlApplyHangLog.Mark("InsertFromFile 前 index=" + destIndex + " pages=" + before);
            object inserted = WppCom.Invoke(slides, "InsertFromFile", path, destIndex, destIndex, destIndex);
            PptHtmlApplyHangLog.Mark("InsertFromFile 返回");
            int after = Convert.ToInt32(WppCom.GetProperty(slides, "Count") ?? 0);
            if (after <= before)
            {
                error = "InsertFromFile 没有插入页";
                return null;
            }

            int newIndex = destIndex + 1;
            try
            {
                object first = inserted == null || inserted is int || inserted is string
                    ? null
                    : TryGetIndexedSafe(inserted, 1);
                if (first != null)
                {
                    newIndex = Convert.ToInt32(WppCom.GetProperty(first, "SlideIndex") ?? newIndex);
                }
            }
            catch (Exception)
            {
            }

            PourLog(warnings, "InsertFromFile 后页 " + before + "→" + after + " 取页=" + newIndex);

            object newSlide = WppCom.GetIndexed(slides, newIndex);
            object newShapes = newSlide == null ? null : WppCom.GetProperty(newSlide, "Shapes");
            object chartShape = TryFindChartShape(newShapes);
            if (chartShape == null)
            {
                TryDeleteSlide(newSlide);
                error = "插入页上没有图表";
                return null;
            }

            PptHtmlApplyHangLog.Mark("Copy 前");
            WppCom.Invoke(chartShape, "Copy");
            PptHtmlApplyHangLog.Mark("Paste 前");
            object pasted = TryPasteChart(destShapes, warnings);
            PptHtmlApplyHangLog.Mark("Paste 返回 pasted=" + (pasted == null ? "null" : "ok"));
            PptHtmlApplyHangLog.Mark("删插入页前");
            TryDeleteSlide(newSlide);
            PptHtmlApplyHangLog.Mark("删插入页后");
            if (pasted == null)
            {
                error = "插入页上的图未能放到当前页";
                return null;
            }

            try
            {
                WppCom.TrySetProperty(pasted, "Left", left);
                WppCom.TrySetProperty(pasted, "Top", top);
                WppCom.TrySetProperty(pasted, "Width", width);
                WppCom.TrySetProperty(pasted, "Height", height);
            }
            catch (Exception)
            {
            }

            return pasted;
        }

        private static object TryFindChartShape(object shapes)
        {
            if (shapes == null)
            {
                return null;
            }

            int count = Convert.ToInt32(WppCom.GetProperty(shapes, "Count") ?? 0);
            for (int i = 1; i <= count; i++)
            {
                object shape = WppCom.GetIndexed(shapes, i);
                try
                {
                    object has = WppCom.GetProperty(shape, "HasChart");
                    if (has != null && Convert.ToInt32(has) != 0)
                    {
                        return shape;
                    }
                }
                catch (Exception)
                {
                }
            }

            return null;
        }

        private static void TryDeleteSlide(object slide)
        {
            if (slide == null)
            {
                return;
            }

            try
            {
                WppCom.Invoke(slide, "Delete");
            }
            catch (Exception)
            {
            }
        }

        private static ZipArchiveEntry FindZipEntry(ZipArchive zip, string part)
        {
            string want = part.Replace('\\', '/');
            foreach (ZipArchiveEntry entry in zip.Entries)
            {
                if (string.Equals(entry.FullName.Replace('\\', '/'), want, StringComparison.OrdinalIgnoreCase))
                {
                    return entry;
                }
            }

            return null;
        }

        private static XDocument LoadZipXml(ZipArchiveEntry entry)
        {
            using (Stream stream = entry.Open())
            {
                return XDocument.Load(stream);
            }
        }

        private static void ReplaceZipXml(ZipArchive zip, string part, XDocument doc)
        {
            ZipArchiveEntry old = FindZipEntry(zip, part);
            if (old != null)
            {
                old.Delete();
            }

            WriteZipBytes(zip, part, XmlUtf8(doc));
        }

        private static void WriteZipBytes(ZipArchive zip, string part, byte[] data)
        {
            ZipArchiveEntry old = FindZipEntry(zip, part);
            if (old != null)
            {
                old.Delete();
            }

            ZipArchiveEntry entry = zip.CreateEntry(part.Replace('\\', '/'));
            using (Stream stream = entry.Open())
            {
                stream.Write(data, 0, data.Length);
            }
        }

        private static byte[] XmlUtf8(XDocument doc)
        {
            using (var ms = new MemoryStream())
            {
                doc.Save(ms);
                return ms.ToArray();
            }
        }

        private static byte[] XmlUtf8(string xml)
        {
            return Encoding.UTF8.GetBytes(xml);
        }

        private static string FindChartRelId(XDocument slideDoc, out int shapeId)
        {
            shapeId = 0;
            foreach (XElement frame in slideDoc.Descendants(PNs + "graphicFrame"))
            {
                XElement chartEl = frame.Descendants(CNs + "chart").FirstOrDefault();
                XElement cNvPr = frame.Descendants(PNs + "cNvPr").FirstOrDefault();
                if (chartEl == null || cNvPr == null)
                {
                    continue;
                }

                if (int.TryParse((string)cNvPr.Attribute("id"), NumberStyles.Integer, CultureInfo.InvariantCulture, out int id))
                {
                    shapeId = id;
                }

                return (string)chartEl.Attribute(RNs + "id");
            }

            return null;
        }

        private static int NextShapeId(XDocument slideDoc)
        {
            int max = 1;
            foreach (XElement cNvPr in slideDoc.Descendants(PNs + "cNvPr"))
            {
                if (int.TryParse((string)cNvPr.Attribute("id"), NumberStyles.Integer, CultureInfo.InvariantCulture, out int id) && id > max)
                {
                    max = id;
                }
            }

            return max + 1;
        }

        private static string NextRelId(XDocument relsDoc)
        {
            int max = 0;
            foreach (XElement rel in relsDoc.Descendants(RelPkgNs + "Relationship"))
            {
                string id = (string)rel.Attribute("Id") ?? "";
                if (id.StartsWith("rId", StringComparison.Ordinal)
                    && int.TryParse(id.Substring(3), NumberStyles.Integer, CultureInfo.InvariantCulture, out int n)
                    && n > max)
                {
                    max = n;
                }
            }

            return "rId" + (max + 1).ToString(CultureInfo.InvariantCulture);
        }

        private static int NextChartNumber(ZipArchive zip)
        {
            int max = 0;
            foreach (ZipArchiveEntry entry in zip.Entries)
            {
                string name = entry.FullName.Replace('\\', '/');
                if (name.StartsWith("ppt/charts/chart", StringComparison.OrdinalIgnoreCase)
                    && name.EndsWith(".xml", StringComparison.OrdinalIgnoreCase)
                    && name.IndexOf("/_rels/", StringComparison.OrdinalIgnoreCase) < 0)
                {
                    string digits = name.Substring("ppt/charts/chart".Length);
                    digits = digits.Substring(0, digits.Length - 4);
                    if (int.TryParse(digits, NumberStyles.Integer, CultureInfo.InvariantCulture, out int n) && n > max)
                    {
                        max = n;
                    }
                }
            }

            return max + 1;
        }

        private static void AddChartFrame(
            XDocument slideDoc,
            int shapeId,
            string relId,
            float left,
            float top,
            float width,
            float height)
        {
            XElement tree = slideDoc.Descendants(PNs + "spTree").FirstOrDefault();
            if (tree == null)
            {
                throw new InvalidOperationException("slide 没有 spTree");
            }

            long x = (long)(left * 12700f);
            long y = (long)(top * 12700f);
            long cx = (long)(width * 12700f);
            long cy = (long)(height * 12700f);
            var frame = new XElement(PNs + "graphicFrame",
                new XElement(PNs + "nvGraphicFramePr",
                    new XElement(PNs + "cNvPr",
                        new XAttribute("id", shapeId.ToString(CultureInfo.InvariantCulture)),
                        new XAttribute("name", "Chart " + shapeId.ToString(CultureInfo.InvariantCulture))),
                    new XElement(PNs + "cNvGraphicFramePr",
                        new XElement(ANs + "graphicFrameLocks", new XAttribute("noGrp", "1"))),
                    new XElement(PNs + "nvPr")),
                new XElement(PNs + "xfrm",
                    new XElement(ANs + "off",
                        new XAttribute("x", x.ToString(CultureInfo.InvariantCulture)),
                        new XAttribute("y", y.ToString(CultureInfo.InvariantCulture))),
                    new XElement(ANs + "ext",
                        new XAttribute("cx", cx.ToString(CultureInfo.InvariantCulture)),
                        new XAttribute("cy", cy.ToString(CultureInfo.InvariantCulture)))),
                new XElement(ANs + "graphic",
                    new XElement(ANs + "graphicData",
                        new XAttribute("uri", "http://schemas.openxmlformats.org/drawingml/2006/chart"),
                        new XElement(CNs + "chart", new XAttribute(RNs + "id", relId)))));
            XElement ext = tree.Element(PNs + "extLst");
            if (ext != null)
            {
                ext.AddBeforeSelf(frame);
            }
            else
            {
                tree.Add(frame);
            }
        }

        private static void AddRelationship(XDocument relsDoc, string id, string type, string target)
        {
            XElement root = relsDoc.Root;
            if (root == null)
            {
                throw new InvalidOperationException("rels 空");
            }

            root.Add(new XElement(RelPkgNs + "Relationship",
                new XAttribute("Id", id),
                new XAttribute("Type", type),
                new XAttribute("Target", target)));
        }

        private static bool ChartHasEmbedding(ZipArchive zip, string chartPart)
        {
            if (string.IsNullOrEmpty(chartPart))
            {
                return false;
            }

            ZipArchiveEntry rels = FindZipEntry(zip, ChartRelsPart(chartPart));
            if (rels == null)
            {
                return false;
            }

            return !string.IsNullOrEmpty(FindEmbeddingTarget(LoadZipXml(rels)));
        }

        private static string EnsurePackageRel(ZipArchive zip, string chartPart, string xlsxPart)
        {
            string relsPart = ChartRelsPart(chartPart);
            ZipArchiveEntry relsEntry = FindZipEntry(zip, relsPart);
            XDocument relsDoc = relsEntry == null
                ? new XDocument(new XElement(RelPkgNs + "Relationships"))
                : LoadZipXml(relsEntry);
            string target = "../embeddings/" + Path.GetFileName(xlsxPart);
            string id = NextRelId(relsDoc);
            AddRelationship(
                relsDoc,
                id,
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/package",
                target);
            foreach (XElement rel in relsDoc.Descendants(RelPkgNs + "Relationship").ToList())
            {
                string type = (string)rel.Attribute("Type") ?? "";
                string relTarget = (string)rel.Attribute("Target") ?? "";
                if (string.Equals((string)rel.Attribute("Id"), id, StringComparison.Ordinal))
                {
                    continue;
                }

                if (type.IndexOf("oleObject", StringComparison.OrdinalIgnoreCase) >= 0
                    && !relTarget.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase)
                    && !relTarget.EndsWith(".xls", StringComparison.OrdinalIgnoreCase))
                {
                    rel.Remove();
                }
            }

            ReplaceZipXml(zip, relsPart, relsDoc);
            return id;
        }

        private static void PointChartAtEmbedding(ZipArchive zip, string chartPart, string relId)
        {
            ZipArchiveEntry entry = FindZipEntry(zip, chartPart);
            if (entry == null || string.IsNullOrEmpty(relId))
            {
                return;
            }

            XDocument doc = LoadZipXml(entry);
            XElement data = doc.Descendants(CNs + "externalData").FirstOrDefault();
            if (data == null)
            {
                XElement chart = doc.Descendants(CNs + "chart").FirstOrDefault();
                if (chart == null)
                {
                    return;
                }

                data = new XElement(CNs + "externalData",
                    new XAttribute(RNs + "id", relId),
                    new XElement(CNs + "autoUpdate", new XAttribute("val", "0")));
                chart.Add(data);
            }
            else
            {
                data.SetAttributeValue(RNs + "id", relId);
            }

            ReplaceZipXml(zip, chartPart, doc);
        }

        private static void EnsureContentTypes(ZipArchive zip, string chartPart, string xlsxPart)
        {
            const string typesPart = "[Content_Types].xml";
            ZipArchiveEntry entry = FindZipEntry(zip, typesPart);
            if (entry == null)
            {
                return;
            }

            XNamespace ct = "http://schemas.openxmlformats.org/package/2006/content-types";
            XDocument doc = LoadZipXml(entry);
            XElement root = doc.Root;
            if (root == null)
            {
                return;
            }

            bool hasXlsx = root.Elements(ct + "Default")
                .Any(e => string.Equals((string)e.Attribute("Extension"), "xlsx", StringComparison.OrdinalIgnoreCase));
            if (!hasXlsx)
            {
                root.Add(new XElement(ct + "Default",
                    new XAttribute("Extension", "xlsx"),
                    new XAttribute("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")));
            }

            if (!string.IsNullOrEmpty(chartPart))
            {
                string partName = "/" + chartPart.Replace('\\', '/');
                bool hasChart = root.Elements(ct + "Override")
                    .Any(e => string.Equals((string)e.Attribute("PartName"), partName, StringComparison.OrdinalIgnoreCase));
                if (!hasChart)
                {
                    root.Add(new XElement(ct + "Override",
                        new XAttribute("PartName", partName),
                        new XAttribute("ContentType", "application/vnd.openxmlformats-officedocument.drawingml.chart+xml")));
                }
            }

            ReplaceZipXml(zip, typesPart, doc);
        }

        private static string MinimalColumnChartXml()
        {
            return "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
                + "<c:chartSpace xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">"
                + "<c:chart><c:plotArea><c:barChart><c:barDir val=\"col\"/><c:grouping val=\"clustered\"/>"
                + "<c:ser><c:idx val=\"0\"/><c:order val=\"0\"/>"
                + "<c:tx><c:strRef><c:f>Sheet1!$B$1</c:f><c:strCache><c:ptCount val=\"1\"/><c:pt idx=\"0\"><c:v>系列</c:v></c:pt></c:strCache></c:strRef></c:tx>"
                + "<c:cat><c:strRef><c:f>Sheet1!$A$2:$A$2</c:f><c:strCache><c:ptCount val=\"1\"/><c:pt idx=\"0\"><c:v></c:v></c:pt></c:strCache></c:strRef></c:cat>"
                + "<c:val><c:numRef><c:f>Sheet1!$B$2:$B$2</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val=\"1\"/><c:pt idx=\"0\"><c:v>0</c:v></c:pt></c:numCache></c:numRef></c:val>"
                + "</c:ser><c:axId val=\"1\"/><c:axId val=\"2\"/></c:barChart>"
                + "<c:catAx><c:axId val=\"1\"/><c:scaling><c:orientation val=\"minMax\"/></c:scaling><c:delete val=\"0\"/><c:axPos val=\"b\"/><c:crossAx val=\"2\"/></c:catAx>"
                + "<c:valAx><c:axId val=\"2\"/><c:scaling><c:orientation val=\"minMax\"/></c:scaling><c:delete val=\"0\"/><c:axPos val=\"l\"/><c:crossAx val=\"1\"/></c:valAx>"
                + "</c:plotArea><c:plotVisOnly val=\"1\"/></c:chart>"
                + "<c:externalData r:id=\"rId1\"><c:autoUpdate val=\"0\"/></c:externalData>"
                + "</c:chartSpace>";
        }

        private static string EmbeddingRelsXml(string id, string target)
        {
            return "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
                + "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
                + "<Relationship Id=\"" + id + "\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/package\" Target=\"" + target + "\"/>"
                + "</Relationships>";
        }

        private static object TryAddChartOnShapes(
            object shapes,
            float left,
            float top,
            float width,
            float height,
            int xlType,
            int chartStyle,
            bool newLayout,
            List<string> warnings,
            out string error)
        {
            error = null;
            try
            {
                if (newLayout)
                {
                    try
                    {
                        return WppCom.Invoke(shapes, "AddChart2", chartStyle, xlType, left, top, width, height, newLayout);
                    }
                    catch (Exception ex1)
                    {
                        PourLog(warnings, "AddChart2(newLayout) 失败: " + ex1.Message);
                    }
                }

                try
                {
                    return WppCom.Invoke(shapes, "AddChart2", chartStyle, xlType, left, top, width, height);
                }
                catch (Exception ex2)
                {
                    PourLog(warnings, "AddChart2 失败: " + ex2.Message);
                    return WppCom.Invoke(shapes, "AddChart", xlType, left, top, width, height);
                }
            }
            catch (Exception ex)
            {
                error = "创建图表失败: " + ex.Message;
                return null;
            }
        }

        private static bool TrySavePresentationTo(object presentation, string path, List<string> warnings)
        {
            if (presentation == null || string.IsNullOrEmpty(path))
            {
                return false;
            }

            try
            {
                WppCom.Invoke(presentation, "Save");
                if (File.Exists(path))
                {
                    return true;
                }
            }
            catch (Exception ex)
            {
                PourLog(warnings, "旁路 Save 失败: " + FormatComError(ex));
            }

            string copyErr;
            if (!WordAddIn1.ChartOoxmlHost.PptChartPackageHost.TrySaveCopy(presentation, path, out copyErr))
            {
                PourLog(warnings, "旁路 SaveCopy 失败: " + copyErr);
                return false;
            }

            return File.Exists(path);
        }

        private static object TryGetChart(object shape)
        {
            if (shape == null)
            {
                return null;
            }

            try
            {
                object has = WppCom.GetProperty(shape, "HasChart");
                if (has != null && !IsTruthy(has))
                {
                    return null;
                }
            }
            catch (Exception)
            {
            }

            try
            {
                return WppCom.GetProperty(shape, "Chart");
            }
            catch (Exception)
            {
                return null;
            }
        }

        private static bool TryReadGrid(object chart, out PptHtmlChartGrid grid, out string error)
        {
            grid = null;
            error = null;
            // 先 Series。PPT 禁止 Activate；WPP 从页上再取的 Chart 常被误判成 PPT（Application.Name=PowerPoint），
            // 系列失败后仍走 OOXML 读 c:pt，不开表。
            if (TryReadGridFromSeries(chart, out grid, out error) && grid != null && grid.IsPourable)
            {
                return true;
            }

            string seriesErr = error;
            bool avoidSheet = HostAvoidsChartDataCom(chart);
            if (avoidSheet)
            {
                if (TryReadGridFromOoxmlCopy(chart, out grid, out error) && grid != null && grid.IsPourable)
                {
                    return true;
                }

                if (!string.IsNullOrEmpty(seriesErr))
                {
                    error = seriesErr;
                }

                return false;
            }

            if (TryWakeAndReadSeries(chart, out grid, out error) && grid != null && grid.IsPourable)
            {
                return true;
            }

            if (!string.IsNullOrEmpty(error))
            {
                seriesErr = error;
            }

            // 系列仍没有：未外链再开内嵌表（外链会弹「链接的文件不可用」）。
            string embeddedErr = null;
            bool fromSheet = !IsChartDataLinked(chart)
                && TryReadGridFromEmbeddedSheetAuto(chart, out grid, out embeddedErr)
                && grid != null
                && grid.IsPourable;
            ReleaseEmbeddedChartWorkbook(chart, null);
            if (fromSheet)
            {
                return true;
            }

            if (TryReadGridFromOoxmlCopy(chart, out grid, out error) && grid != null && grid.IsPourable)
            {
                return true;
            }

            if (!string.IsNullOrEmpty(seriesErr))
            {
                error = seriesErr;
            }
            else if (!string.IsNullOrEmpty(embeddedErr))
            {
                error = embeddedErr;
            }
            else if (string.IsNullOrEmpty(error))
            {
                error = "无法读取图表数据";
            }

            return false;
        }

        private static bool TryWakeAndReadSeries(
            object chart,
            out PptHtmlChartGrid grid,
            out string error)
        {
            grid = null;
            error = null;
            try
            {
                object chartData = WppCom.GetProperty(chart, "ChartData");
                TryActivateChartData(chartData);
                PumpChartUi(120);
            }
            catch (Exception)
            {
            }

            bool ok = TryReadGridFromSeries(chart, out grid, out error)
                && grid != null
                && grid.IsPourable;
            DismissChartExcelUiForChart(chart);
            return ok;
        }

        private static bool TryReadGridFromSeries(object chart, out PptHtmlChartGrid grid, out string error)
        {
            grid = null;
            if (!WordAddIn1.LateBoundChartSeries.TryRead(chart, out WordAddIn1.LateBoundChartSeries.Grid raw, out error)
                || raw == null
                || raw.Categories == null
                || raw.Series == null)
            {
                return false;
            }

            int count = raw.Series.Count;
            int useSeries = Math.Min(count, MaxCols - 1);
            var columns = new List<PptHtmlChartColumn>
            {
                new PptHtmlChartColumn { Role = "category", Name = "类别" }
            };
            var seriesValues = new List<List<string>>();
            for (int i = 0; i < useSeries; i++)
            {
                WordAddIn1.LateBoundChartSeries.Col one = raw.Series[i];
                columns.Add(new PptHtmlChartColumn
                {
                    Role = "value",
                    Name = one.Name,
                    Color = one.Color
                });
                seriesValues.Add(one.Values ?? new List<string>());
            }

            int useRows = Math.Min(raw.Categories.Count, MaxRows - 1);
            var rows = new List<List<string>>();
            for (int r = 0; r < useRows; r++)
            {
                var row = new List<string> { raw.Categories[r] };
                for (int s = 0; s < seriesValues.Count; s++)
                {
                    row.Add(r < seriesValues[s].Count ? seriesValues[s][r] : "");
                }

                rows.Add(row);
            }

            grid = new PptHtmlChartGrid
            {
                Columns = columns,
                Rows = rows,
                Truncated = raw.Categories.Count + 1 > MaxRows || count + 1 > MaxCols
            };
            return grid.IsPourable;
        }

        private static void ApplySeriesExtras(object chart, PptHtmlChartGrid grid)
        {
            try
            {
                object sc = TryInvoke(chart, "SeriesCollection");
                int si = 0;
                for (int i = 0; i < grid.Columns.Count; i++)
                {
                    PptHtmlChartColumn col = grid.Columns[i];
                    if (col.Role == "category")
                    {
                        continue;
                    }

                    si++;
                    object series = WppCom.GetIndexed(sc, si);
                    if (series == null)
                    {
                        continue;
                    }

                    if (!IsPieChart(chart) && !string.IsNullOrEmpty(col.Color))
                    {
                        try
                        {
                            object fmt = WppCom.GetProperty(series, "Format");
                            object fill = WppCom.GetProperty(fmt, "Fill");
                            if (string.Equals(col.Color, "none", StringComparison.OrdinalIgnoreCase)
                                || string.Equals(col.Color, "transparent", StringComparison.OrdinalIgnoreCase))
                            {
                                WppCom.TrySetProperty(fill, "Visible", 0);
                            }
                            else if (TryParseHexToOffice(col.Color, out int rgb))
                            {
                                WppCom.TrySetProperty(fill, "Visible", -1);
                                object fc = WppCom.GetProperty(fill, "ForeColor");
                                WppCom.TrySetProperty(fc, "RGB", rgb);
                            }
                        }
                        catch (Exception)
                        {
                        }
                    }

                    if (!string.IsNullOrWhiteSpace(col.SeriesType)
                        && TryParseType(col.SeriesType, out int xl, out _, out _))
                    {
                        WppCom.TrySetProperty(series, "ChartType", xl);
                    }

                    if (IsSeriesAxisY2(col.AxisY))
                    {
                        WppCom.TrySetProperty(series, "AxisGroup", XlSecondary);
                    }
                    else if (IsSeriesAxisY(col.AxisY))
                    {
                        WppCom.TrySetProperty(series, "AxisGroup", XlPrimary);
                    }
                }
            }
            catch (Exception)
            {
            }
        }

        private static void TryReadTitle(object chart, PptHtmlChartFormat format)
        {
            try
            {
                object has = WppCom.GetProperty(chart, "HasTitle");
                if (!IsTruthy(has))
                {
                    return;
                }

                object title = WppCom.GetProperty(chart, "ChartTitle");
                format.Title = Convert.ToString(WppCom.GetProperty(title, "Text") ?? "");
                object font = WppCom.GetProperty(title, "Font");
                if (font != null)
                {
                    object sz = WppCom.GetProperty(font, "Size");
                    if (sz != null)
                    {
                        format.TitleFontSize = Convert.ToDouble(sz).ToString("0.##", CultureInfo.InvariantCulture);
                    }

                    object bold = WppCom.GetProperty(font, "Bold");
                    if (bold != null)
                    {
                        format.TitleFontBold = IsTruthy(bold) ? "true" : "false";
                    }

                    object color = WppCom.GetProperty(font, "Color");
                    if (color != null)
                    {
                        format.TitleFontColor = OfficeRgbToHex(Convert.ToInt32(Convert.ToDouble(color)));
                    }
                }
            }
            catch (Exception)
            {
            }
        }

        private static void TryReadLegend(object chart, PptHtmlChartFormat format)
        {
            try
            {
                object has = WppCom.GetProperty(chart, "HasLegend");
                if (!IsTruthy(has))
                {
                    format.Legend = "none";
                    return;
                }

                object legend = WppCom.GetProperty(chart, "Legend");
                object pos = WppCom.GetProperty(legend, "Position");
                format.Legend = LegendFromXl(Convert.ToInt32(pos));
            }
            catch (Exception)
            {
            }
        }

        private static void TryReadPlotAndLabels(object chart, PptHtmlChartFormat format)
        {
            try
            {
                object plot = WppCom.GetProperty(chart, "PlotArea");
                object fill = WppCom.GetProperty(WppCom.GetProperty(plot, "Format"), "Fill");
                if (!IsTruthy(WppCom.GetProperty(fill, "Visible")))
                {
                    return;
                }

                object fc = WppCom.GetProperty(fill, "ForeColor");
                object type = WppCom.GetProperty(fc, "Type");
                // msoColorTypeRGB = 1；主题/自动色读出来常是 0，不能当成 #000000
                if (type != null && Convert.ToInt32(type) != 1)
                {
                    return;
                }

                object rgb = WppCom.GetProperty(fc, "RGB");
                if (rgb == null)
                {
                    return;
                }

                int rgbVal = Convert.ToInt32(Convert.ToDouble(rgb));
                if (rgbVal == 0 && type == null)
                {
                    return;
                }

                format.PlotColor = OfficeRgbToHex(rgbVal);
            }
            catch (Exception)
            {
            }
        }

        private static void TryReadAxes(object chart, PptHtmlChartFormat format)
        {
            try
            {
                object x = TryInvoke(chart, "Axes", XlCategory, XlPrimary);
                if (x != null && IsTruthy(WppCom.GetProperty(x, "HasTitle")))
                {
                    object t = WppCom.GetProperty(x, "AxisTitle");
                    format.AxisX = Convert.ToString(WppCom.GetProperty(t, "Text") ?? "");
                }
            }
            catch (Exception)
            {
            }

            try
            {
                object y = TryInvoke(chart, "Axes", XlValue, XlPrimary);
                if (y != null && IsTruthy(WppCom.GetProperty(y, "HasTitle")))
                {
                    object t = WppCom.GetProperty(y, "AxisTitle");
                    format.AxisY = Convert.ToString(WppCom.GetProperty(t, "Text") ?? "");
                }
            }
            catch (Exception)
            {
            }
        }

        private static void TryReadValueScale(object chart, PptHtmlChartFormat format)
        {
            try
            {
                object y = TryInvoke(chart, "Axes", XlValue, XlPrimary);
                if (y == null)
                {
                    return;
                }

                if (!IsTruthy(WppCom.GetProperty(y, "MinimumScaleIsAuto")))
                {
                    object mn = WppCom.GetProperty(y, "MinimumScale");
                    if (mn != null)
                    {
                        format.AxisYMin = Convert.ToDouble(mn).ToString("0.##", CultureInfo.InvariantCulture);
                    }
                }

                if (!IsTruthy(WppCom.GetProperty(y, "MaximumScaleIsAuto")))
                {
                    object mx = WppCom.GetProperty(y, "MaximumScale");
                    if (mx != null)
                    {
                        format.AxisYMax = Convert.ToDouble(mx).ToString("0.##", CultureInfo.InvariantCulture);
                    }
                }

                if (!IsTruthy(WppCom.GetProperty(y, "MajorUnitIsAuto")))
                {
                    object un = WppCom.GetProperty(y, "MajorUnit");
                    if (un != null)
                    {
                        format.AxisYMajorUnit = Convert.ToDouble(un).ToString("0.##", CultureInfo.InvariantCulture);
                    }
                }
            }
            catch (Exception)
            {
            }
        }

        private static void TryReadExplosion(object chart, PptHtmlChartFormat format)
        {
            try
            {
                // 仅各瓣相同才报整图值；不对称不输出（换数也不继承）。
                int? uniform = TryReadUniformPieExplosion(chart);
                if (uniform.HasValue)
                {
                    format.Explosion = uniform.Value.ToString(CultureInfo.InvariantCulture);
                }
            }
            catch (Exception)
            {
            }
        }

        private static void TrySetTitle(object chart, string title, List<string> warnings)
        {
            try
            {
                if (string.IsNullOrEmpty(title))
                {
                    WppCom.TrySetProperty(chart, "HasTitle", false);
                    try
                    {
                        if (IsTruthy(WppCom.GetProperty(chart, "HasTitle")))
                        {
                            object leftover = WppCom.GetProperty(chart, "ChartTitle");
                            if (leftover != null)
                            {
                                WppCom.TrySetProperty(leftover, "Text", "");
                            }

                            WppCom.TrySetProperty(chart, "HasTitle", false);
                        }
                    }
                    catch (Exception)
                    {
                    }

                    return;
                }

                WppCom.TrySetProperty(chart, "HasTitle", true);
                object ct = WppCom.GetProperty(chart, "ChartTitle");
                WppCom.TrySetProperty(ct, "Text", title);
            }
            catch (Exception ex)
            {
                warnings?.Add("未能套用 data-title: " + ex.Message);
            }
        }

        private static void TrySetOptional(object chart, PptHtmlChartFormat fmt, List<string> warnings)
        {
            if (!string.IsNullOrWhiteSpace(fmt.Theme))
            {
                try
                {
                    int style = string.Equals(fmt.Theme, "office", StringComparison.OrdinalIgnoreCase) ? 1 : 1;
                    object cur = WppCom.GetProperty(chart, "ChartStyle");
                    if (cur == null || Convert.ToInt32(cur) != style)
                    {
                        WppCom.TrySetProperty(chart, "ChartStyle", style);
                    }
                }
                catch (Exception)
                {
                    Warn(warnings, "data-theme");
                }
            }

            if (!string.IsNullOrWhiteSpace(fmt.Legend))
            {
                try
                {
                    if (string.Equals(fmt.Legend, "none", StringComparison.OrdinalIgnoreCase))
                    {
                        WppCom.TrySetProperty(chart, "HasLegend", false);
                    }
                    else
                    {
                        WppCom.TrySetProperty(chart, "HasLegend", true);
                        object legend = WppCom.GetProperty(chart, "Legend");
                        WppCom.TrySetProperty(legend, "Position", LegendToXl(fmt.Legend));
                    }
                }
                catch (Exception)
                {
                    Warn(warnings, "data-legend");
                }
            }

            if (!string.IsNullOrWhiteSpace(fmt.ShowDataLabels))
            {
                try
                {
                    bool on = IsTrue(fmt.ShowDataLabels);
                    object sc = TryInvoke(chart, "SeriesCollection");
                    int n = Convert.ToInt32(WppCom.GetProperty(sc, "Count"));
                    for (int i = 1; i <= n; i++)
                    {
                        object s = WppCom.GetIndexed(sc, i);
                        WppCom.TrySetProperty(s, "HasDataLabels", on);
                    }
                }
                catch (Exception)
                {
                    Warn(warnings, "data-show-data-labels");
                }
            }

            if (!string.IsNullOrWhiteSpace(fmt.ShowValue) || !string.IsNullOrWhiteSpace(fmt.ShowPercentage))
            {
                try
                {
                    bool? showValue = ParseOptionalBool(fmt.ShowValue);
                    bool? showPct = ParseOptionalBool(fmt.ShowPercentage);
                    object sc = TryInvoke(chart, "SeriesCollection");
                    int n = Convert.ToInt32(WppCom.GetProperty(sc, "Count"));
                    for (int i = 1; i <= n; i++)
                    {
                        object s = WppCom.GetIndexed(sc, i);
                        if (showValue == true || showPct == true)
                        {
                            object has = WppCom.GetProperty(s, "HasDataLabels");
                            if (!IsTruthy(has))
                            {
                                TryInvoke(s, "ApplyDataLabels", XlDataLabelsShowValue);
                            }
                        }

                        object dls = TryGetDataLabels(s);
                        if (dls == null)
                        {
                            continue;
                        }

                        if (showValue.HasValue)
                        {
                            WppCom.TrySetProperty(dls, "ShowValue", showValue.Value);
                        }

                        if (showPct.HasValue)
                        {
                            WppCom.TrySetProperty(dls, "ShowPercentage", showPct.Value);
                        }
                    }
                }
                catch (Exception)
                {
                    Warn(warnings, "data-show-value/percentage");
                }
            }

            if (string.Equals(fmt.PlotColor, "none", StringComparison.OrdinalIgnoreCase)
                || string.Equals(fmt.PlotColor, "transparent", StringComparison.OrdinalIgnoreCase))
            {
                try
                {
                    TryWriteAreaFill(chart, "PlotArea", false, null);
                }
                catch (Exception)
                {
                    Warn(warnings, "data-plot-color");
                }
            }
            else if (!string.IsNullOrWhiteSpace(fmt.PlotColor) && TryParseHexToOffice(fmt.PlotColor, out int plotRgb))
            {
                try
                {
                    object plot = WppCom.GetProperty(chart, "PlotArea");
                    object fill = WppCom.GetProperty(WppCom.GetProperty(plot, "Format"), "Fill");
                    if (!FillRgbAlreadyMatches(fill, plotRgb))
                    {
                        TryInvoke(fill, "Solid");
                        object fc = WppCom.GetProperty(fill, "ForeColor");
                        WppCom.TrySetProperty(fc, "RGB", plotRgb);
                    }
                }
                catch (Exception)
                {
                    Warn(warnings, "data-plot-color");
                }
            }

            if (!string.IsNullOrWhiteSpace(fmt.TitleFontSize)
                || !string.IsNullOrWhiteSpace(fmt.TitleFontBold)
                || !string.IsNullOrWhiteSpace(fmt.TitleFontColor))
            {
                try
                {
                    object title = WppCom.GetProperty(chart, "ChartTitle");
                    object font = WppCom.GetProperty(title, "Font");
                    if (!string.IsNullOrWhiteSpace(fmt.TitleFontSize)
                        && double.TryParse(fmt.TitleFontSize, NumberStyles.Float, CultureInfo.InvariantCulture, out double sz))
                    {
                        WppCom.TrySetProperty(font, "Size", sz);
                    }

                    if (!string.IsNullOrWhiteSpace(fmt.TitleFontBold))
                    {
                        WppCom.TrySetProperty(font, "Bold", IsTrue(fmt.TitleFontBold));
                    }

                    if (!string.IsNullOrWhiteSpace(fmt.TitleFontColor)
                        && TryParseHexToOffice(fmt.TitleFontColor, out int tRgb))
                    {
                        WppCom.TrySetProperty(font, "Color", tRgb);
                    }
                }
                catch (Exception)
                {
                    Warn(warnings, "data-title-font-*");
                }
            }

            TrySetAxisTitle(chart, XlCategory, XlPrimary, fmt.AxisX, warnings, "data-axis-x");
            TrySetAxisTitle(chart, XlValue, XlPrimary, fmt.AxisY, warnings, "data-axis-y");
            TrySetAxisTitle(chart, XlValue, XlSecondary, fmt.AxisYSecondary, warnings, "data-axis-y-secondary");

            if (!string.IsNullOrWhiteSpace(fmt.AxisYMin)
                || !string.IsNullOrWhiteSpace(fmt.AxisYMax)
                || !string.IsNullOrWhiteSpace(fmt.AxisYMajorUnit))
            {
                TryApplyValueAxisScaleFromFormat(chart, fmt, warnings);
            }

            if (!string.IsNullOrWhiteSpace(fmt.GapWidth)
                && int.TryParse(fmt.GapWidth, NumberStyles.Integer, CultureInfo.InvariantCulture, out int gap))
            {
                try
                {
                    object g = TryInvoke(chart, "ChartGroups", 1);
                    WppCom.TrySetProperty(g, "GapWidth", gap);
                }
                catch (Exception)
                {
                    Warn(warnings, "data-gap-width");
                }
            }

            if (!string.IsNullOrWhiteSpace(fmt.Explosion)
                && int.TryParse(fmt.Explosion, NumberStyles.Integer, CultureInfo.InvariantCulture, out int exp))
            {
                TryApplyPieExplosion(chart, exp, warnings);
            }
        }

        /// <summary>饼/环：灌数后点数已定，给当前全部瓣写同一爆炸值。</summary>
        private static void TryApplyPieExplosion(object chart, int explosion, List<string> warnings)
        {
            if (chart == null || !IsPieChart(chart))
            {
                return;
            }

            int? pieXl = null;
            if (TryReadChartXl(chart, out int cur) && IsPieXl(cur))
            {
                pieXl = cur;
            }

            try
            {
                object sc = TryInvoke(chart, "SeriesCollection");
                if (sc == null)
                {
                    sc = WppCom.GetProperty(chart, "SeriesCollection");
                }

                object s = sc == null ? null : WppCom.GetIndexed(sc, 1);
                if (s == null)
                {
                    return;
                }

                // 只写 Point.Explosion。系列级 Explosion 在部分环境会把饼图 ChartType 打成柱图。
                object pts = TryGetPoints(s);
                if (pts == null)
                {
                    // 退回逐点 Points(i)
                    int guess = 0;
                    try
                    {
                        object cnt = WppCom.GetProperty(s, "Points") ?? TryInvoke(s, "Points");
                        if (cnt != null)
                        {
                            guess = Convert.ToInt32(WppCom.GetProperty(cnt, "Count"));
                        }
                    }
                    catch (Exception)
                    {
                    }

                    if (guess < 1)
                    {
                        guess = 8;
                    }

                    for (int i = 1; i <= guess; i++)
                    {
                        object pt = TryInvoke(s, "Points", i);
                        if (pt == null)
                        {
                            break;
                        }

                        WppCom.TrySetProperty(pt, "Explosion", explosion);
                    }

                    return;
                }

                int n = Convert.ToInt32(WppCom.GetProperty(pts, "Count"));
                for (int i = 1; i <= n; i++)
                {
                    object pt = WppCom.GetIndexed(pts, i);
                    if (pt == null)
                    {
                        pt = TryInvoke(pts, "Item", i);
                    }

                    if (pt == null)
                    {
                        pt = TryInvoke(s, "Points", i);
                    }

                    if (pt != null)
                    {
                        WppCom.TrySetProperty(pt, "Explosion", explosion);
                    }
                }
            }
            catch (Exception)
            {
                Warn(warnings, "data-explosion");
            }
            finally
            {
                // 部分环境下写 Point.Explosion 会把整图 ChartType 打成柱图，须钉回。
                if (pieXl.HasValue && !IsPieChart(chart))
                {
                    try
                    {
                        WppCom.TrySetProperty(chart, "ChartType", pieXl.Value);
                    }
                    catch (Exception)
                    {
                    }
                }
            }
        }

        private static void TryApplyPieExplosionFromFormat(
            object chart,
            PptHtmlChartFormat fmt,
            List<string> warnings)
        {
            if (fmt == null || string.IsNullOrWhiteSpace(fmt.Explosion))
            {
                return;
            }

            if (!int.TryParse(fmt.Explosion, NumberStyles.Integer, CultureInfo.InvariantCulture, out int exp))
            {
                return;
            }

            TryApplyPieExplosion(chart, exp, warnings);
        }

        /// <summary>
        /// 各瓣爆炸值相同则返回该值；不对称或非饼返回 null（换数不继承不对称）。
        /// </summary>
        private static int? TryReadUniformPieExplosion(object chart)
        {
            if (chart == null || !IsPieChart(chart))
            {
                return null;
            }

            try
            {
                object sc = TryInvoke(chart, "SeriesCollection");
                if (sc == null)
                {
                    sc = WppCom.GetProperty(chart, "SeriesCollection");
                }

                object s = sc == null ? null : WppCom.GetIndexed(sc, 1);
                if (s == null)
                {
                    return null;
                }

                object pts = TryGetPoints(s);
                if (pts == null)
                {
                    // 点集合不可用时退回系列级（常为 0）
                    object rawSeries = WppCom.GetProperty(s, "Explosion");
                    return rawSeries == null ? 0 : Convert.ToInt32(Convert.ToDouble(rawSeries));
                }

                int n = Convert.ToInt32(WppCom.GetProperty(pts, "Count"));
                if (n < 1)
                {
                    return null;
                }

                int? uniform = null;
                for (int i = 1; i <= n; i++)
                {
                    object pt = WppCom.GetIndexed(pts, i)
                        ?? TryInvoke(pts, "Item", i)
                        ?? TryInvoke(s, "Points", i);
                    object raw = pt == null ? null : WppCom.GetProperty(pt, "Explosion");
                    int v = raw == null ? 0 : Convert.ToInt32(Convert.ToDouble(raw));
                    if (uniform == null)
                    {
                        uniform = v;
                    }
                    else if (uniform.Value != v)
                    {
                        return null;
                    }
                }

                return uniform;
            }
            catch (Exception)
            {
                return null;
            }
        }

        private static void TrySetAxisTitle(
            object chart,
            int axisType,
            int group,
            string title,
            List<string> warnings,
            string field)
        {
            if (title == null)
            {
                return;
            }

            try
            {
                object axis = TryInvoke(chart, "Axes", axisType, group);
                if (string.IsNullOrEmpty(title))
                {
                    WppCom.TrySetProperty(axis, "HasTitle", false);
                    return;
                }

                WppCom.TrySetProperty(axis, "HasTitle", true);
                object at = WppCom.GetProperty(axis, "AxisTitle");
                WppCom.TrySetProperty(at, "Text", title);
            }
            catch (Exception)
            {
                Warn(warnings, field);
            }
        }

    }
}
