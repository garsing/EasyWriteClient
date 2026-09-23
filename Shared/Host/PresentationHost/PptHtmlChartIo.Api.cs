using System;
using System.Collections.Generic;
using System.Globalization;
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
        public static bool LooksLikeChart(object shape)
        {
            if (shape == null)
            {
                return false;
            }

            try
            {
                return IsTruthy(WppCom.GetProperty(shape, "HasChart"));
            }
            catch (Exception)
            {
                return false;
            }
        }

        public static bool TryRead(object shape, out PptHtmlChartReadModel model, out string error)
        {
            model = null;
            error = null;
            object chart = TryGetChart(shape);
            if (chart == null)
            {
                error = "无法读取图表 Chart 对象";
                return false;
            }

            var format = new PptHtmlChartFormat();
            try
            {
                object t = WppCom.GetProperty(chart, "ChartType");
                if (t != null)
                {
                    format.ChartType = CanonicalTypeFromXl(Convert.ToInt32(t));
                }
            }
            catch (Exception)
            {
            }

            TryReadTitle(chart, format);
            TryReadLegend(chart, format);
            TryReadPlotAndLabels(chart, format);
            TryReadAxes(chart, format);
            TryReadValueScale(chart, format);
            TryReadExplosion(chart, format);
            ChartStyleSnap snap = TryCaptureStyle(chart);
            LogAxisFormats(chart, null, "读COM");
            TryEnrichSnapFromOoxml(shape, snap);
            LogAxisFormats(chart, null, "读OOXML后");

            if (!TryReadGrid(chart, out PptHtmlChartGrid grid, out error))
            {
                return false;
            }

            ProjectSnapToFormat(format, snap);
            ProjectSnapToColumns(grid, snap);
            PourLog(null, "轴格式 投影 cat=" + (format.AxisXStyle == null ? "-" : (format.AxisXStyle.Format ?? format.AxisXFormat ?? "null"))
                + " val=" + (format.AxisYStyle == null ? "-" : (format.AxisYStyle.Format ?? "null"))
                + " val2=" + (format.AxisY2Style == null ? "-" : (format.AxisY2Style.Format ?? "null")));
            model = new PptHtmlChartReadModel
            {
                Format = format,
                Grid = grid
            };
            return true;
        }

        public static bool TryPourGrid(
            object chart,
            PptHtmlChartGrid grid,
            out string error,
            List<string> warnings = null)
        {
            error = null;
            if (chart == null || grid == null || !grid.IsPourable)
            {
                error = "新建 chart 必须在节点内嵌 <table> 灌数，不能建空图";
                return false;
            }

            if (HostAvoidsChartDataCom(chart))
            {
                error = "PPT 灌数走 OOXML，不打开 ChartData";
                PourLog(warnings, error);
                return false;
            }

            PourLog(warnings, "开始灌数（写 ChartData 内嵌表） " + DescribeWantGrid(grid));
            object excelApp = null;
            try
            {
                if (!TryPourViaChartData(chart, grid, out excelApp, out error, warnings))
                {
                    PourLog(warnings, "写内嵌表失败: " + (error ?? ""));
                    return false;
                }

                int wantSeries = CountValueColumns(grid);
                if (!TryEnsureSeriesCount(chart, wantSeries, out error))
                {
                    PourLog(warnings, "扩系列失败: " + (error ?? ""));
                    return false;
                }

                TrimExtraSeries(chart, wantSeries, out _);
                if (ReadSeriesRowCount(chart) == grid.Rows.Count)
                {
                    PourLog(warnings, "系列点数已对齐，跳过数组灌入以免截成默认 4 点 | "
                        + DescribeLiveSeries(chart) + " | " + DescribeSeriesExtra(chart));
                }
                else if (!TryPourSeriesLikeWord(chart, grid, out error, warnings))
                {
                    PourLog(warnings, "Series 同步失败: " + (error ?? ""));
                    return false;
                }

                EnsureCategoryAxisLabels(chart, grid);
                PourLog(warnings, "灌数后 " + DescribeLiveSeries(chart));

                if (!TryVerifyPouredGrid(chart, grid, out error, warnings))
                {
                    PourLog(warnings, "灌后校验失败: " + (error ?? ""));
                    return false;
                }

                ApplySeriesExtras(chart, grid);
                return true;
            }
            catch (Exception ex)
            {
                error = "灌入图表数据失败: " + ex.Message;
                return false;
            }
            finally
            {
                ReleaseEmbeddedChartWorkbook(chart, excelApp, warnings);
                DismissChartExcelUiForChart(chart);
            }
        }

        public static bool TryApplyFormat(object chart, PptHtmlChartFormat fmt, List<string> warnings, out string error)
        {
            error = null;
            if (chart == null || fmt == null)
            {
                return true;
            }

            if (!string.IsNullOrWhiteSpace(fmt.ChartType))
            {
                if (!TryParseType(fmt.ChartType, out int xl, out string canon, out error))
                {
                    return false;
                }

                if (!IsSameChartType(chart, canon))
                {
                    TryCaptureTickLabelFont(chart, out object tickColor, out object tickSize);
                    try
                    {
                        WppCom.TrySetProperty(chart, "ChartType", xl);
                    }
                    catch (Exception ex)
                    {
                        error = "无法改 data-chart-type: " + ex.Message;
                        return false;
                    }

                    TryRestoreTickLabelFont(chart, tickColor, tickSize);
                }
            }

            if (fmt.Title != null)
            {
                TrySetTitle(chart, fmt.Title, warnings);
            }

            TrySetOptional(chart, fmt, warnings);
            return true;
        }

        /// <summary>
        /// 改已有图 = 新建图。拍旧图外观白名单 + HTML/表结构，AddChart2 灌数，再删旧图。
        /// </summary>
        public static bool TryReplaceOnSlide(
            object shapes,
            object oldShape,
            PptHtmlChartGrid grid,
            PptHtmlChartFormat format,
            float? left,
            float? top,
            float? width,
            float? height,
            List<string> warnings,
            out object newShape,
            out string error)
        {
            newShape = null;
            error = null;
            if (shapes == null || oldShape == null)
            {
                error = "无法定位要替换的图表";
                return false;
            }

            if (!TryReadBox(oldShape, out float oldLeft, out float oldTop, out float oldWidth, out float oldHeight))
            {
                error = "无法读取原图位置";
                return false;
            }

            // 组内 chart 的 HTML style 是相对组的 %；瘦稿平铺到 section 后会当成整页坐标
            //（常见 top:0% → 飞到页顶）。换数沿用原图幻灯片盒，不听稿上的几何。
            float useLeft;
            float useTop;
            float useWidth;
            float useHeight;
            if (IsInsideGroup(oldShape))
            {
                if (left.HasValue || top.HasValue || width.HasValue || height.HasValue)
                {
                    StyleLog(warnings, "组内 chart 忽略稿上 style，沿用原图幻灯片坐标 "
                        + oldLeft.ToString("0.#", CultureInfo.InvariantCulture) + ","
                        + oldTop.ToString("0.#", CultureInfo.InvariantCulture));
                }

                useLeft = oldLeft;
                useTop = oldTop;
                useWidth = oldWidth;
                useHeight = oldHeight;
            }
            else
            {
                useLeft = left ?? oldLeft;
                useTop = top ?? oldTop;
                useWidth = width ?? oldWidth;
                useHeight = height ?? oldHeight;
            }

            PptHtmlChartFormat useFormat = format ?? new PptHtmlChartFormat();
            bool htmlWroteType = format != null && !string.IsNullOrWhiteSpace(format.ChartType);
            bool gridFromHtml = grid != null && grid.IsPourable;
            PptHtmlChartGrid useGrid = gridFromHtml ? grid : null;
            if (useGrid == null)
            {
                if (TryRead(oldShape, out PptHtmlChartReadModel model, out _)
                    && model != null
                    && model.Grid != null
                    && model.Grid.IsPourable)
                {
                    useGrid = model.Grid;
                }
            }

            if (useGrid == null || !useGrid.IsPourable)
            {
                error = "改已有 chart 须带内嵌 <table>（将删旧图重建，避免继承外链）";
                return false;
            }

            object oldChart = TryGetChart(oldShape);

            ChartStyleSnap oldSnap = null;
            try
            {
                oldSnap = TryCaptureStyle(oldChart, warnings, "旧图");
            }
            catch (Exception ex)
            {
                StyleLog(warnings, "拍旧图样式异常: " + ex.Message);
            }

            ChartStyleSnap htmlSnap = SnapFromFormat(useFormat, useGrid);
            ChartStyleSnap snap = MergeReplaceSnap(oldSnap, htmlSnap, useFormat, warnings);
            if (snap == null)
            {
                snap = htmlSnap ?? oldSnap;
            }

            int xlType;
            string typeFrom;
            if (htmlWroteType)
            {
                if (!TryParseType(useFormat.ChartType, out xlType, out string canon, out error))
                {
                    return false;
                }

                useFormat.ChartType = canon;
                PinChartSeriesType(snap, htmlSnap, xlType);
                typeFrom = "html";
            }
            else if (TryReadChartXl(oldChart, out xlType) && IsBuildableAddChartXl(xlType))
            {
                useFormat.ChartType = CanonicalTypeFromXl(xlType);
                typeFrom = "old";
            }
            else
            {
                int oldXl = 0;
                TryReadChartXl(oldChart, out oldXl);
                if (!TryDeriveAddChartBaseType(
                    gridFromHtml,
                    useGrid,
                    oldSnap,
                    out xlType,
                    out string canon,
                    out string deriveSource,
                    out error))
                {
                    return false;
                }

                useFormat.ChartType = canon;
                PinChartSeriesType(snap, htmlSnap, xlType);
                typeFrom = "old-derived:" + deriveSource;
                if (oldXl == XlCombination)
                {
                    StyleLog(warnings, "整图 xlCombination 不可建，推导底型=" + canon + " xl=" + xlType);
                }
            }

            StyleLog(warnings, "重建开始 type=" + (useFormat.ChartType ?? "")
                + " xl=" + xlType
                + " from=" + typeFrom
                + " box=" + useLeft.ToString("0.#", CultureInfo.InvariantCulture)
                + "," + useTop.ToString("0.#", CultureInfo.InvariantCulture)
                + " " + useWidth.ToString("0.#", CultureInfo.InvariantCulture)
                + "x" + useHeight.ToString("0.#", CultureInfo.InvariantCulture));

            SyncFormatToOldSnap(useFormat, xlType, warnings);
            MergeGridlinesFromFormat(snap, useFormat);

            if (!TryCreateOnSlide(
                shapes,
                useLeft,
                useTop,
                useWidth,
                useHeight,
                xlType,
                useGrid,
                useFormat,
                warnings,
                out newShape,
                out error,
                -1,
                newLayout: false,
                applyHtmlChrome: false))
            {
                return false;
            }

            object newChart = TryGetChart(newShape);
            FinishLineChartLayout(newChart, xlType, useFormat, warnings);
            // 关标题须在 Dismiss Excel 之后：Dismiss 会把默认标题再掀开成「系列A」，套样式阶段先跳过关。
            bool deferTitleOff = snap != null
                && (snap.HasTitle == false || (snap.Title != null && snap.Title.Length == 0));
            TryApplyStyleSnap(newChart, snap, warnings, useGrid, useFormat, deferTitleOff);
            PourLog(warnings, "套快照后 " + DescribeLiveSeries(newChart));
            // 套快照（含 3D）可能把数据打回字面量；必须再验，对不上再灌一次，仍不对就失败并删新图。
            if (!TryVerifyPouredGrid(newChart, useGrid, out error, warnings))
            {
                StyleLog(warnings, "套回后数据对不上，再按新建图灌数: " + (error ?? ""));
                bool poured = false;
                if (HostAvoidsChartDataCom(newChart))
                {
                    if (TryFixSeriesViaOoxml(newShape, useGrid, warnings, out object fixedShape)
                        && fixedShape != null)
                    {
                        newShape = fixedShape;
                        newChart = TryGetChart(newShape);
                        poured = TryVerifyPouredGrid(newChart, useGrid, out error, warnings);
                    }
                    else if (string.IsNullOrEmpty(error))
                    {
                        error = "PPT 换数 OOXML 失败";
                    }
                }
                else
                {
                    poured = TryPourGrid(newChart, useGrid, out error, warnings)
                        && TryVerifyPouredGrid(newChart, useGrid, out error, warnings);
                }

                if (!poured)
                {
                    TryDelete(newShape);
                    newShape = null;
                    if (string.IsNullOrEmpty(error))
                    {
                        error = "图表换数后内嵌表与稿不一致";
                    }

                    return false;
                }
            }

            // 轴皮是最后一层：前面结构/开标签/再灌数都可能按 ChartStyle 掀字色。
            TryInheritAxisChrome(newChart, snap, warnings);
            // 换数跳过 applyHtmlChrome，尺度不在快照旧字段里时仍以稿为准再钉一次。
            TryApplyValueAxisScaleFromFormat(newChart, useFormat, warnings);
            // 再灌数会重建瓣点：统一爆炸须按当前点数最后再钉（稿优先，否则继承快照）。
            if (useFormat != null && !string.IsNullOrWhiteSpace(useFormat.Explosion))
            {
                TryApplyPieExplosionFromFormat(newChart, useFormat, warnings);
            }
            else if (snap != null && snap.Explosion.HasValue)
            {
                TryApplyPieExplosion(newChart, snap.Explosion.Value, warnings);
            }

            TryRestorePieChartType(newChart, snap);

            DismissChartExcelUiForChart(newChart);
            if (deferTitleOff)
            {
                TrySetTitle(newChart, "", warnings);
            }

            TryStripWppSrgbTintAfterChrome(ref newShape, ref newChart, snap, warnings);

            TryDelete(oldShape);
            warnings?.Add("已按旧图属性新建图表（HTML 未写的属性用快照补上）");
            if (!string.IsNullOrEmpty(EasyWriteLog.CurrentLogPath))
            {
                warnings?.Add("chart 样式日志: " + EasyWriteLog.CurrentLogPath);
            }

            return true;
        }

    }
}
