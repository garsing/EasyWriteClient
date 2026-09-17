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
        private static void ProjectSnapToFormat(PptHtmlChartFormat format, ChartStyleSnap snap)
        {
            if (format == null || snap == null)
            {
                return;
            }

            if (snap.ChartStyle.HasValue)
            {
                format.ChartStyle = snap.ChartStyle.Value.ToString(CultureInfo.InvariantCulture);
            }

            if (snap.HasLegend == false)
            {
                format.Legend = "none";
            }
            else if (snap.LegendPosition.HasValue)
            {
                format.Legend = LegendFromXl(snap.LegendPosition.Value);
            }

            if (!string.IsNullOrEmpty(snap.LegendFontColor))
            {
                format.LegendFontColor = snap.LegendFontColor;
            }

            if (!string.IsNullOrEmpty(snap.TitleFontSize))
            {
                format.TitleFontSize = snap.TitleFontSize;
            }

            if (snap.TitleFontBold.HasValue)
            {
                format.TitleFontBold = snap.TitleFontBold.Value ? "true" : "false";
            }

            if (!string.IsNullOrEmpty(snap.TitleFontColor))
            {
                format.TitleFontColor = snap.TitleFontColor;
            }

            if (snap.HasTitle == true && !string.IsNullOrEmpty(snap.Title)
                && string.IsNullOrEmpty(format.Title))
            {
                Console.WriteLine("  [标题调试] ProjectSnapToFormat 补 Title=[" + snap.Title + "]");
                format.Title = snap.Title;
            }

            format.ChartAreaColor = AreaColorFromSnap(snap.ChartAreaFillVisible, snap.ChartAreaFillRgb);
            format.PlotColor = AreaColorFromSnap(snap.PlotFillVisible, snap.PlotFillRgb);
            if (snap.GapWidth.HasValue)
            {
                format.GapWidth = snap.GapWidth.Value.ToString(CultureInfo.InvariantCulture);
            }

            if (snap.Overlap.HasValue)
            {
                format.Overlap = snap.Overlap.Value.ToString(CultureInfo.InvariantCulture);
            }

            format.PlotBox = BoxFromSnap(
                snap.PlotLeft, snap.PlotTop, snap.PlotWidth, snap.PlotHeight);
            format.PlotInside = BoxFromSnap(
                snap.PlotInsideLeft, snap.PlotInsideTop, snap.PlotInsideWidth, snap.PlotInsideHeight);
            format.AxisXStyle = AxisExtrasFromSnap(snap.Category);
            format.AxisYStyle = AxisExtrasFromSnap(snap.Value);
            format.AxisY2Style = AxisExtrasFromSnap(snap.ValueSecondary);
            if (format.AxisXStyle != null && !string.IsNullOrEmpty(format.AxisXStyle.Format))
            {
                format.AxisXFormat = format.AxisXStyle.Format;
            }

            if (snap.Category != null && snap.Category.HasTitle == true
                && !string.IsNullOrEmpty(snap.Category.Title))
            {
                format.AxisX = snap.Category.Title;
            }

            if (snap.Value != null && snap.Value.HasTitle == true
                && !string.IsNullOrEmpty(snap.Value.Title))
            {
                format.AxisY = snap.Value.Title;
            }

            if (snap.ValueSecondary != null && snap.ValueSecondary.HasTitle == true
                && !string.IsNullOrEmpty(snap.ValueSecondary.Title))
            {
                format.AxisYSecondary = snap.ValueSecondary.Title;
            }

            if (snap.Value != null && snap.Value.HasMajorGridlines.HasValue)
            {
                format.Gridlines = snap.Value.HasMajorGridlines.Value ? "true" : "false";
            }

            MarkDisplaySwitchMentions(format);
        }

        private static void ProjectSnapToColumns(PptHtmlChartGrid grid, ChartStyleSnap snap)
        {
            if (grid == null || grid.Columns == null || snap == null || snap.Series == null)
            {
                return;
            }

            int si = 0;
            for (int i = 0; i < grid.Columns.Count && si < snap.Series.Count; i++)
            {
                PptHtmlChartColumn col = grid.Columns[i];
                if (col == null || col.Role == "category")
                {
                    continue;
                }

                SeriesStyleSnap one = snap.Series[si++];
                if (one.ChartType.HasValue)
                {
                    col.SeriesType = SeriesTypeFromXl(one.ChartType.Value);
                }

                if (one.AxisGroup == XlSecondary)
                {
                    col.AxisY = "y2";
                }
                else if (one.AxisGroup == XlPrimary)
                {
                    col.AxisY = "y";
                }

                if (one.Fill != null)
                {
                    if (one.Fill.Stops != null && one.Fill.Stops.Count > 0)
                    {
                        col.FillGradient = EncodeGradient(one.Fill.Stops);
                        if (one.Fill.Angle.HasValue)
                        {
                            col.FillAngle = one.Fill.Angle.Value.ToString("0.##", CultureInfo.InvariantCulture);
                        }
                    }
                    else if (one.Fill.SolidRgb.HasValue)
                    {
                        col.Color = OfficeRgbToHex(one.Fill.SolidRgb.Value);
                    }
                    else if (one.Fill.Visible == false)
                    {
                        col.Color = "none";
                    }
                }

                if (one.Line != null)
                {
                    if (one.Line.Visible == false)
                    {
                        col.Line = "none";
                    }
                    else if (one.Line.Rgb.HasValue)
                    {
                        col.Line = OfficeRgbToHex(one.Line.Rgb.Value);
                    }

                    if (one.Line.Weight.HasValue && !IsPhantomWeight(one.Line.Weight.Value))
                    {
                        col.LineWeight = one.Line.Weight.Value.ToString("0.##", CultureInfo.InvariantCulture);
                    }
                }

                if (one.MarkerStyle.HasValue)
                {
                    col.Marker = MarkerFromXl(one.MarkerStyle.Value);
                }

                if (one.MarkerSize.HasValue)
                {
                    col.MarkerSize = one.MarkerSize.Value.ToString(CultureInfo.InvariantCulture);
                }

                if (one.MarkerForeRgb.HasValue)
                {
                    col.MarkerColor = OfficeRgbToHex(one.MarkerForeRgb.Value);
                }

                if (one.MarkerBackRgb.HasValue)
                {
                    col.MarkerFill = OfficeRgbToHex(one.MarkerBackRgb.Value);
                }

                if (one.HasDataLabels.HasValue)
                {
                    col.ShowDataLabels = one.HasDataLabels.Value ? "true" : "false";
                }

                if (one.ShowValue.HasValue)
                {
                    col.ShowValue = one.ShowValue.Value ? "true" : "false";
                }

                if (one.ShowPercentage.HasValue)
                {
                    col.ShowPercentage = one.ShowPercentage.Value ? "true" : "false";
                }

                if (one.DataLabelPosition.HasValue)
                {
                    col.LabelPosition = LabelPosFromXl(one.DataLabelPosition.Value);
                }

                if (!string.IsNullOrEmpty(one.DataLabelFontName))
                {
                    col.LabelFont = one.DataLabelFontName;
                }

                if (one.DataLabelFontSize.HasValue)
                {
                    col.LabelSize = one.DataLabelFontSize.Value.ToString("0.##", CultureInfo.InvariantCulture);
                }

                if (one.DataLabelFontColor.HasValue)
                {
                    col.LabelColor = OfficeRgbToHex(one.DataLabelFontColor.Value);
                }

                if (!string.IsNullOrEmpty(one.DataLabelNumberFormat)
                    && one.DataLabelNumberFormat != ";;;")
                {
                    col.LabelFormat = one.DataLabelNumberFormat;
                }

                MarkColumnDataLabelMentions(col);
            }
        }

        private static ChartStyleSnap SnapFromFormat(PptHtmlChartFormat format, PptHtmlChartGrid grid)
        {
            var snap = new ChartStyleSnap { Series = new List<SeriesStyleSnap>() };
            if (format != null)
            {
                if (int.TryParse(format.ChartStyle, NumberStyles.Integer, CultureInfo.InvariantCulture, out int style))
                {
                    snap.ChartStyle = style;
                }

                if (format.Legend != null)
                {
                    if (string.Equals(format.Legend, "none", StringComparison.OrdinalIgnoreCase))
                    {
                        snap.HasLegend = false;
                    }
                    else if (!string.IsNullOrWhiteSpace(format.Legend))
                    {
                        snap.HasLegend = true;
                        snap.LegendPosition = LegendToXl(format.Legend);
                    }
                }

                if (!string.IsNullOrWhiteSpace(format.LegendFontColor))
                {
                    snap.LegendFontColor = format.LegendFontColor;
                }

                if (format.Title != null)
                {
                    snap.HasTitle = !string.IsNullOrEmpty(format.Title);
                    snap.Title = format.Title;
                }

                if (!string.IsNullOrWhiteSpace(format.TitleFontSize))
                {
                    snap.TitleFontSize = format.TitleFontSize;
                }

                if (!string.IsNullOrWhiteSpace(format.TitleFontBold))
                {
                    snap.TitleFontBold = IsTrue(format.TitleFontBold);
                }

                if (!string.IsNullOrWhiteSpace(format.TitleFontColor))
                {
                    snap.TitleFontColor = format.TitleFontColor;
                }

                ApplyAreaColor(format.ChartAreaColor, out bool? areaVis, out int? areaRgb);
                snap.ChartAreaFillVisible = areaVis;
                snap.ChartAreaFillRgb = areaRgb;
                ApplyAreaColor(format.PlotColor, out bool? plotVis, out int? plotRgb);
                snap.PlotFillVisible = plotVis;
                snap.PlotFillRgb = plotRgb;
                if (int.TryParse(format.GapWidth, NumberStyles.Integer, CultureInfo.InvariantCulture, out int gap))
                {
                    snap.GapWidth = gap;
                }

                if (int.TryParse(format.Overlap, NumberStyles.Integer, CultureInfo.InvariantCulture, out int overlap))
                {
                    snap.Overlap = overlap;
                }

                ParseBox(format.PlotBox, out float? l, out float? t, out float? w, out float? h);
                snap.PlotLeft = l;
                snap.PlotTop = t;
                snap.PlotWidth = w;
                snap.PlotHeight = h;
                ParseBox(format.PlotInside, out float? il, out float? it, out float? iw, out float? ih);
                snap.PlotInsideLeft = il;
                snap.PlotInsideTop = it;
                snap.PlotInsideWidth = iw;
                snap.PlotInsideHeight = ih;
                snap.Category = AxisSnapFromExtras(format.AxisXStyle, format.AxisX, format.AxisXFormat);
                snap.Value = AxisSnapFromExtras(format.AxisYStyle, format.AxisY, null);
                snap.ValueSecondary = AxisSnapFromExtras(format.AxisY2Style, format.AxisYSecondary, null);
                if (!string.IsNullOrWhiteSpace(format.Gridlines) && snap.Value != null
                    && snap.Value.HasMajorGridlines == null)
                {
                    snap.Value.HasMajorGridlines = IsTrue(format.Gridlines);
                }
            }

            if (grid != null && grid.Columns != null)
            {
                foreach (PptHtmlChartColumn col in grid.Columns)
                {
                    if (col == null || col.Role == "category")
                    {
                        continue;
                    }

                    snap.Series.Add(SeriesSnapFromColumn(col));
                }
            }

            ApplyChartLevelLabelContent(format, snap);
            return snap;
        }

        /// <summary>图级 data-show-value / percentage 落到列上未写的系列。</summary>
        private static void ApplyChartLevelLabelContent(PptHtmlChartFormat format, ChartStyleSnap snap)
        {
            if (format == null || snap == null || snap.Series == null)
            {
                return;
            }

            bool? chartValue = ParseOptionalBool(format.ShowValue);
            bool? chartPct = ParseOptionalBool(format.ShowPercentage);
            if (chartValue == null && chartPct == null)
            {
                return;
            }

            for (int i = 0; i < snap.Series.Count; i++)
            {
                SeriesStyleSnap one = snap.Series[i];
                if (one == null)
                {
                    continue;
                }

                if (chartValue.HasValue && !one.ShowValue.HasValue)
                {
                    one.ShowValue = chartValue;
                }

                if (chartPct.HasValue && !one.ShowPercentage.HasValue)
                {
                    one.ShowPercentage = chartPct;
                }
            }
        }

        private static bool? ParseOptionalBool(string raw)
        {
            if (string.IsNullOrWhiteSpace(raw))
            {
                return null;
            }

            return IsTrue(raw);
        }

        /// <summary>
        /// 换数套回：结构听 HTML/表；外观白名单从旧图贴皮；HTML 显式皮再盖。
        /// 不再全盘 OverlaySnap（避免旧图次轴 Deleted、关标签等打乱新图）。
        /// </summary>
        private static ChartStyleSnap MergeReplaceSnap(
            ChartStyleSnap oldSnap,
            ChartStyleSnap htmlSnap,
            PptHtmlChartFormat format,
            List<string> warnings)
        {
            ChartStyleSnap snap = htmlSnap ?? new ChartStyleSnap { Series = new List<SeriesStyleSnap>() };
            if (snap.Series == null)
            {
                snap.Series = new List<SeriesStyleSnap>();
            }

            // 展示开关必须在白名单之前；「是否提到」已由解析层 MarkDisplaySwitchMentions 钉死。
            EnsureDisplaySwitches(oldSnap, format, snap, warnings);

            if (oldSnap != null)
            {
                ApplyAppearanceWhitelist(oldSnap, snap, warnings);
            }

            EnsureReplaceAxisStructure(snap, format, warnings);
            NormalizeAxisChrome(snap);
            return snap;
        }

        /// <summary>
        /// 图例/标题/标签/网格/主轴显隐：解析层 Mentioned + 显式关 → 开/关；完全未提 → 跟旧图。
        /// 「是否提到」只信 format / 列解析标志，不从 htmlSnap 反推。
        /// </summary>
        private static void EnsureDisplaySwitches(
            ChartStyleSnap oldSnap,
            PptHtmlChartFormat format,
            ChartStyleSnap snap,
            List<string> warnings)
        {
            if (snap == null)
            {
                return;
            }

            EnsureLegendSwitch(oldSnap, format, snap, warnings);
            EnsureTitleSwitch(oldSnap, format, snap, warnings);
            EnsureDataLabelSwitches(oldSnap, format, snap, warnings);
            EnsureAxisVisibilitySwitches(oldSnap, format, snap, warnings);
            EnsureGridlineSwitch(oldSnap, format, snap, warnings);
        }

        /// <summary>
        /// 主轴显隐：显式 visible=false → 关；解析层 Axis*VisibilityMentioned 且非显式关 → 开；未提 → 跟旧图 Deleted。
        /// 次轴有无仍由 EnsureReplaceAxisStructure 听系列挂 y2 / 稿 y2-visible，不走本套。
        /// 「是否提到」只信 format 解析标志，不读 htmlSnap。
        /// </summary>
        private static void EnsureAxisVisibilitySwitches(
            ChartStyleSnap oldSnap,
            PptHtmlChartFormat format,
            ChartStyleSnap snap,
            List<string> warnings)
        {
            if (snap == null)
            {
                return;
            }

            if (snap.Category == null)
            {
                snap.Category = new AxisStyleSnap();
            }

            if (snap.Value == null)
            {
                snap.Value = new AxisStyleSnap();
            }

            EnsureOnePrimaryAxisVisibility(
                oldSnap != null ? oldSnap.Category : null,
                format,
                forCategory: true,
                dest: snap.Category,
                warnings: warnings);

            EnsureOnePrimaryAxisVisibility(
                oldSnap != null ? oldSnap.Value : null,
                format,
                forCategory: false,
                dest: snap.Value,
                warnings: warnings);
        }

        private static void EnsureOnePrimaryAxisVisibility(
            AxisStyleSnap oldAx,
            PptHtmlChartFormat format,
            bool forCategory,
            AxisStyleSnap dest,
            List<string> warnings)
        {
            if (dest == null)
            {
                return;
            }

            PptHtmlAxisExtras extras = null;
            bool mentioned = false;
            if (format != null)
            {
                if (forCategory)
                {
                    extras = format.AxisXStyle;
                    mentioned = format.AxisXVisibilityMentioned;
                }
                else
                {
                    extras = format.AxisYStyle;
                    mentioned = format.AxisYVisibilityMentioned;
                }
            }

            bool explicitOff = extras != null
                && !string.IsNullOrWhiteSpace(extras.Visible)
                && !IsTrue(extras.Visible);

            string axisName = forCategory ? "横轴" : "纵轴";
            if (explicitOff)
            {
                dest.Deleted = true;
                StyleLog(warnings, axisName + "显隐=关（稿显式 visible=false）");
                return;
            }

            if (mentioned)
            {
                dest.Deleted = false;
                StyleLog(warnings, axisName + "显隐=开（稿写了轴属性）");
                return;
            }

            // 跟旧图：仅当旧图明确 Deleted=true 才继续藏；否则显示
            dest.Deleted = oldAx != null && oldAx.Deleted == true;
            StyleLog(warnings, axisName + "显隐="
                + (dest.Deleted == true ? "关" : "开")
                + "（跟旧图）");
        }

        /// <summary>
        /// 解析层：按「支持属性 ↔ 展示开关」登记表钉死各 Mentioned 标志。
        /// 合并 Ensure* 只读这些标志，禁止从 htmlSnap 反推。
        /// </summary>
        public static void MarkDisplaySwitchMentions(PptHtmlChartFormat format)
        {
            if (format == null)
            {
                return;
            }

            // 图例：data-legend、data-legend-font-color
            format.LegendMentioned = format.Legend != null
                || !string.IsNullOrWhiteSpace(format.LegendFontColor);

            // 标题：data-title（含空串=显式关）、字色/字号/粗体
            format.TitleMentioned = format.Title != null
                || !string.IsNullOrWhiteSpace(format.TitleFontColor)
                || !string.IsNullOrWhiteSpace(format.TitleFontSize)
                || !string.IsNullOrWhiteSpace(format.TitleFontBold);

            // 图级数据标签：data-show-data-labels / show-value / show-percentage
            format.DataLabelsMentioned = !string.IsNullOrWhiteSpace(format.ShowDataLabels)
                || !string.IsNullOrWhiteSpace(format.ShowValue)
                || !string.IsNullOrWhiteSpace(format.ShowPercentage);

            // 网格：data-gridlines、轴 extras 的 grid/grid-color
            format.GridlinesMentioned = !string.IsNullOrWhiteSpace(format.Gridlines)
                || AxisExtrasMentionsGrid(format.AxisYStyle)
                || AxisExtrasMentionsGrid(format.AxisXStyle)
                || AxisExtrasMentionsGrid(format.AxisY2Style);

            format.AxisXVisibilityMentioned = MentionsPrimaryAxisXVisibility(format);
            format.AxisYVisibilityMentioned = MentionsPrimaryAxisYVisibility(format);
        }

        /// <summary>兼容旧调用名；等价 MarkDisplaySwitchMentions。</summary>
        public static void MarkAxisVisibilityMentions(PptHtmlChartFormat format)
        {
            MarkDisplaySwitchMentions(format);
        }

        /// <summary>解析层：列上与数据标签开关相关的支持属性。</summary>
        public static void MarkColumnDataLabelMentions(PptHtmlChartColumn col)
        {
            if (col == null)
            {
                return;
            }

            if (string.Equals(col.Role, "category", StringComparison.OrdinalIgnoreCase))
            {
                col.DataLabelsMentioned = false;
                return;
            }

            col.DataLabelsMentioned = !string.IsNullOrWhiteSpace(col.ShowDataLabels)
                || !string.IsNullOrWhiteSpace(col.ShowValue)
                || !string.IsNullOrWhiteSpace(col.ShowPercentage)
                || !string.IsNullOrWhiteSpace(col.LabelPosition)
                || !string.IsNullOrWhiteSpace(col.LabelFont)
                || !string.IsNullOrWhiteSpace(col.LabelSize)
                || !string.IsNullOrWhiteSpace(col.LabelColor)
                || !string.IsNullOrWhiteSpace(col.LabelFormat);
        }

        private static bool MentionsPrimaryAxisXVisibility(PptHtmlChartFormat format)
        {
            return AxisExtrasMentionsVisibilitySwitch(format.AxisXStyle)
                || !string.IsNullOrWhiteSpace(format.AxisX)
                || !string.IsNullOrWhiteSpace(format.AxisXType)
                || !string.IsNullOrWhiteSpace(format.AxisXFormat)
                || !string.IsNullOrWhiteSpace(format.AxisXTickCount)
                || !string.IsNullOrWhiteSpace(format.AxisXTickSpacing)
                || !string.IsNullOrWhiteSpace(format.AxisXBetween);
        }

        private static bool MentionsPrimaryAxisYVisibility(PptHtmlChartFormat format)
        {
            return AxisExtrasMentionsVisibilitySwitch(format.AxisYStyle)
                || !string.IsNullOrWhiteSpace(format.AxisY)
                || !string.IsNullOrWhiteSpace(format.AxisYMin)
                || !string.IsNullOrWhiteSpace(format.AxisYMax)
                || !string.IsNullOrWhiteSpace(format.AxisYMajorUnit);
        }

        /// <summary>
        /// data-axis-{x|y}-* extras 中与主轴显隐相关的支持项：
        /// visible / tick-* / major|minor-tick / format / grid|grid-color / line|line-weight。
        /// </summary>
        private static bool AxisExtrasMentionsVisibilitySwitch(PptHtmlAxisExtras extras)
        {
            if (extras == null)
            {
                return false;
            }

            return !string.IsNullOrWhiteSpace(extras.Visible)
                || !string.IsNullOrWhiteSpace(extras.TickFont)
                || !string.IsNullOrWhiteSpace(extras.TickColor)
                || !string.IsNullOrWhiteSpace(extras.TickSize)
                || !string.IsNullOrWhiteSpace(extras.TickPosition)
                || !string.IsNullOrWhiteSpace(extras.MajorTick)
                || !string.IsNullOrWhiteSpace(extras.MinorTick)
                || !string.IsNullOrWhiteSpace(extras.Format)
                || !string.IsNullOrWhiteSpace(extras.Grid)
                || !string.IsNullOrWhiteSpace(extras.GridColor)
                || !string.IsNullOrWhiteSpace(extras.Line)
                || !string.IsNullOrWhiteSpace(extras.LineWeight);
        }

        private static void EnsureLegendSwitch(
            ChartStyleSnap oldSnap,
            PptHtmlChartFormat format,
            ChartStyleSnap snap,
            List<string> warnings)
        {
            bool explicitOff = format != null
                && format.Legend != null
                && string.Equals(format.Legend.Trim(), "none", StringComparison.OrdinalIgnoreCase);

            bool mentioned = format != null && format.LegendMentioned;

            if (explicitOff)
            {
                snap.HasLegend = false;
                StyleLog(warnings, "图例开关=关（稿显式 none）");
                return;
            }

            if (mentioned)
            {
                snap.HasLegend = true;
                StyleLog(warnings, "图例开关=开（稿写了图例属性）");
                return;
            }

            snap.HasLegend = oldSnap != null && oldSnap.HasLegend == true;
            StyleLog(warnings, "图例开关=" + (snap.HasLegend == true ? "开" : "关") + "（跟旧图）");
        }

        private static void EnsureTitleSwitch(
            ChartStyleSnap oldSnap,
            PptHtmlChartFormat format,
            ChartStyleSnap snap,
            List<string> warnings)
        {
            bool explicitOff = format != null
                && format.Title != null
                && string.IsNullOrEmpty(format.Title);

            bool mentioned = format != null && format.TitleMentioned;

            if (explicitOff)
            {
                snap.HasTitle = false;
                snap.Title = "";
                StyleLog(warnings, "标题开关=关（稿空标题） TitleMentioned="
                    + (format != null && format.TitleMentioned)
                    + " format.Title=[" + (format.Title ?? "(null)") + "]");
                return;
            }

            if (mentioned)
            {
                snap.HasTitle = true;
                StyleLog(warnings, "标题开关=开（稿写了标题属性） format.Title=["
                    + (format.Title ?? "(null)") + "] snap.Title=[" + (snap.Title ?? "(null)") + "]");
                return;
            }

            // 完全未提：有旧图则跟旧图；新建无旧图 → 默认关（压住 AddChart2 自带标题）
            snap.HasTitle = oldSnap != null && oldSnap.HasTitle == true;
            StyleLog(warnings, "标题开关=" + (snap.HasTitle == true ? "开" : "关")
                + (oldSnap == null ? "（新建默认关）" : "（跟旧图）"));
        }

        private static void EnsureDataLabelSwitches(
            ChartStyleSnap oldSnap,
            PptHtmlChartFormat format,
            ChartStyleSnap snap,
            List<string> warnings)
        {
            if (snap.Series == null)
            {
                return;
            }

            bool chartLevelOff = format != null
                && !string.IsNullOrWhiteSpace(format.ShowDataLabels)
                && !IsTrue(format.ShowDataLabels);
            bool chartMentioned = format != null && format.DataLabelsMentioned;

            for (int i = 0; i < snap.Series.Count; i++)
            {
                SeriesStyleSnap one = snap.Series[i];
                if (one == null)
                {
                    continue;
                }

                bool explicitOff = chartLevelOff || one.DataLabelsExplicitOff;
                bool mentioned = chartMentioned || one.DataLabelsMentioned;

                if (explicitOff)
                {
                    one.HasDataLabels = false;
                    continue;
                }

                if (mentioned)
                {
                    one.HasDataLabels = true;
                    continue;
                }

                SeriesStyleSnap oldOne = oldSnap != null
                    && oldSnap.Series != null
                    && i < oldSnap.Series.Count
                    ? oldSnap.Series[i]
                    : null;
                one.HasDataLabels = oldOne != null && oldOne.HasDataLabels == true;
            }

            StyleLog(warnings, "数据标签开关已按稿/旧图推断 series=" + snap.Series.Count);
        }

        private static void EnsureGridlineSwitch(
            ChartStyleSnap oldSnap,
            PptHtmlChartFormat format,
            ChartStyleSnap snap,
            List<string> warnings)
        {
            if (snap.Value == null)
            {
                snap.Value = new AxisStyleSnap();
            }

            bool explicitOff = (format != null
                    && !string.IsNullOrWhiteSpace(format.Gridlines)
                    && !IsTrue(format.Gridlines))
                || AxisExtrasGridOff(format != null ? format.AxisYStyle : null);

            bool mentioned = format != null && format.GridlinesMentioned;

            if (explicitOff)
            {
                snap.Value.HasMajorGridlines = false;
                StyleLog(warnings, "网格开关=关（稿显式）");
                return;
            }

            if (mentioned)
            {
                snap.Value.HasMajorGridlines = true;
                StyleLog(warnings, "网格开关=开（稿写了网格属性）");
                return;
            }

            bool oldOn = oldSnap != null
                && oldSnap.Value != null
                && oldSnap.Value.HasMajorGridlines == true
                && oldSnap.Value.GridlineVisible != false;
            snap.Value.HasMajorGridlines = oldOn;
            StyleLog(warnings, "网格开关=" + (oldOn ? "开" : "关") + "（跟旧图）");
        }

        private static bool AxisExtrasGridOff(PptHtmlAxisExtras extras)
        {
            return extras != null
                && !string.IsNullOrWhiteSpace(extras.Grid)
                && !IsTrue(extras.Grid);
        }

        private static bool AxisExtrasMentionsGrid(PptHtmlAxisExtras extras)
        {
            return extras != null
                && (!string.IsNullOrWhiteSpace(extras.Grid)
                    || !string.IsNullOrWhiteSpace(extras.GridColor));
        }

        private static void ApplyAppearanceWhitelist(
            ChartStyleSnap oldSnap,
            ChartStyleSnap snap,
            List<string> warnings)
        {
            if (oldSnap == null || snap == null)
            {
                return;
            }

            if (!snap.ChartStyle.HasValue && oldSnap.ChartStyle.HasValue)
            {
                snap.ChartStyle = oldSnap.ChartStyle;
            }

            if (!snap.ChartColor.HasValue && oldSnap.ChartColor.HasValue)
            {
                snap.ChartColor = oldSnap.ChartColor;
            }

            // 标题/图例皮可补旧图；开关由 EnsureDisplaySwitches 推断（写了皮⇒开，沉默跟旧图）
            if (snap.Title == null && oldSnap.Title != null)
            {
                snap.Title = oldSnap.Title;
            }

            if (string.IsNullOrEmpty(snap.TitleFontColor) && !string.IsNullOrEmpty(oldSnap.TitleFontColor))
            {
                snap.TitleFontColor = oldSnap.TitleFontColor;
            }

            if (string.IsNullOrEmpty(snap.TitleFontSize) && !string.IsNullOrEmpty(oldSnap.TitleFontSize))
            {
                snap.TitleFontSize = oldSnap.TitleFontSize;
            }

            if (!snap.TitleFontBold.HasValue && oldSnap.TitleFontBold.HasValue)
            {
                snap.TitleFontBold = oldSnap.TitleFontBold;
            }

            if (!snap.LegendPosition.HasValue && oldSnap.LegendPosition.HasValue)
            {
                snap.LegendPosition = oldSnap.LegendPosition;
            }

            if (string.IsNullOrEmpty(snap.LegendFontColor) && !string.IsNullOrEmpty(oldSnap.LegendFontColor))
            {
                snap.LegendFontColor = oldSnap.LegendFontColor;
            }

            if (!snap.ChartAreaFillVisible.HasValue && oldSnap.ChartAreaFillVisible.HasValue)
            {
                snap.ChartAreaFillVisible = oldSnap.ChartAreaFillVisible;
                snap.ChartAreaFillRgb = oldSnap.ChartAreaFillRgb;
            }

            if (!snap.PlotFillVisible.HasValue && oldSnap.PlotFillVisible.HasValue)
            {
                snap.PlotFillVisible = oldSnap.PlotFillVisible;
                snap.PlotFillRgb = oldSnap.PlotFillRgb;
            }

            if (!snap.GapWidth.HasValue && oldSnap.GapWidth.HasValue)
            {
                snap.GapWidth = oldSnap.GapWidth;
            }

            if (!snap.Overlap.HasValue && oldSnap.Overlap.HasValue)
            {
                snap.Overlap = oldSnap.Overlap;
            }

            CopyPlotBoxIfEmpty(oldSnap, snap);

            snap.Category = MergeAxisAppearance(oldSnap.Category, snap.Category);
            snap.Value = MergeAxisAppearance(oldSnap.Value, snap.Value);
            snap.ValueSecondary = MergeAxisAppearance(oldSnap.ValueSecondary, snap.ValueSecondary);

            if (oldSnap.Series != null)
            {
                for (int i = 0; i < snap.Series.Count && i < oldSnap.Series.Count; i++)
                {
                    MergeSeriesAppearance(oldSnap.Series[i], snap.Series[i]);
                }
            }

            StyleLog(warnings, "换数白名单贴皮 "
                + "style=" + snap.ChartStyle
                + " chartColor=" + snap.ChartColor
                + " gap=" + snap.GapWidth
                + " plotBox=" + (snap.PlotLeft.HasValue ? "yes" : "no")
                + " series=" + snap.Series.Count);
        }

        private static void CopyPlotBoxIfEmpty(ChartStyleSnap oldSnap, ChartStyleSnap snap)
        {
            if (!snap.PlotLeft.HasValue && oldSnap.PlotLeft.HasValue)
            {
                snap.PlotLeft = oldSnap.PlotLeft;
            }

            if (!snap.PlotTop.HasValue && oldSnap.PlotTop.HasValue)
            {
                snap.PlotTop = oldSnap.PlotTop;
            }

            if (!snap.PlotWidth.HasValue && oldSnap.PlotWidth.HasValue)
            {
                snap.PlotWidth = oldSnap.PlotWidth;
            }

            if (!snap.PlotHeight.HasValue && oldSnap.PlotHeight.HasValue)
            {
                snap.PlotHeight = oldSnap.PlotHeight;
            }

            if (!snap.PlotInsideLeft.HasValue && oldSnap.PlotInsideLeft.HasValue)
            {
                snap.PlotInsideLeft = oldSnap.PlotInsideLeft;
            }

            if (!snap.PlotInsideTop.HasValue && oldSnap.PlotInsideTop.HasValue)
            {
                snap.PlotInsideTop = oldSnap.PlotInsideTop;
            }

            if (!snap.PlotInsideWidth.HasValue && oldSnap.PlotInsideWidth.HasValue)
            {
                snap.PlotInsideWidth = oldSnap.PlotInsideWidth;
            }

            if (!snap.PlotInsideHeight.HasValue && oldSnap.PlotInsideHeight.HasValue)
            {
                snap.PlotInsideHeight = oldSnap.PlotInsideHeight;
            }
        }

        /// <summary>
        /// 轴皮从旧图补；主轴 Deleted 由 EnsureAxisVisibilitySwitches 推断，此处不改结构。
        /// </summary>
        private static AxisStyleSnap MergeAxisAppearance(AxisStyleSnap oldAx, AxisStyleSnap htmlAx)
        {
            AxisStyleSnap ax = htmlAx ?? new AxisStyleSnap();
            if (oldAx == null)
            {
                return ax;
            }

            if (string.IsNullOrEmpty(ax.TickFontName) && !string.IsNullOrEmpty(oldAx.TickFontName))
            {
                ax.TickFontName = oldAx.TickFontName;
            }

            if (!ax.TickFontColor.HasValue && oldAx.TickFontColor.HasValue)
            {
                ax.TickFontColor = oldAx.TickFontColor;
            }

            if (!ax.TickFontSize.HasValue && oldAx.TickFontSize.HasValue)
            {
                ax.TickFontSize = oldAx.TickFontSize;
            }

            if (!ax.TickLabelPosition.HasValue && oldAx.TickLabelPosition.HasValue)
            {
                ax.TickLabelPosition = oldAx.TickLabelPosition;
            }

            if (!ax.MajorTickMark.HasValue && oldAx.MajorTickMark.HasValue)
            {
                ax.MajorTickMark = oldAx.MajorTickMark;
            }

            if (!ax.MinorTickMark.HasValue && oldAx.MinorTickMark.HasValue)
            {
                ax.MinorTickMark = oldAx.MinorTickMark;
            }

            if (!ax.LineVisible.HasValue && oldAx.LineVisible.HasValue)
            {
                ax.LineVisible = oldAx.LineVisible;
            }

            if (!ax.LineRgb.HasValue && oldAx.LineRgb.HasValue)
            {
                ax.LineRgb = oldAx.LineRgb;
            }

            if (!ax.LineWeight.HasValue && oldAx.LineWeight.HasValue)
            {
                ax.LineWeight = oldAx.LineWeight;
            }

            if (!ax.MajorGridlineRgb.HasValue && oldAx.MajorGridlineRgb.HasValue)
            {
                ax.MajorGridlineRgb = oldAx.MajorGridlineRgb;
            }

            if (!ax.GridlineWeight.HasValue && oldAx.GridlineWeight.HasValue)
            {
                ax.GridlineWeight = oldAx.GridlineWeight;
            }

            return ax;
        }

        /// <summary>
        /// 系列：类型/挂轴/标签开关只听 HTML；填色/线/标记从旧图补。
        /// </summary>
        private static void MergeSeriesAppearance(SeriesStyleSnap oldS, SeriesStyleSnap htmlS)
        {
            if (oldS == null || htmlS == null)
            {
                return;
            }

            if (htmlS.Fill == null && oldS.Fill != null)
            {
                htmlS.Fill = CloneFillSnap(oldS.Fill);
            }
            else if (htmlS.Fill != null && oldS.Fill != null)
            {
                htmlS.Fill = OverlayFill(CloneFillSnap(oldS.Fill), htmlS.Fill);
            }

            if (htmlS.Line == null && oldS.Line != null)
            {
                htmlS.Line = CloneLineSnap(oldS.Line);
            }
            else if (htmlS.Line != null && oldS.Line != null)
            {
                htmlS.Line = OverlayLine(CloneLineSnap(oldS.Line), htmlS.Line);
            }

            if ((htmlS.PointFills == null || htmlS.PointFills.Count == 0)
                && oldS.PointFills != null
                && oldS.PointFills.Count > 0)
            {
                htmlS.PointFills = oldS.PointFills;
            }

            if (!htmlS.MarkerStyle.HasValue && oldS.MarkerStyle.HasValue)
            {
                htmlS.MarkerStyle = oldS.MarkerStyle;
            }

            if (!htmlS.MarkerSize.HasValue && oldS.MarkerSize.HasValue)
            {
                htmlS.MarkerSize = oldS.MarkerSize;
            }

            if (!htmlS.MarkerForeRgb.HasValue && oldS.MarkerForeRgb.HasValue)
            {
                htmlS.MarkerForeRgb = oldS.MarkerForeRgb;
            }

            if (!htmlS.MarkerBackRgb.HasValue && oldS.MarkerBackRgb.HasValue)
            {
                htmlS.MarkerBackRgb = oldS.MarkerBackRgb;
            }

            if (!htmlS.ShowValue.HasValue && oldS.ShowValue.HasValue)
            {
                htmlS.ShowValue = oldS.ShowValue;
            }

            if (!htmlS.ShowPercentage.HasValue && oldS.ShowPercentage.HasValue)
            {
                htmlS.ShowPercentage = oldS.ShowPercentage;
            }
        }

        private static FillSnap CloneFillSnap(FillSnap src)
        {
            if (src == null)
            {
                return null;
            }

            return new FillSnap
            {
                Visible = src.Visible,
                FillType = src.FillType,
                SolidRgb = src.SolidRgb,
                Angle = src.Angle,
                Stops = src.Stops
            };
        }

        private static LineSnap CloneLineSnap(LineSnap src)
        {
            if (src == null)
            {
                return null;
            }

            return new LineSnap
            {
                Visible = src.Visible,
                Rgb = src.Rgb,
                Weight = src.Weight,
                Stops = src.Stops
            };
        }

        /// <summary>
        /// 有无次轴：听系列挂 y2 或稿 data-axis-y2-visible；不听旧图 Deleted。
        /// 主轴显隐已由 EnsureAxisVisibilitySwitches 推断，此处不再把 Category/Value 的 null 打成 false。
        /// </summary>
        private static void EnsureReplaceAxisStructure(
            ChartStyleSnap snap,
            PptHtmlChartFormat format,
            List<string> warnings)
        {
            if (snap == null)
            {
                return;
            }

            bool needY2 = false;
            if (snap.Series != null)
            {
                for (int i = 0; i < snap.Series.Count; i++)
                {
                    if (snap.Series[i] != null && snap.Series[i].AxisGroup == XlSecondary)
                    {
                        needY2 = true;
                        break;
                    }
                }
            }

            bool? htmlY2Visible = null;
            if (format != null
                && format.AxisY2Style != null
                && !string.IsNullOrWhiteSpace(format.AxisY2Style.Visible))
            {
                htmlY2Visible = IsTrue(format.AxisY2Style.Visible);
            }

            if (snap.Category == null)
            {
                snap.Category = new AxisStyleSnap();
            }

            if (snap.Value == null)
            {
                snap.Value = new AxisStyleSnap();
            }

            if (htmlY2Visible == true || needY2)
            {
                if (snap.ValueSecondary == null)
                {
                    snap.ValueSecondary = new AxisStyleSnap();
                }

                snap.ValueSecondary.Deleted = false;
                StyleLog(warnings, "次轴结构打开 needY2=" + needY2
                    + " htmlVisible=" + (htmlY2Visible.HasValue ? htmlY2Visible.Value.ToString() : "null"));
            }
            else if (htmlY2Visible == false)
            {
                if (snap.ValueSecondary == null)
                {
                    snap.ValueSecondary = new AxisStyleSnap { Deleted = true };
                }
                else
                {
                    snap.ValueSecondary.Deleted = true;
                }
            }
            else
            {
                if (snap.ValueSecondary == null)
                {
                    snap.ValueSecondary = new AxisStyleSnap { Deleted = true };
                }
                else if (snap.ValueSecondary.Deleted != false)
                {
                    snap.ValueSecondary.Deleted = true;
                }
            }
        }

        private static ChartStyleSnap OverlaySnap(ChartStyleSnap oldSnap, ChartStyleSnap htmlSnap)
        {
            if (oldSnap == null)
            {
                return htmlSnap;
            }

            if (htmlSnap == null)
            {
                return oldSnap;
            }

            if (htmlSnap.ChartStyle.HasValue)
            {
                oldSnap.ChartStyle = htmlSnap.ChartStyle;
            }

            if (htmlSnap.ChartColor.HasValue)
            {
                oldSnap.ChartColor = htmlSnap.ChartColor;
            }

            if (htmlSnap.HasTitle.HasValue)
            {
                oldSnap.HasTitle = htmlSnap.HasTitle;
            }

            if (htmlSnap.Title != null)
            {
                oldSnap.Title = htmlSnap.Title;
            }

            if (!string.IsNullOrEmpty(htmlSnap.TitleFontColor))
            {
                oldSnap.TitleFontColor = htmlSnap.TitleFontColor;
            }

            if (!string.IsNullOrEmpty(htmlSnap.TitleFontSize))
            {
                oldSnap.TitleFontSize = htmlSnap.TitleFontSize;
            }

            if (htmlSnap.TitleFontBold.HasValue)
            {
                oldSnap.TitleFontBold = htmlSnap.TitleFontBold;
            }

            if (htmlSnap.HasLegend.HasValue)
            {
                oldSnap.HasLegend = htmlSnap.HasLegend;
            }

            if (htmlSnap.LegendPosition.HasValue)
            {
                oldSnap.LegendPosition = htmlSnap.LegendPosition;
            }

            if (!string.IsNullOrEmpty(htmlSnap.LegendFontColor))
            {
                oldSnap.LegendFontColor = htmlSnap.LegendFontColor;
            }

            if (htmlSnap.ChartAreaFillVisible.HasValue)
            {
                oldSnap.ChartAreaFillVisible = htmlSnap.ChartAreaFillVisible;
                oldSnap.ChartAreaFillRgb = htmlSnap.ChartAreaFillRgb;
            }

            if (htmlSnap.PlotFillVisible.HasValue)
            {
                oldSnap.PlotFillVisible = htmlSnap.PlotFillVisible;
                oldSnap.PlotFillRgb = htmlSnap.PlotFillRgb;
            }

            if (htmlSnap.GapWidth.HasValue)
            {
                oldSnap.GapWidth = htmlSnap.GapWidth;
            }

            if (htmlSnap.Overlap.HasValue)
            {
                oldSnap.Overlap = htmlSnap.Overlap;
            }

            OverlayBox(htmlSnap, oldSnap);
            oldSnap.Category = OverlayAxis(oldSnap.Category, htmlSnap.Category);
            oldSnap.Value = OverlayAxis(oldSnap.Value, htmlSnap.Value);
            oldSnap.ValueSecondary = OverlayAxis(oldSnap.ValueSecondary, htmlSnap.ValueSecondary);
            if (htmlSnap.Series != null && htmlSnap.Series.Count > 0)
            {
                if (oldSnap.Series == null)
                {
                    oldSnap.Series = new List<SeriesStyleSnap>();
                }

                for (int i = 0; i < htmlSnap.Series.Count; i++)
                {
                    if (i < oldSnap.Series.Count)
                    {
                        oldSnap.Series[i] = OverlaySeries(oldSnap.Series[i], htmlSnap.Series[i]);
                    }
                    else
                    {
                        oldSnap.Series.Add(htmlSnap.Series[i]);
                    }
                }
            }

            return oldSnap;
        }

        private static void OverlayBox(ChartStyleSnap src, ChartStyleSnap dest)
        {
            if (src.PlotLeft.HasValue)
            {
                dest.PlotLeft = src.PlotLeft;
            }

            if (src.PlotTop.HasValue)
            {
                dest.PlotTop = src.PlotTop;
            }

            if (src.PlotWidth.HasValue)
            {
                dest.PlotWidth = src.PlotWidth;
            }

            if (src.PlotHeight.HasValue)
            {
                dest.PlotHeight = src.PlotHeight;
            }

            if (src.PlotInsideLeft.HasValue)
            {
                dest.PlotInsideLeft = src.PlotInsideLeft;
            }

            if (src.PlotInsideTop.HasValue)
            {
                dest.PlotInsideTop = src.PlotInsideTop;
            }

            if (src.PlotInsideWidth.HasValue)
            {
                dest.PlotInsideWidth = src.PlotInsideWidth;
            }

            if (src.PlotInsideHeight.HasValue)
            {
                dest.PlotInsideHeight = src.PlotInsideHeight;
            }
        }

        private static AxisStyleSnap OverlayAxis(AxisStyleSnap dest, AxisStyleSnap src)
        {
            if (src == null)
            {
                return dest;
            }

            if (dest == null)
            {
                return null;
            }

            // 有无轴、有无标题、标题文案以旧图为准，不跟 HTML 增删
            if (dest.Deleted == true)
            {
                return dest;
            }

            if (src.Deleted == true)
            {
                return dest;
            }

            if (!string.IsNullOrEmpty(src.TickFontName))
            {
                dest.TickFontName = src.TickFontName;
            }

            if (src.TickFontColor.HasValue)
            {
                dest.TickFontColor = src.TickFontColor;
            }

            if (src.TickFontSize.HasValue)
            {
                dest.TickFontSize = src.TickFontSize;
            }

            if (src.TickLabelPosition.HasValue)
            {
                dest.TickLabelPosition = src.TickLabelPosition;
            }

            if (src.MajorTickMark.HasValue)
            {
                dest.MajorTickMark = src.MajorTickMark;
            }

            if (src.MinorTickMark.HasValue)
            {
                dest.MinorTickMark = src.MinorTickMark;
            }

            if (!string.IsNullOrEmpty(src.NumberFormat))
            {
                dest.NumberFormat = src.NumberFormat;
            }

            if (src.HasMajorGridlines.HasValue)
            {
                dest.HasMajorGridlines = src.HasMajorGridlines;
            }

            if (src.MajorGridlineRgb.HasValue)
            {
                dest.MajorGridlineRgb = src.MajorGridlineRgb;
            }

            if (src.GridlineVisible.HasValue)
            {
                dest.GridlineVisible = src.GridlineVisible;
            }

            if (src.GridlineWeight.HasValue)
            {
                dest.GridlineWeight = src.GridlineWeight;
            }

            if (src.GridlineTransparency.HasValue)
            {
                dest.GridlineTransparency = src.GridlineTransparency;
            }

            if (src.LineVisible.HasValue)
            {
                dest.LineVisible = src.LineVisible;
            }

            if (src.LineRgb.HasValue)
            {
                dest.LineRgb = src.LineRgb;
            }

            if (src.LineWeight.HasValue)
            {
                dest.LineWeight = src.LineWeight;
            }

            return dest;
        }

        private static SeriesStyleSnap OverlaySeries(SeriesStyleSnap dest, SeriesStyleSnap src)
        {
            if (src == null)
            {
                return dest;
            }

            if (dest == null)
            {
                return src;
            }

            if (src.ChartType.HasValue)
            {
                // th data-series-type="pie" 只表示饼族，不能把旧图 pie3d 压成 pie2d。
                // 2D/3D 只听节点 data-chart-type。
                if (!(dest.ChartType.HasValue
                    && IsPieXl(dest.ChartType.Value)
                    && IsPieXl(src.ChartType.Value)))
                {
                    dest.ChartType = src.ChartType;
                }
            }

            if (src.AxisGroup.HasValue)
            {
                dest.AxisGroup = src.AxisGroup;
            }

            dest.Fill = OverlayFill(dest.Fill, src.Fill);
            dest.Line = OverlayLine(dest.Line, src.Line);
            if (src.PointFills != null && src.PointFills.Count > 0)
            {
                dest.PointFills = src.PointFills;
            }
            if (src.MarkerStyle.HasValue)
            {
                dest.MarkerStyle = src.MarkerStyle;
            }

            if (src.MarkerSize.HasValue)
            {
                dest.MarkerSize = src.MarkerSize;
            }

            if (src.MarkerForeRgb.HasValue)
            {
                dest.MarkerForeRgb = src.MarkerForeRgb;
            }

            if (src.MarkerBackRgb.HasValue)
            {
                dest.MarkerBackRgb = src.MarkerBackRgb;
            }

            if (src.HasDataLabels.HasValue)
            {
                dest.HasDataLabels = src.HasDataLabels;
            }

            if (src.ShowValue.HasValue)
            {
                dest.ShowValue = src.ShowValue;
            }

            if (src.ShowPercentage.HasValue)
            {
                dest.ShowPercentage = src.ShowPercentage;
            }

            if (src.DataLabelPosition.HasValue)
            {
                dest.DataLabelPosition = src.DataLabelPosition;
            }

            if (!string.IsNullOrEmpty(src.DataLabelFontName))
            {
                dest.DataLabelFontName = src.DataLabelFontName;
            }

            if (src.DataLabelFontSize.HasValue)
            {
                dest.DataLabelFontSize = src.DataLabelFontSize;
            }

            if (src.DataLabelFontColor.HasValue)
            {
                dest.DataLabelFontColor = src.DataLabelFontColor;
            }

            if (!string.IsNullOrEmpty(src.DataLabelNumberFormat))
            {
                dest.DataLabelNumberFormat = src.DataLabelNumberFormat;
            }

            return dest;
        }

        private static FillSnap OverlayFill(FillSnap dest, FillSnap src)
        {
            if (src == null)
            {
                return dest;
            }

            if (dest == null)
            {
                return src;
            }

            if (src.Visible.HasValue)
            {
                dest.Visible = src.Visible;
            }

            if (src.FillType.HasValue)
            {
                dest.FillType = src.FillType;
            }

            if (src.SolidRgb.HasValue)
            {
                dest.SolidRgb = src.SolidRgb;
            }

            if (src.Angle.HasValue)
            {
                dest.Angle = src.Angle;
            }

            if (src.Stops != null && src.Stops.Count > 0)
            {
                dest.Stops = src.Stops;
            }

            return dest;
        }

        private static LineSnap OverlayLine(LineSnap dest, LineSnap src)
        {
            if (src == null)
            {
                return dest;
            }

            if (dest == null)
            {
                return src;
            }

            if (src.Visible.HasValue)
            {
                dest.Visible = src.Visible;
            }

            if (src.Rgb.HasValue)
            {
                dest.Rgb = src.Rgb;
            }

            if (src.Weight.HasValue)
            {
                dest.Weight = src.Weight;
            }

            if (src.Stops != null && src.Stops.Count > 0)
            {
                dest.Stops = src.Stops;
            }

            return dest;
        }

        private static SeriesStyleSnap SeriesSnapFromColumn(PptHtmlChartColumn col)
        {
            var one = new SeriesStyleSnap();
            if (TryParseSeriesXl(col.SeriesType, col.Marker, out int xl))
            {
                one.ChartType = xl;
            }

            if (IsSeriesAxisY2(col.AxisY))
            {
                one.AxisGroup = XlSecondary;
            }
            else if (IsSeriesAxisY(col.AxisY))
            {
                one.AxisGroup = XlPrimary;
            }

            one.Fill = FillFromColumn(col);
            one.Line = LineFromColumn(col);
            if (!string.IsNullOrWhiteSpace(col.Marker))
            {
                one.MarkerStyle = MarkerToXl(col.Marker);
            }

            if (int.TryParse(col.MarkerSize, NumberStyles.Integer, CultureInfo.InvariantCulture, out int ms))
            {
                one.MarkerSize = ms;
            }

            if (TryParseHexToOffice(col.MarkerColor, out int mFore))
            {
                one.MarkerForeRgb = mFore;
            }

            if (TryParseHexToOffice(col.MarkerFill, out int mBack))
            {
                one.MarkerBackRgb = mBack;
            }

            if (!string.IsNullOrWhiteSpace(col.ShowDataLabels))
            {
                one.HasDataLabels = IsTrue(col.ShowDataLabels);
            }

            if (!string.IsNullOrWhiteSpace(col.ShowValue))
            {
                one.ShowValue = IsTrue(col.ShowValue);
            }

            if (!string.IsNullOrWhiteSpace(col.ShowPercentage))
            {
                one.ShowPercentage = IsTrue(col.ShowPercentage);
            }

            if (!string.IsNullOrWhiteSpace(col.LabelPosition))
            {
                one.DataLabelPosition = LabelPosToXl(col.LabelPosition);
            }

            if (!string.IsNullOrWhiteSpace(col.LabelFont))
            {
                one.DataLabelFontName = col.LabelFont;
            }

            if (double.TryParse(col.LabelSize, NumberStyles.Float, CultureInfo.InvariantCulture, out double lsz))
            {
                one.DataLabelFontSize = lsz;
            }

            if (TryParseHexToOffice(col.LabelColor, out int lRgb))
            {
                one.DataLabelFontColor = lRgb;
            }

            if (!string.IsNullOrWhiteSpace(col.LabelFormat))
            {
                one.DataLabelNumberFormat = col.LabelFormat;
            }

            MarkColumnDataLabelMentions(col);
            one.DataLabelsMentioned = col.DataLabelsMentioned;
            one.DataLabelsExplicitOff = !string.IsNullOrWhiteSpace(col.ShowDataLabels)
                && !IsTrue(col.ShowDataLabels);

            return one;
        }

        private static FillSnap FillFromColumn(PptHtmlChartColumn col)
        {
            List<GradientStopSnap> stops = DecodeGradient(col.FillGradient);
            if (stops != null && stops.Count > 0)
            {
                var fill = new FillSnap
                {
                    Visible = true,
                    FillType = MsoFillGradient,
                    Stops = stops
                };
                if (double.TryParse(col.FillAngle, NumberStyles.Float, CultureInfo.InvariantCulture, out double ang))
                {
                    fill.Angle = ang;
                }

                return fill;
            }

            if (string.Equals(col.Color, "none", StringComparison.OrdinalIgnoreCase))
            {
                return new FillSnap { Visible = false };
            }

            if (TryParseHexToOffice(col.Color, out int rgb))
            {
                return new FillSnap { Visible = true, SolidRgb = rgb };
            }

            return null;
        }

        private static LineSnap LineFromColumn(PptHtmlChartColumn col)
        {
            if (string.IsNullOrWhiteSpace(col.Line) && string.IsNullOrWhiteSpace(col.LineWeight))
            {
                return null;
            }

            var line = new LineSnap();
            if (string.Equals(col.Line, "none", StringComparison.OrdinalIgnoreCase))
            {
                line.Visible = false;
            }
            else if (TryParseHexToOffice(col.Line, out int rgb))
            {
                line.Visible = true;
                line.Rgb = rgb;
            }

            if (double.TryParse(col.LineWeight, NumberStyles.Float, CultureInfo.InvariantCulture, out double w))
            {
                line.Weight = w;
            }

            return line;
        }

        private static PptHtmlAxisExtras AxisExtrasFromSnap(AxisStyleSnap ax)
        {
            if (ax == null)
            {
                return null;
            }

            var extras = new PptHtmlAxisExtras();
            if (ax.Deleted == true)
            {
                extras.Visible = "false";
            }
            else if (ax.Deleted == false)
            {
                extras.Visible = "true";
            }

            extras.TickFont = ax.TickFontName;
            if (ax.TickFontColor.HasValue)
            {
                extras.TickColor = OfficeRgbToHex(ax.TickFontColor.Value);
            }

            if (ax.TickFontSize.HasValue)
            {
                extras.TickSize = ax.TickFontSize.Value.ToString("0.##", CultureInfo.InvariantCulture);
            }

            if (ax.TickLabelPosition.HasValue)
            {
                extras.TickPosition = TickPosFromXl(ax.TickLabelPosition.Value);
            }

            if (ax.MajorTickMark.HasValue)
            {
                extras.MajorTick = TickMarkFromXl(ax.MajorTickMark.Value);
            }

            if (ax.MinorTickMark.HasValue)
            {
                extras.MinorTick = TickMarkFromXl(ax.MinorTickMark.Value);
            }

            extras.Format = ax.NumberFormat;
            if (ax.HasMajorGridlines.HasValue)
            {
                extras.Grid = ax.HasMajorGridlines.Value ? "true" : "false";
            }

            if (ax.MajorGridlineRgb.HasValue)
            {
                extras.GridColor = OfficeRgbToHex(ax.MajorGridlineRgb.Value);
            }

            if (ax.LineVisible == false)
            {
                extras.Line = "none";
            }
            else if (ax.LineRgb.HasValue)
            {
                extras.Line = OfficeRgbToHex(ax.LineRgb.Value);
            }

            if (ax.LineWeight.HasValue && !IsPhantomWeight(ax.LineWeight.Value))
            {
                extras.LineWeight = ax.LineWeight.Value.ToString("0.##", CultureInfo.InvariantCulture);
            }

            return extras;
        }

        private static AxisStyleSnap AxisSnapFromExtras(PptHtmlAxisExtras extras, string title, string formatFallback)
        {
            if (extras == null && title == null && string.IsNullOrWhiteSpace(formatFallback))
            {
                return null;
            }

            extras = extras ?? new PptHtmlAxisExtras();
            var ax = new AxisStyleSnap();
            if (!string.IsNullOrWhiteSpace(extras.Visible))
            {
                ax.Deleted = !IsTrue(extras.Visible);
            }

            if (title != null)
            {
                ax.HasTitle = !string.IsNullOrEmpty(title);
                ax.Title = title;
            }

            ax.TickFontName = extras.TickFont;
            if (TryParseHexToOffice(extras.TickColor, out int tickRgb))
            {
                ax.TickFontColor = tickRgb;
            }

            if (double.TryParse(extras.TickSize, NumberStyles.Float, CultureInfo.InvariantCulture, out double tsz))
            {
                ax.TickFontSize = tsz;
            }

            if (!string.IsNullOrWhiteSpace(extras.TickPosition))
            {
                ax.TickLabelPosition = TickPosToXl(extras.TickPosition);
            }

            if (!string.IsNullOrWhiteSpace(extras.MajorTick))
            {
                ax.MajorTickMark = TickMarkToXl(extras.MajorTick);
            }

            if (!string.IsNullOrWhiteSpace(extras.MinorTick))
            {
                ax.MinorTickMark = TickMarkToXl(extras.MinorTick);
            }

            ax.NumberFormat = !string.IsNullOrWhiteSpace(extras.Format) ? extras.Format : formatFallback;
            if (!string.IsNullOrWhiteSpace(extras.Grid))
            {
                ax.HasMajorGridlines = IsTrue(extras.Grid);
            }

            if (TryParseHexToOffice(extras.GridColor, out int gRgb))
            {
                ax.MajorGridlineRgb = gRgb;
            }

            if (string.Equals(extras.Line, "none", StringComparison.OrdinalIgnoreCase))
            {
                ax.LineVisible = false;
            }
            else if (TryParseHexToOffice(extras.Line, out int lRgb))
            {
                ax.LineVisible = true;
                ax.LineRgb = lRgb;
            }

            if (double.TryParse(extras.LineWeight, NumberStyles.Float, CultureInfo.InvariantCulture, out double lw))
            {
                ax.LineWeight = lw;
            }

            return ax;
        }

        private static string AreaColorFromSnap(bool? visible, int? rgb)
        {
            if (visible == false)
            {
                return "none";
            }

            if (visible == true && rgb.HasValue)
            {
                return OfficeRgbToHex(rgb.Value);
            }

            return null;
        }

        private static void ApplyAreaColor(string raw, out bool? visible, out int? rgb)
        {
            visible = null;
            rgb = null;
            if (string.IsNullOrWhiteSpace(raw))
            {
                return;
            }

            if (string.Equals(raw, "none", StringComparison.OrdinalIgnoreCase)
                || string.Equals(raw, "transparent", StringComparison.OrdinalIgnoreCase))
            {
                visible = false;
                return;
            }

            if (TryParseHexToOffice(raw, out int parsed))
            {
                visible = true;
                rgb = parsed;
            }
        }

        private static string BoxFromSnap(float? left, float? top, float? width, float? height)
        {
            if (!left.HasValue || !top.HasValue || !width.HasValue || !height.HasValue)
            {
                return null;
            }

            return left.Value.ToString("0.##", CultureInfo.InvariantCulture)
                + "," + top.Value.ToString("0.##", CultureInfo.InvariantCulture)
                + "," + width.Value.ToString("0.##", CultureInfo.InvariantCulture)
                + "," + height.Value.ToString("0.##", CultureInfo.InvariantCulture);
        }

        private static void ParseBox(string raw, out float? left, out float? top, out float? width, out float? height)
        {
            left = top = width = height = null;
            if (string.IsNullOrWhiteSpace(raw))
            {
                return;
            }

            string[] parts = raw.Split(new[] { ',', ' ' }, StringSplitOptions.RemoveEmptyEntries);
            if (parts.Length < 4)
            {
                return;
            }

            if (float.TryParse(parts[0], NumberStyles.Float, CultureInfo.InvariantCulture, out float l)
                && float.TryParse(parts[1], NumberStyles.Float, CultureInfo.InvariantCulture, out float t)
                && float.TryParse(parts[2], NumberStyles.Float, CultureInfo.InvariantCulture, out float w)
                && float.TryParse(parts[3], NumberStyles.Float, CultureInfo.InvariantCulture, out float h))
            {
                left = l;
                top = t;
                width = w;
                height = h;
            }
        }

        private static string EncodeGradient(List<GradientStopSnap> stops)
        {
            if (stops == null || stops.Count == 0)
            {
                return null;
            }

            var sb = new StringBuilder();
            for (int i = 0; i < stops.Count; i++)
            {
                if (i > 0)
                {
                    sb.Append(';');
                }

                GradientStopSnap s = stops[i];
                sb.Append(s.Position.ToString("0.##", CultureInfo.InvariantCulture))
                    .Append(':')
                    .Append(OfficeRgbToHex(s.Rgb))
                    .Append('@')
                    .Append(s.Transparency.ToString("0.##", CultureInfo.InvariantCulture));
            }

            return sb.ToString();
        }

        private static List<GradientStopSnap> DecodeGradient(string raw)
        {
            if (string.IsNullOrWhiteSpace(raw))
            {
                return null;
            }

            var list = new List<GradientStopSnap>();
            foreach (string part in raw.Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries))
            {
                int colon = part.IndexOf(':');
                int at = part.LastIndexOf('@');
                if (colon <= 0)
                {
                    continue;
                }

                if (!double.TryParse(part.Substring(0, colon).Trim(), NumberStyles.Float, CultureInfo.InvariantCulture, out double pos))
                {
                    continue;
                }

                string hex = at > colon
                    ? part.Substring(colon + 1, at - colon - 1).Trim()
                    : part.Substring(colon + 1).Trim();
                if (!TryParseHexToOffice(hex, out int rgb))
                {
                    continue;
                }

                double trans = 0;
                if (at > colon)
                {
                    double.TryParse(part.Substring(at + 1).Trim(), NumberStyles.Float, CultureInfo.InvariantCulture, out trans);
                }

                list.Add(new GradientStopSnap { Position = pos, Rgb = rgb, Transparency = trans });
            }

            return list.Count == 0 ? null : list;
        }

        private static bool TryParseSeriesXl(string seriesType, string marker, out int xl)
        {
            xl = XlColumnClustered;
            if (!string.IsNullOrWhiteSpace(seriesType))
            {
                if (int.TryParse(seriesType.Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out int raw)
                    && raw != 0)
                {
                    xl = raw;
                    return true;
                }

                if (TryParseType(seriesType, out xl, out _, out _))
                {
                    if (xl == XlLine && !string.IsNullOrWhiteSpace(marker)
                        && !string.Equals(marker, "none", StringComparison.OrdinalIgnoreCase))
                    {
                        xl = XlLineMarkers;
                    }

                    return true;
                }

                return false;
            }

            if (!string.IsNullOrWhiteSpace(marker)
                && !string.Equals(marker, "none", StringComparison.OrdinalIgnoreCase))
            {
                xl = XlLineMarkers;
                return true;
            }

            return false;
        }

        private static string SeriesTypeFromXl(int xl)
        {
            if (xl == XlBarClustered)
            {
                return "bar";
            }

            if (xl == XlLine || xl == XlLineMarkers)
            {
                return "line";
            }

            if (xl == Xl3DPie)
            {
                return "pie3d";
            }

            if (xl == XlPie)
            {
                return "pie2d";
            }

            if (xl == XlColumnClustered)
            {
                return "column";
            }

            return xl.ToString(CultureInfo.InvariantCulture);
        }

        private static string MarkerFromXl(int style)
        {
            switch (style)
            {
                case -4142:
                    return "none";
                case 8:
                    return "circle";
                case 2:
                    return "diamond";
                case 1:
                    return "square";
                case 3:
                    return "triangle";
                case 5:
                    return "star";
                case 9:
                    return "plus";
                case -4168:
                    return "x";
                case -4118:
                    return "dot";
                default:
                    return style.ToString(CultureInfo.InvariantCulture);
            }
        }

        private static int MarkerToXl(string raw)
        {
            switch ((raw ?? "").Trim().ToLowerInvariant())
            {
                case "none":
                    return -4142;
                case "circle":
                    return 8;
                case "diamond":
                    return 2;
                case "square":
                    return 1;
                case "triangle":
                    return 3;
                case "star":
                    return 5;
                case "plus":
                    return 9;
                case "x":
                    return -4168;
                case "dot":
                    return -4118;
                default:
                    return int.TryParse(raw, NumberStyles.Integer, CultureInfo.InvariantCulture, out int n)
                        ? n
                        : 8;
            }
        }

        private static string TickPosFromXl(int pos)
        {
            if (pos == XlTickLabelPositionNone)
            {
                return "none";
            }

            if (pos == -4134)
            {
                return "low";
            }

            if (pos == -4127)
            {
                return "high";
            }

            return "next";
        }

        private static int TickPosToXl(string raw)
        {
            switch ((raw ?? "").Trim().ToLowerInvariant())
            {
                case "none":
                    return XlTickLabelPositionNone;
                case "low":
                    return -4134;
                case "high":
                    return -4127;
                default:
                    return XlTickLabelPositionNextToAxis;
            }
        }

        private static string TickMarkFromXl(int mark)
        {
            if (mark == XlTickMarkNone)
            {
                return "none";
            }

            if (mark == 2)
            {
                return "inside";
            }

            if (mark == 4)
            {
                return "cross";
            }

            return "outside";
        }

        private static int TickMarkToXl(string raw)
        {
            switch ((raw ?? "").Trim().ToLowerInvariant())
            {
                case "none":
                    return XlTickMarkNone;
                case "inside":
                    return 2;
                case "cross":
                    return 4;
                default:
                    return 3;
            }
        }

        private static string LabelPosFromXl(int pos)
        {
            switch (pos)
            {
                case -4108:
                    return "center";
                case 0:
                    return "above";
                case 1:
                    return "below";
                case -4131:
                    return "left";
                case -4152:
                    return "right";
                case 2:
                    return "outside";
                case 3:
                    return "inside-end";
                case 4:
                    return "inside-base";
                default:
                    return pos.ToString(CultureInfo.InvariantCulture);
            }
        }

        private static int LabelPosToXl(string raw)
        {
            switch ((raw ?? "").Trim().ToLowerInvariant())
            {
                case "center":
                    return -4108;
                case "above":
                    return 0;
                case "below":
                    return 1;
                case "left":
                    return -4131;
                case "right":
                    return -4152;
                case "outside":
                    return 2;
                case "inside-end":
                    return 3;
                case "inside-base":
                    return 4;
                default:
                    return int.TryParse(raw, NumberStyles.Integer, CultureInfo.InvariantCulture, out int n)
                        ? n
                        : 0;
            }
        }

        private static bool IsPhantomWeight(double weight)
        {
            return weight < -1000 || weight > 1000;
        }

        public static bool TryParseType(string raw, out int xlType, out string canonical, out string error)
        {
            xlType = XlColumnClustered;
            canonical = "column";
            error = null;
            if (string.IsNullOrWhiteSpace(raw))
            {
                return true;
            }

            switch (raw.Trim().ToLowerInvariant())
            {
                case "column":
                case "column_clustered":
                    xlType = XlColumnClustered;
                    canonical = "column";
                    return true;
                case "bar":
                    xlType = XlBarClustered;
                    canonical = "bar";
                    return true;
                case "line":
                    xlType = XlLine;
                    canonical = "line";
                    return true;
                case "pie":
                case "pie2d":
                    xlType = XlPie;
                    canonical = "pie2d";
                    return true;
                case "pie3d":
                    xlType = Xl3DPie;
                    canonical = "pie3d";
                    return true;
                default:
                    error = "data-chart-type 仅支持 column/bar/line/pie2d/pie3d（pie 为 pie2d 别名）";
                    return false;
            }
        }

        public static string CanonicalTypeFromXl(int xlType)
        {
            if (xlType == XlBarClustered)
            {
                return "bar";
            }

            if (xlType == XlLine || xlType == XlLineMarkers)
            {
                return "line";
            }

            if (xlType == Xl3DPie)
            {
                return "pie3d";
            }

            if (xlType == XlPie)
            {
                return "pie2d";
            }

            return "column";
        }

        private static bool IsPieXl(int xlType)
        {
            return xlType == XlPie || xlType == Xl3DPie;
        }

        private static bool IsPieChart(object chart)
        {
            if (chart == null)
            {
                return false;
            }

            try
            {
                object t = WppCom.GetProperty(chart, "ChartType");
                return t != null && IsPieXl(Convert.ToInt32(t));
            }
            catch (Exception)
            {
                return false;
            }
        }

        private static bool TryReadChartXl(object chart, out int xlType)
        {
            xlType = XlColumnClustered;
            if (chart == null)
            {
                return false;
            }

            try
            {
                object t = WppCom.GetProperty(chart, "ChartType");
                if (t == null)
                {
                    return false;
                }

                xlType = Convert.ToInt32(t);
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        private static bool IsBuildableAddChartXl(int xlType)
        {
            switch (xlType)
            {
                case XlColumnClustered:
                case XlBarClustered:
                case XlLine:
                case XlLineMarkers:
                case XlPie:
                case Xl3DPie:
                    return true;
                default:
                    return false;
            }
        }

        private static string NormalizeSeriesTypeHint(string raw)
        {
            if (string.IsNullOrWhiteSpace(raw))
            {
                return null;
            }

            string s = raw.Trim().ToLowerInvariant();
            if (s == "pie")
            {
                return "pie2d";
            }

            if (s == "column_clustered")
            {
                return "column";
            }

            return s;
        }

        private static int BaseTypePreferenceRank(string normalizedHint)
        {
            switch (normalizedHint)
            {
                case "column":
                    return 0;
                case "bar":
                    return 1;
                case "line":
                    return 2;
                case "pie2d":
                    return 3;
                case "pie3d":
                    return 4;
                default:
                    return 10;
            }
        }

        private static string PickDeriveBaseSeriesType(IList<string> hintsInOrder)
        {
            if (hintsInOrder == null || hintsInOrder.Count == 0)
            {
                return null;
            }

            int bestRank = int.MaxValue;
            string best = null;
            foreach (string raw in hintsInOrder)
            {
                string hint = NormalizeSeriesTypeHint(raw);
                if (hint == null)
                {
                    continue;
                }

                int rank = BaseTypePreferenceRank(hint);
                if (rank < bestRank)
                {
                    bestRank = rank;
                    best = hint;
                }
            }

            return best;
        }

        private static List<string> CollectGridSeriesTypeHints(PptHtmlChartGrid grid)
        {
            var list = new List<string>();
            if (grid?.Columns == null)
            {
                return list;
            }

            foreach (PptHtmlChartColumn col in grid.Columns)
            {
                if (col == null || col.Role == "category")
                {
                    continue;
                }

                if (!string.IsNullOrWhiteSpace(col.SeriesType))
                {
                    list.Add(col.SeriesType);
                }
            }

            return list;
        }

        private static List<string> CollectSnapSeriesTypeHints(ChartStyleSnap snap)
        {
            var list = new List<string>();
            if (snap?.Series == null)
            {
                return list;
            }

            foreach (SeriesStyleSnap one in snap.Series)
            {
                if (one?.ChartType.HasValue == true)
                {
                    list.Add(SeriesTypeFromXl(one.ChartType.Value));
                }
            }

            return list;
        }

        private static bool TryDeriveAddChartBaseType(
            bool gridFromHtml,
            PptHtmlChartGrid grid,
            ChartStyleSnap oldSnap,
            out int xlType,
            out string canonical,
            out string deriveSource,
            out string error)
        {
            xlType = XlColumnClustered;
            canonical = "column";
            deriveSource = "default";
            error = null;

            string hint = null;
            if (gridFromHtml)
            {
                hint = PickDeriveBaseSeriesType(CollectGridSeriesTypeHints(grid));
                if (!string.IsNullOrEmpty(hint))
                {
                    deriveSource = "html-grid";
                }
            }

            if (string.IsNullOrEmpty(hint) && oldSnap != null)
            {
                hint = PickDeriveBaseSeriesType(CollectSnapSeriesTypeHints(oldSnap));
                if (!string.IsNullOrEmpty(hint))
                {
                    deriveSource = "old-series";
                }
            }

            if (string.IsNullOrEmpty(hint))
            {
                hint = "column";
                deriveSource = "default";
            }

            if (!TryParseType(hint, out xlType, out canonical, out error))
            {
                if (!TryParseType("column", out xlType, out canonical, out error))
                {
                    return false;
                }

                deriveSource = "default";
            }

            return true;
        }

        /// <summary>
        /// 整图 data-chart-type 只铺底：列上已写 data-series-type 的系列不被盖掉。
        /// </summary>
        private static void PinChartSeriesType(ChartStyleSnap snap, ChartStyleSnap htmlSnap, int xlType)
        {
            if (snap == null || snap.Series == null)
            {
                return;
            }

            for (int i = 0; i < snap.Series.Count; i++)
            {
                if (snap.Series[i] == null)
                {
                    continue;
                }

                if (htmlSnap != null
                    && htmlSnap.Series != null
                    && i < htmlSnap.Series.Count
                    && htmlSnap.Series[i] != null
                    && htmlSnap.Series[i].ChartType.HasValue)
                {
                    continue;
                }

                snap.Series[i].ChartType = xlType;
            }
        }

        /// <summary>
        /// 列上系列绑轴：规范 data-axis=y|y2；兼容旧 data-axis-y / data-axis-y2。
        /// 返回规范槽位 "y" / "y2"，未写返回 null。
        /// </summary>
        private static string ResolveSeriesAxisSlot(XElement cell)
        {
            if (cell == null)
            {
                return null;
            }

            string axis = GetAttr(cell, "data-axis");
            if (!string.IsNullOrWhiteSpace(axis))
            {
                string a = axis.Trim().ToLowerInvariant();
                if (a == "y2" || a == "secondary")
                {
                    return "y2";
                }

                if (a == "y" || a == "primary" || a == "x")
                {
                    // x 不当纵轴次轴；当主轴忽略（系列只挂 y/y2）
                    return a == "x" ? null : "y";
                }
            }

            string ay = GetAttr(cell, "data-axis-y");
            if (!string.IsNullOrWhiteSpace(ay))
            {
                string a = ay.Trim().ToLowerInvariant();
                if (a == "secondary" || a == "y2")
                {
                    return "y2";
                }

                if (a == "primary" || a == "y")
                {
                    return "y";
                }
            }

            string ay2 = GetAttr(cell, "data-axis-y2");
            if (ay2 == null)
            {
                return null;
            }

            if (string.IsNullOrWhiteSpace(ay2))
            {
                return "y2";
            }

            string v = ay2.Trim().ToLowerInvariant();
            if (v == "false" || v == "0" || v == "primary" || v == "y" || v == "no")
            {
                return "y";
            }

            return "y2";
        }

        private static bool IsSeriesAxisY2(string slot)
        {
            if (string.IsNullOrWhiteSpace(slot))
            {
                return false;
            }

            string a = slot.Trim().ToLowerInvariant();
            return a == "y2" || a == "secondary";
        }

        private static bool IsSeriesAxisY(string slot)
        {
            if (string.IsNullOrWhiteSpace(slot))
            {
                return false;
            }

            string a = slot.Trim().ToLowerInvariant();
            return a == "y" || a == "primary";
        }

        private static string CanonicalSeriesAxisSlot(string slot)
        {
            if (IsSeriesAxisY2(slot))
            {
                return "y2";
            }

            if (IsSeriesAxisY(slot))
            {
                return "y";
            }

            return null;
        }

        public static IEnumerable<XElement> EnumerateTableRows(XElement table)
        {
            if (table == null)
            {
                yield break;
            }

            foreach (XElement child in table.Elements())
            {
                string name = child.Name.LocalName;
                if (string.Equals(name, "tr", StringComparison.OrdinalIgnoreCase))
                {
                    yield return child;
                    continue;
                }

                if (string.Equals(name, "thead", StringComparison.OrdinalIgnoreCase)
                    || string.Equals(name, "tbody", StringComparison.OrdinalIgnoreCase)
                    || string.Equals(name, "tfoot", StringComparison.OrdinalIgnoreCase))
                {
                    foreach (XElement tr in child.Elements().Where(e =>
                        string.Equals(e.Name.LocalName, "tr", StringComparison.OrdinalIgnoreCase)))
                    {
                        yield return tr;
                    }
                }
            }
        }

        public static bool TryParseGrid(XElement table, bool enforceLimitFail, out PptHtmlChartGrid grid, out string error)
        {
            grid = null;
            error = null;
            if (table == null)
            {
                error = "新建 chart 必须在节点内嵌 <table> 灌数，不能建空图";
                return false;
            }

            var columns = new List<PptHtmlChartColumn>();
            var dataRows = new List<List<string>>();
            bool first = true;
            foreach (XElement tr in EnumerateTableRows(table))
            {
                var cells = tr.Elements().Where(e =>
                    string.Equals(e.Name.LocalName, "td", StringComparison.OrdinalIgnoreCase)
                    || string.Equals(e.Name.LocalName, "th", StringComparison.OrdinalIgnoreCase)).ToList();
                if (cells.Count == 0)
                {
                    continue;
                }

                if (first)
                {
                    for (int i = 0; i < cells.Count; i++)
                    {
                        string role = GetAttr(cells[i], "data-col");
                        if (string.IsNullOrWhiteSpace(role))
                        {
                            role = i == 0 ? "category" : "value";
                        }

                        role = role.Trim().ToLowerInvariant();
                        if (role != "category" && role != "value")
                        {
                            error = "th data-col 只能是 category 或 value";
                            return false;
                        }

                        columns.Add(new PptHtmlChartColumn
                        {
                            Role = role,
                            Name = InnerText(cells[i]),
                            Color = ParseColorOrNone(GetAttr(cells[i], "data-color")),
                            SeriesType = GetAttr(cells[i], "data-series-type"),
                            AxisY = ResolveSeriesAxisSlot(cells[i]),
                            ShowDataLabels = GetAttr(cells[i], "data-show-data-labels"),
                            ShowValue = GetAttr(cells[i], "data-show-value"),
                            ShowPercentage = GetAttr(cells[i], "data-show-percentage"),
                            FillGradient = GetAttr(cells[i], "data-fill-gradient"),
                            FillAngle = GetAttr(cells[i], "data-fill-angle"),
                            Line = GetAttr(cells[i], "data-line"),
                            LineWeight = GetAttr(cells[i], "data-line-weight"),
                            Marker = GetAttr(cells[i], "data-marker"),
                            MarkerSize = GetAttr(cells[i], "data-marker-size"),
                            MarkerColor = GetAttr(cells[i], "data-marker-color"),
                            MarkerFill = GetAttr(cells[i], "data-marker-fill"),
                            LabelPosition = GetAttr(cells[i], "data-label-position"),
                            LabelFont = GetAttr(cells[i], "data-label-font"),
                            LabelSize = GetAttr(cells[i], "data-label-size"),
                            LabelColor = GetAttr(cells[i], "data-label-color"),
                            LabelFormat = GetAttr(cells[i], "data-label-format")
                        });
                        MarkColumnDataLabelMentions(columns[columns.Count - 1]);
                    }

                    first = false;
                    continue;
                }

                var row = new List<string>();
                for (int i = 0; i < cells.Count; i++)
                {
                    row.Add(InnerText(cells[i]));
                }

                dataRows.Add(row);
            }

            if (columns.Count < 2 || dataRows.Count < 1)
            {
                error = "chart 内嵌表至少要有表头行 + 一行数字。"
                    + "写成 <tr><th>类别</th><th>系列</th></tr><tr><td>Q1</td><td>120</td></tr>；"
                    + "thead/tbody 也可以";
                return false;
            }

            if (!columns.Any(c => c.Role == "value"))
            {
                error = "chart 内嵌表至少要有一列 value";
                return false;
            }

            bool over = columns.Count > MaxCols || (dataRows.Count + 1) > MaxRows;
            if (over && enforceLimitFail)
            {
                error = "chart 内嵌表最多 " + MaxRows + " 行 × " + MaxCols + " 列";
                return false;
            }

            if (over)
            {
                if (columns.Count > MaxCols)
                {
                    columns = columns.GetRange(0, MaxCols);
                }

                int maxData = MaxRows - 1;
                if (dataRows.Count > maxData)
                {
                    dataRows = dataRows.GetRange(0, maxData);
                }
            }

            int colCount = columns.Count;
            for (int r = 0; r < dataRows.Count; r++)
            {
                while (dataRows[r].Count < colCount)
                {
                    dataRows[r].Add("");
                }

                if (dataRows[r].Count > colCount)
                {
                    dataRows[r] = dataRows[r].GetRange(0, colCount);
                }

                for (int c = 0; c < colCount; c++)
                {
                    if (columns[c].Role != "value")
                    {
                        continue;
                    }

                    string raw = dataRows[r][c];
                    if (string.IsNullOrWhiteSpace(raw)
                        || !double.TryParse(raw.Trim(), NumberStyles.Float, CultureInfo.InvariantCulture, out _))
                    {
                        error = "第 " + (r + 2) + " 行第 " + (c + 1) + " 列不是数字";
                        return false;
                    }
                }
            }

            grid = new PptHtmlChartGrid
            {
                Columns = columns,
                Rows = dataRows,
                Truncated = over
            };
            return true;
        }

        public static PptHtmlChartFormat ParseFormat(XElement el)
        {
            if (el == null)
            {
                return new PptHtmlChartFormat();
            }

            var format = new PptHtmlChartFormat
            {
                ChartType = GetAttr(el, "data-chart-type"),
                Title = GetAttr(el, "data-title"),
                Theme = GetAttr(el, "data-theme"),
                Legend = GetAttr(el, "data-legend"),
                ShowDataLabels = GetAttr(el, "data-show-data-labels"),
                ShowValue = GetAttr(el, "data-show-value"),
                ShowPercentage = GetAttr(el, "data-show-percentage"),
                PlotColor = GetAttr(el, "data-plot-color"),
                TitleFontSize = GetAttr(el, "data-title-font-size"),
                TitleFontBold = GetAttr(el, "data-title-font-bold"),
                TitleFontColor = GetAttr(el, "data-title-font-color"),
                Gridlines = GetAttr(el, "data-gridlines"),
                GapWidth = GetAttr(el, "data-gap-width"),
                DataMarkers = GetAttr(el, "data-data-markers"),
                MarkerSize = GetAttr(el, "data-marker-size"),
                ChartLineWeight = GetAttr(el, "data-chart-line-weight"),
                AxisYMin = GetAttr(el, "data-axis-y-min"),
                AxisYMax = GetAttr(el, "data-axis-y-max"),
                AxisYMajorUnit = GetAttr(el, "data-axis-y-major-unit"),
                Explosion = GetAttr(el, "data-explosion"),
                FillMissing = GetAttr(el, "data-fill-missing"),
                AxisX = GetAttr(el, "data-axis-x"),
                AxisXType = GetAttr(el, "data-axis-x-type"),
                AxisXFormat = GetAttr(el, "data-axis-x-format"),
                AxisXTickCount = GetAttr(el, "data-axis-x-tick-count"),
                AxisXTickSpacing = GetAttr(el, "data-axis-x-tick-spacing"),
                AxisXBetween = GetAttr(el, "data-axis-x-between"),
                AxisY = GetAttr(el, "data-axis-y"),
                AxisYSecondary = GetAttr(el, "data-axis-y-secondary"),
                ChartStyle = GetAttr(el, "data-chart-style"),
                LegendFontColor = GetAttr(el, "data-legend-font-color"),
                ChartAreaColor = GetAttr(el, "data-chart-area-color"),
                Overlap = GetAttr(el, "data-overlap"),
                PlotBox = GetAttr(el, "data-plot-box"),
                PlotInside = GetAttr(el, "data-plot-inside"),
                AxisXStyle = ParseAxisExtras(el, "data-axis-x"),
                AxisYStyle = ParseAxisExtras(el, "data-axis-y"),
                AxisY2Style = ParseAxisExtras(el, "data-axis-y2")
            };
            MarkDisplaySwitchMentions(format);
            return format;
        }

        private static PptHtmlAxisExtras ParseAxisExtras(XElement el, string prefix)
        {
            return new PptHtmlAxisExtras
            {
                Visible = GetAttr(el, prefix + "-visible"),
                TickFont = GetAttr(el, prefix + "-tick-font"),
                TickColor = GetAttr(el, prefix + "-tick-color"),
                TickSize = GetAttr(el, prefix + "-tick-size"),
                TickPosition = GetAttr(el, prefix + "-tick-position"),
                MajorTick = GetAttr(el, prefix + "-major-tick"),
                MinorTick = GetAttr(el, prefix + "-minor-tick"),
                Format = prefix == "data-axis-x"
                    ? GetAttr(el, "data-axis-x-format")
                    : GetAttr(el, prefix + "-format"),
                Grid = GetAttr(el, prefix + "-grid"),
                GridColor = GetAttr(el, prefix + "-grid-color"),
                Line = GetAttr(el, prefix + "-line"),
                LineWeight = GetAttr(el, prefix + "-line-weight")
            };
        }

        public static string BuildInnerHtml(PptHtmlChartGrid grid)
        {
            if (grid == null || grid.Columns == null)
            {
                return "";
            }

            var sb = new StringBuilder();
            sb.AppendLine("    <tr>");
            foreach (PptHtmlChartColumn col in grid.Columns)
            {
                sb.Append("      <th data-col=\"").Append(EscapeAttr(col.Role ?? "value")).Append("\"");
                if (!string.IsNullOrEmpty(col.Color))
                {
                    sb.Append(" data-color=\"").Append(EscapeAttr(col.Color)).Append("\"");
                }

                if (!string.IsNullOrEmpty(col.SeriesType))
                {
                    sb.Append(" data-series-type=\"").Append(EscapeAttr(col.SeriesType)).Append("\"");
                }

                string axisSlot = CanonicalSeriesAxisSlot(col.AxisY);
                if (!string.IsNullOrEmpty(axisSlot))
                {
                    sb.Append(" data-axis=\"").Append(EscapeAttr(axisSlot)).Append("\"");
                }

                if (!string.IsNullOrEmpty(col.ShowDataLabels))
                {
                    sb.Append(" data-show-data-labels=\"").Append(EscapeAttr(col.ShowDataLabels)).Append("\"");
                }

                if (!string.IsNullOrEmpty(col.ShowValue))
                {
                    sb.Append(" data-show-value=\"").Append(EscapeAttr(col.ShowValue)).Append("\"");
                }

                if (!string.IsNullOrEmpty(col.ShowPercentage))
                {
                    sb.Append(" data-show-percentage=\"").Append(EscapeAttr(col.ShowPercentage)).Append("\"");
                }

                WriteRawAttr(sb, "data-fill-gradient", col.FillGradient);
                WriteRawAttr(sb, "data-fill-angle", col.FillAngle);
                WriteRawAttr(sb, "data-line", col.Line);
                WriteRawAttr(sb, "data-line-weight", col.LineWeight);
                WriteRawAttr(sb, "data-marker", col.Marker);
                WriteRawAttr(sb, "data-marker-size", col.MarkerSize);
                WriteRawAttr(sb, "data-marker-color", col.MarkerColor);
                WriteRawAttr(sb, "data-marker-fill", col.MarkerFill);
                WriteRawAttr(sb, "data-label-position", col.LabelPosition);
                WriteRawAttr(sb, "data-label-font", col.LabelFont);
                WriteRawAttr(sb, "data-label-size", col.LabelSize);
                WriteRawAttr(sb, "data-label-color", col.LabelColor);
                WriteRawAttr(sb, "data-label-format", col.LabelFormat);

                sb.Append(">").Append(EscapeText(col.Name)).AppendLine("</th>");
            }

            sb.AppendLine("    </tr>");
            if (grid.Rows != null)
            {
                foreach (List<string> row in grid.Rows)
                {
                    sb.Append("    <tr>");
                    int n = grid.Columns.Count;
                    for (int i = 0; i < n; i++)
                    {
                        string cell = row != null && i < row.Count ? row[i] : "";
                        sb.Append("<td>").Append(EscapeText(cell)).Append("</td>");
                    }

                    sb.AppendLine("</tr>");
                }
            }

            return sb.ToString();
        }

        public static void AppendFormatAttrs(StringBuilder sb, PptHtmlChartFormat fmt)
        {
            if (sb == null || fmt == null)
            {
                return;
            }

            WriteAttr(sb, "data-chart-type", fmt.ChartType);
            WriteAttr(sb, "data-title", fmt.Title);
            WriteAttr(sb, "data-theme", fmt.Theme);
            WriteAttr(sb, "data-legend", fmt.Legend);
            WriteAttr(sb, "data-show-data-labels", fmt.ShowDataLabels);
            WriteAttr(sb, "data-show-value", fmt.ShowValue);
            WriteAttr(sb, "data-show-percentage", fmt.ShowPercentage);
            WriteAttr(sb, "data-plot-color", fmt.PlotColor);
            WriteAttr(sb, "data-title-font-size", fmt.TitleFontSize);
            WriteAttr(sb, "data-title-font-bold", fmt.TitleFontBold);
            WriteAttr(sb, "data-title-font-color", fmt.TitleFontColor);
            WriteAttr(sb, "data-gridlines", fmt.Gridlines);
            WriteAttr(sb, "data-gap-width", fmt.GapWidth);
            WriteAttr(sb, "data-data-markers", fmt.DataMarkers);
            WriteAttr(sb, "data-marker-size", fmt.MarkerSize);
            WriteAttr(sb, "data-chart-line-weight", fmt.ChartLineWeight);
            WriteAttr(sb, "data-axis-y-min", fmt.AxisYMin);
            WriteAttr(sb, "data-axis-y-max", fmt.AxisYMax);
            WriteAttr(sb, "data-axis-y-major-unit", fmt.AxisYMajorUnit);
            WriteAttr(sb, "data-explosion", fmt.Explosion);
            WriteAttr(sb, "data-fill-missing", fmt.FillMissing);
            WriteAttr(sb, "data-axis-x", fmt.AxisX);
            WriteAttr(sb, "data-axis-x-type", fmt.AxisXType);
            WriteAttr(sb, "data-axis-x-format", fmt.AxisXFormat);
            WriteAttr(sb, "data-axis-x-tick-count", fmt.AxisXTickCount);
            WriteAttr(sb, "data-axis-x-tick-spacing", fmt.AxisXTickSpacing);
            WriteAttr(sb, "data-axis-x-between", fmt.AxisXBetween);
            WriteAttr(sb, "data-axis-y", fmt.AxisY);
            WriteAttr(sb, "data-axis-y-secondary", fmt.AxisYSecondary);
            WriteAttr(sb, "data-chart-style", fmt.ChartStyle);
            WriteAttr(sb, "data-legend-font-color", fmt.LegendFontColor);
            WriteAttr(sb, "data-chart-area-color", fmt.ChartAreaColor);
            WriteAttr(sb, "data-overlap", fmt.Overlap);
            WriteAttr(sb, "data-plot-box", fmt.PlotBox);
            WriteAttr(sb, "data-plot-inside", fmt.PlotInside);
            WriteAxisExtras(sb, "data-axis-x", fmt.AxisXStyle, skipFormat: true);
            WriteAxisExtras(sb, "data-axis-y", fmt.AxisYStyle, skipFormat: false);
            WriteAxisExtras(sb, "data-axis-y2", fmt.AxisY2Style, skipFormat: false);
        }

        private static void WriteAxisExtras(StringBuilder sb, string prefix, PptHtmlAxisExtras ax, bool skipFormat)
        {
            if (ax == null)
            {
                return;
            }

            WriteAttr(sb, prefix + "-visible", ax.Visible);
            WriteAttr(sb, prefix + "-tick-font", ax.TickFont);
            WriteAttr(sb, prefix + "-tick-color", ax.TickColor);
            WriteAttr(sb, prefix + "-tick-size", ax.TickSize);
            WriteAttr(sb, prefix + "-tick-position", ax.TickPosition);
            WriteAttr(sb, prefix + "-major-tick", ax.MajorTick);
            WriteAttr(sb, prefix + "-minor-tick", ax.MinorTick);
            if (!skipFormat)
            {
                WriteAttr(sb, prefix + "-format", ax.Format);
            }

            WriteAttr(sb, prefix + "-grid", ax.Grid);
            WriteAttr(sb, prefix + "-grid-color", ax.GridColor);
            WriteAttr(sb, prefix + "-line", ax.Line);
            WriteAttr(sb, prefix + "-line-weight", ax.LineWeight);
        }

        private static void WriteRawAttr(StringBuilder sb, string name, string value)
        {
            if (string.IsNullOrEmpty(value))
            {
                return;
            }

            sb.Append(" ").Append(name).Append("=\"").Append(EscapeAttr(value)).Append("\"");
        }

    }
}
