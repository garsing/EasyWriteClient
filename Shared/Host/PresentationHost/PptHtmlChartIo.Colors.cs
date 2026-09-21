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
        private static void TryInheritColors(object chart, ChartStyleSnap snap, List<string> warnings)
        {
            if (chart == null || snap == null)
            {
                return;
            }

            if (IsPieChart(chart))
            {
                object series = GetSeries(chart, 1);
                SeriesStyleSnap one = snap.Series != null && snap.Series.Count > 0
                    ? snap.Series[0]
                    : null;
                int live = TryGetPointCount(series);
                int oldPts = one == null || one.PointFills == null ? 0 : one.PointFills.Count;
                if (one != null && one.PointFills != null && live > 0 && oldPts == live)
                {
                    StyleLog(warnings, "饼图在主题上覆盖扇区色 " + live + " 个");
                    TryApplyPointFills(series, one.PointFills, warnings, "S1");
                    return;
                }

                StyleLog(warnings, "饼图已继承主题，扇区色 " + oldPts + " 个对 " + live + " 瓣，按旧色带扩/收");
                TrySetVaryByCategories(chart, true, warnings);
                List<FillSnap> expanded = ExpandPieSliceFills(one == null ? null : one.PointFills, live);
                if (expanded != null && expanded.Count == live)
                {
                    TryApplyPointFills(series, expanded, warnings, "S1");
                }

                return;
            }

            if (snap.Series == null)
            {
                return;
            }

            for (int i = 0; i < snap.Series.Count; i++)
            {
                object series = GetSeries(chart, i + 1);
                if (series == null)
                {
                    continue;
                }

                TryApplyFill(series, snap.Series[i].Fill, warnings, "S" + (i + 1));
                TryApplyPointFills(series, snap.Series[i].PointFills, warnings, "S" + (i + 1));
            }
        }

        private static void TryApplyChartTheme(object chart, ChartStyleSnap snap)
        {
            if (snap.ChartStyle.HasValue)
            {
                WppCom.TrySetProperty(chart, "ChartStyle", snap.ChartStyle.Value);
            }

            if (snap.ChartColor.HasValue)
            {
                WppCom.TrySetProperty(chart, "ChartColor", snap.ChartColor.Value);
            }
        }

        private static List<FillSnap> ExpandPieSliceFills(List<FillSnap> oldFills, int want)
        {
            if (want <= 0)
            {
                return null;
            }

            var seeds = new List<int>();
            if (oldFills != null)
            {
                for (int i = 0; i < oldFills.Count; i++)
                {
                    if (oldFills[i] != null && oldFills[i].SolidRgb.HasValue)
                    {
                        seeds.Add(oldFills[i].SolidRgb.Value);
                    }
                }
            }

            var fills = new List<FillSnap>();
            if (seeds.Count == 0)
            {
                return null;
            }

            int[] rgb = ExpandOfficeRgbRamp(seeds, want);
            for (int i = 0; i < rgb.Length; i++)
            {
                fills.Add(new FillSnap { Visible = true, SolidRgb = rgb[i] });
            }

            return fills;
        }

        /// <summary>
        /// 旧扇区色当种子：少了沿色带 HSL 插值拉长，多了均匀取样（含首尾）。
        /// 只有 1 个种子时绕色相、拉明度生成可区分的瓣。
        /// </summary>
        private static int[] ExpandOfficeRgbRamp(List<int> seeds, int want)
        {
            var dest = new int[want];
            if (seeds.Count == 1)
            {
                RgbToHsl(OfficeRgbToChannels(seeds[0]), out double h, out double s, out double l);
                for (int i = 0; i < want; i++)
                {
                    double t = want == 1 ? 0 : i / (double)(want - 1);
                    double nh = (h + (t - 0.5) * 0.22 + 1.0) % 1.0;
                    double nl = Clamp01(l + (t - 0.5) * 0.28);
                    dest[i] = ChannelsToOfficeRgb(HslToRgb(nh, Math.Max(0.25, s), nl));
                }

                return dest;
            }

            if (seeds.Count >= want)
            {
                for (int i = 0; i < want; i++)
                {
                    double t = want == 1 ? 0 : i / (double)(want - 1);
                    dest[i] = seeds[(int)Math.Round(t * (seeds.Count - 1))];
                }

                return dest;
            }

            int segments = seeds.Count - 1;
            for (int i = 0; i < want; i++)
            {
                double t = want == 1 ? 0 : i / (double)(want - 1);
                double pos = t * segments;
                int a = (int)Math.Floor(pos);
                if (a >= segments)
                {
                    dest[i] = seeds[seeds.Count - 1];
                    continue;
                }

                double local = pos - a;
                dest[i] = LerpOfficeRgbHsl(seeds[a], seeds[a + 1], local);
            }

            return dest;
        }

        private static int LerpOfficeRgbHsl(int from, int to, double t)
        {
            RgbToHsl(OfficeRgbToChannels(from), out double h1, out double s1, out double l1);
            RgbToHsl(OfficeRgbToChannels(to), out double h2, out double s2, out double l2);
            double dh = h2 - h1;
            if (dh > 0.5)
            {
                dh -= 1;
            }
            else if (dh < -0.5)
            {
                dh += 1;
            }

            double h = (h1 + dh * t + 1.0) % 1.0;
            return ChannelsToOfficeRgb(HslToRgb(h, s1 + (s2 - s1) * t, l1 + (l2 - l1) * t));
        }

        private static int[] OfficeRgbToChannels(int office)
        {
            return new[] { office & 0xFF, (office >> 8) & 0xFF, (office >> 16) & 0xFF };
        }

        private static int ChannelsToOfficeRgb(int[] rgb)
        {
            return (rgb[0] & 0xFF) | ((rgb[1] & 0xFF) << 8) | ((rgb[2] & 0xFF) << 16);
        }

        private static void RgbToHsl(int[] rgb, out double h, out double s, out double l)
        {
            double r = rgb[0] / 255.0;
            double g = rgb[1] / 255.0;
            double b = rgb[2] / 255.0;
            double max = Math.Max(r, Math.Max(g, b));
            double min = Math.Min(r, Math.Min(g, b));
            l = (max + min) / 2.0;
            if (Math.Abs(max - min) < 0.0001)
            {
                h = 0;
                s = 0;
                return;
            }

            double d = max - min;
            s = l > 0.5 ? d / (2.0 - max - min) : d / (max + min);
            if (max == r)
            {
                h = ((g - b) / d + (g < b ? 6 : 0)) / 6.0;
            }
            else if (max == g)
            {
                h = ((b - r) / d + 2) / 6.0;
            }
            else
            {
                h = ((r - g) / d + 4) / 6.0;
            }
        }

        private static int[] HslToRgb(double h, double s, double l)
        {
            s = Clamp01(s);
            l = Clamp01(l);
            if (s <= 0)
            {
                int g = (int)Math.Round(l * 255);
                return new[] { g, g, g };
            }

            double q = l < 0.5 ? l * (1 + s) : l + s - l * s;
            double p = 2 * l - q;
            return new[]
            {
                (int)Math.Round(HueToRgb(p, q, h + 1.0 / 3.0) * 255),
                (int)Math.Round(HueToRgb(p, q, h) * 255),
                (int)Math.Round(HueToRgb(p, q, h - 1.0 / 3.0) * 255)
            };
        }

        private static double HueToRgb(double p, double q, double t)
        {
            if (t < 0)
            {
                t += 1;
            }

            if (t > 1)
            {
                t -= 1;
            }

            if (t < 1.0 / 6.0)
            {
                return p + (q - p) * 6 * t;
            }

            if (t < 0.5)
            {
                return q;
            }

            if (t < 2.0 / 3.0)
            {
                return p + (q - p) * (2.0 / 3.0 - t) * 6;
            }

            return p;
        }

        private static double Clamp01(double x)
        {
            if (x < 0)
            {
                return 0;
            }

            return x > 1 ? 1 : x;
        }

        /// <summary>
        /// 仅新建饼图：稿未写系列色时按类别分色。改已有图不走这里。
        /// </summary>
        private static void EnsureNewPieVariesByCategory(
            object chart,
            PptHtmlChartGrid grid,
            List<string> warnings)
        {
            if (!IsPieChart(chart) || HtmlGridHasSeriesColor(grid))
            {
                return;
            }

            TrySetVaryByCategories(chart, true, warnings);
            StyleLog(warnings, "新建饼图未写系列色，按类别分色");
        }

        private static bool HtmlGridHasSeriesColor(PptHtmlChartGrid grid)
        {
            if (grid == null || grid.Columns == null)
            {
                return false;
            }

            for (int i = 0; i < grid.Columns.Count; i++)
            {
                PptHtmlChartColumn col = grid.Columns[i];
                if (col == null || col.Role == "category")
                {
                    continue;
                }

                if (!string.IsNullOrWhiteSpace(col.Color))
                {
                    return true;
                }
            }

            return false;
        }

        private static void TrySetVaryByCategories(object chart, bool on, List<string> warnings = null)
        {
            try
            {
                object g = TryInvoke(chart, "ChartGroups", 1);
                WppCom.TrySetProperty(g, "VaryByCategories", on);
                object live = g == null ? null : WppCom.GetProperty(g, "VaryByCategories");
                StyleLog(warnings, "VaryByCategories=" + on + " live=" + live);
            }
            catch (Exception ex)
            {
                StyleLog(warnings, "VaryByCategories 异常: " + ex.Message);
            }
        }

        /// <summary>
        /// 饼图已是饼则不改 2D↔3D。未写类型时建图已跟旧图；这里只兜底「建成了非饼、快照却是饼」。
        /// </summary>
        private static void TryRestorePieChartType(object chart, ChartStyleSnap snap)
        {
            if (chart == null || snap == null || snap.Series == null)
            {
                return;
            }

            if (IsPieChart(chart))
            {
                return;
            }

            int? want = null;
            for (int i = 0; i < snap.Series.Count; i++)
            {
                if (snap.Series[i].ChartType.HasValue && IsPieXl(snap.Series[i].ChartType.Value))
                {
                    want = snap.Series[i].ChartType.Value;
                    break;
                }
            }

            if (!want.HasValue)
            {
                return;
            }

            try
            {
                WppCom.TrySetProperty(chart, "ChartType", want.Value);
            }
            catch (Exception)
            {
            }
        }

        private static List<string> CategoryLabels(PptHtmlChartGrid grid)
        {
            var cats = new List<string>();
            if (grid == null || grid.Rows == null)
            {
                return cats;
            }

            int catCol = 0;
            if (grid.Columns != null)
            {
                for (int i = 0; i < grid.Columns.Count; i++)
                {
                    if (grid.Columns[i].Role == "category")
                    {
                        catCol = i;
                        break;
                    }
                }
            }

            foreach (List<string> row in grid.Rows)
            {
                cats.Add(row != null && catCol < row.Count ? (row[catCol] ?? "") : "");
            }

            return cats;
        }

        private static void EnsureCategoryAxisLabels(object chart, PptHtmlChartGrid grid)
        {
            if (chart == null)
            {
                return;
            }

            try
            {
                object axis = TryInvoke(chart, "Axes", XlCategory, XlPrimary);
                if (axis == null)
                {
                    return;
                }

                // 2018/2019 这类标签灌进 XValues 后常被收成时间轴，横坐标字会空白
                WppCom.TrySetProperty(axis, "CategoryType", XlCategoryScale);

                List<string> cats = CategoryLabels(grid);
                if (cats.Count > 0)
                {
                    WppCom.TrySetProperty(axis, "CategoryNames", cats.ToArray());
                }

                object pos = WppCom.GetProperty(axis, "TickLabelPosition");
                if (pos == null || Convert.ToInt32(pos) == XlTickLabelPositionNone)
                {
                    WppCom.TrySetProperty(axis, "TickLabelPosition", XlTickLabelPositionNextToAxis);
                }

                try
                {
                    object ticks = WppCom.GetProperty(axis, "TickLabels");
                    WppCom.TrySetProperty(ticks, "NumberFormat", "@");
                }
                catch (Exception)
                {
                }
            }
            catch (Exception)
            {
            }
        }

        private static void TryCaptureTickLabelFont(object chart, out object color, out object size)
        {
            color = null;
            size = null;
            try
            {
                object axis = TryInvoke(chart, "Axes", XlCategory, XlPrimary);
                object ticks = axis == null ? null : WppCom.GetProperty(axis, "TickLabels");
                object font = ticks == null ? null : WppCom.GetProperty(ticks, "Font");
                if (font == null)
                {
                    return;
                }

                color = WppCom.GetProperty(font, "Color");
                size = WppCom.GetProperty(font, "Size");
            }
            catch (Exception)
            {
            }
        }

        private static void TryRestoreTickLabelFont(object chart, object color, object size)
        {
            if (color == null && size == null)
            {
                return;
            }

            try
            {
                object axis = TryInvoke(chart, "Axes", XlCategory, XlPrimary);
                object ticks = axis == null ? null : WppCom.GetProperty(axis, "TickLabels");
                object font = ticks == null ? null : WppCom.GetProperty(ticks, "Font");
                if (font == null)
                {
                    return;
                }

                if (color != null)
                {
                    WppCom.TrySetProperty(font, "Color", color);
                }

                if (size != null)
                {
                    WppCom.TrySetProperty(font, "Size", size);
                }
            }
            catch (Exception)
            {
            }
        }

        private static bool TryEnsureSeriesCount(object chart, int wantSeries, out string error)
        {
            error = null;
            if (wantSeries < 1)
            {
                return true;
            }

            try
            {
                int count = GetSeriesCount(chart);
                for (int i = count; i < wantSeries; i++)
                {
                    object sc = TryInvoke(chart, "SeriesCollection");
                    if (sc == null)
                    {
                        sc = WppCom.GetProperty(chart, "SeriesCollection");
                    }

                    if (TryInvoke(sc, "NewSeries") == null && GetSeriesCount(chart) <= i)
                    {
                        error = "无法把图表扩到 " + wantSeries + " 个系列";
                        return false;
                    }
                }

                return GetSeriesCount(chart) >= wantSeries;
            }
            catch (Exception ex)
            {
                error = "扩展图表系列失败: " + ex.Message;
                return false;
            }
        }

        private static bool TryPourSeriesLikeWord(
            object chart,
            PptHtmlChartGrid grid,
            out string error,
            List<string> warnings = null)
        {
            return TryAssignGridArrays(chart, grid, warnings, out error);
        }

        private static bool TryAssignGridArrays(
            object chart,
            PptHtmlChartGrid grid,
            List<string> warnings,
            out string error)
        {
            error = null;
            if (chart == null || grid == null || grid.Rows == null)
            {
                error = "无法按数组灌系列";
                return false;
            }

            var cats = new List<string>();
            foreach (List<string> row in grid.Rows)
            {
                cats.Add(row != null && row.Count > 0 ? (row[0] ?? "") : "");
            }

            int si = 0;
            for (int c = 0; c < grid.Columns.Count; c++)
            {
                if (grid.Columns[c].Role == "category")
                {
                    continue;
                }

                si++;
                object series = GetSeries(chart, si);
                if (series == null)
                {
                    error = "无法访问图表系列 " + si;
                    return false;
                }

                var values = new List<double>();
                foreach (List<string> row in grid.Rows)
                {
                    string raw = row != null && c < row.Count ? row[c] : "";
                    if (!double.TryParse(raw, NumberStyles.Float, CultureInfo.InvariantCulture, out double n))
                    {
                        n = 0;
                    }

                    values.Add(n);
                }

                if (!TryAssignSeriesArrayShapes(series, values, cats, warnings, out error))
                {
                    return false;
                }

                if (!string.IsNullOrEmpty(grid.Columns[c].Name))
                {
                    WppCom.TrySetProperty(series, "Name", grid.Columns[c].Name);
                }
            }

            if (ReadSeriesRowCount(chart) != grid.Rows.Count)
            {
                error = "系列点数 " + ReadSeriesRowCount(chart) + " 与稿行数 " + grid.Rows.Count + " 不一致";
                return false;
            }

            return true;
        }

        private static bool TryAssignSeriesArrayShapes(
            object series,
            List<double> values,
            List<string> cats,
            List<string> warnings,
            out string error)
        {
            error = null;
            int n = values.Count;
            var yCol = new object[n, 1];
            var yRow = new object[1, n];
            var xCol = new object[n, 1];
            var xRow = new object[1, n];
            var yObj = new object[n];
            var xObj = new object[n];
            for (int i = 0; i < n; i++)
            {
                yCol[i, 0] = values[i];
                yRow[0, i] = values[i];
                yObj[i] = values[i];
                string cat = i < cats.Count ? cats[i] : "";
                xCol[i, 0] = cat;
                xRow[0, i] = cat;
                xObj[i] = cat;
            }

            object[] yCandidates = { yCol, yRow, yObj, values.ToArray() };
            object[] xCandidates = { xCol, xRow, xObj, cats.ToArray() };
            string[] tags = { "col2d", "row2d", "obj1d", "dbl1d" };
            for (int i = 0; i < yCandidates.Length; i++)
            {
                if (TryAssignSeriesValues(series, yCandidates[i], xCandidates[i], out string bindErr))
                {
                    int pts = TryGetPointCount(series);
                    int yn = ReadSeriesValuesCount(series);
                    PourLog(warnings, "数组 " + tags[i] + " y=" + DescribeComArg(yCandidates[i])
                        + " 后 pts=" + pts + " Yn=" + yn);
                    if (pts >= n || yn >= n)
                    {
                        return true;
                    }
                }
                else
                {
                    PourLog(warnings, "数组 " + tags[i] + " 失败 y=" + DescribeComArg(yCandidates[i])
                        + " | " + bindErr);
                }
            }

            error = "各形状数组都未能把系列扩到 " + n + " 点";
            return false;
        }

        private static int ReadSeriesValuesCount(object series)
        {
            try
            {
                return ToStringList(WppCom.GetProperty(series, "Values")).Count;
            }
            catch (Exception)
            {
                return 0;
            }
        }

        private static bool TrimExtraSeries(object chart, int wantSeries, out string error)
        {
            error = null;
            if (wantSeries < 1)
            {
                return true;
            }

            try
            {
                int total = GetSeriesCount(chart);
                for (int i = total; i > wantSeries; i--)
                {
                    try
                    {
                        object series = GetSeries(chart, i);
                        if (series != null)
                        {
                            WppCom.Invoke(series, "Delete");
                        }
                    }
                    catch (Exception)
                    {
                    }
                }

                return true;
            }
            catch (Exception ex)
            {
                error = "删除图表多余系列失败: " + ex.Message;
                return false;
            }
        }

        private static void TryClearSheet(object ws)
        {
            try
            {
                object block = TryGetExcelRange(ws, "A1", "H40");
                if (block == null)
                {
                    block = TryGetExcelRange(ws, "A1:H40", null);
                }

                TryInvoke(block, "Clear");
                TryInvoke(block, "ClearContents");
            }
            catch (Exception)
            {
            }

            try
            {
                object used = WppCom.GetProperty(ws, "UsedRange");
                TryInvoke(used, "Clear");
            }
            catch (Exception)
            {
            }
        }

        private static object GetCellObject(object ws, int row, int col)
        {
            try
            {
                object cells = WppCom.GetProperty(ws, "Cells");
                if (cells != null)
                {
                    return cells.GetType().InvokeMember(
                        "Item",
                        BindingFlags.GetProperty
                        | BindingFlags.InvokeMethod
                        | BindingFlags.Instance
                        | BindingFlags.Public,
                        null,
                        cells,
                        new object[] { row, col });
                }
            }
            catch (Exception)
            {
            }

            try
            {
                return WppCom.Invoke(ws, "Cells", row, col);
            }
            catch (Exception)
            {
                return null;
            }
        }

        private static void SetCell(object ws, int row, int col, object value)
        {
            object cell = GetCellObject(ws, row, col);
            WppCom.TrySetProperty(cell, "Value", value);
            WppCom.TrySetProperty(cell, "Value2", value);
        }

        private static object GetCell(object ws, int row, int col)
        {
            try
            {
                object cell = GetCellObject(ws, row, col);
                object v = WppCom.GetProperty(cell, "Value2");
                return v ?? WppCom.GetProperty(cell, "Value");
            }
            catch (Exception)
            {
                return null;
            }
        }

        private static object TryInvoke(object target, string name, params object[] args)
        {
            if (target == null)
            {
                return null;
            }

            try
            {
                return WppCom.Invoke(target, name, args);
            }
            catch (Exception)
            {
                return null;
            }
        }

        private static void TryDelete(object shape)
        {
            try
            {
                WppCom.Invoke(shape, "Delete");
            }
            catch (Exception)
            {
            }
        }

        private static List<string> ToStringList(object values)
        {
            var list = new List<string>();
            if (values == null)
            {
                return list;
            }

            if (values is Array arr)
            {
                foreach (object item in arr)
                {
                    list.Add(FormatCell(item));
                }

                return list;
            }

            list.Add(FormatCell(values));
            return list;
        }

        private static string FormatCell(object v)
        {
            if (v == null || v is DBNull)
            {
                return "";
            }

            if (v is double d)
            {
                return d.ToString("0.####", CultureInfo.InvariantCulture);
            }

            if (v is float f)
            {
                return f.ToString("0.####", CultureInfo.InvariantCulture);
            }

            if (v is decimal m)
            {
                return m.ToString("0.####", CultureInfo.InvariantCulture);
            }

            return Convert.ToString(v, CultureInfo.InvariantCulture) ?? "";
        }

        private static string TryReadSeriesColor(object series)
        {
            try
            {
                object fmt = WppCom.GetProperty(series, "Format");
                object fill = WppCom.GetProperty(fmt, "Fill");
                object vis = fill == null ? null : WppCom.GetProperty(fill, "Visible");
                if (vis != null && Convert.ToInt32(vis) == 0)
                {
                    return "none";
                }

                object fc = WppCom.GetProperty(fill, "ForeColor");
                object rgb = WppCom.GetProperty(fc, "RGB");
                if (rgb != null)
                {
                    return OfficeRgbToHex(Convert.ToInt32(Convert.ToDouble(rgb)));
                }
            }
            catch (Exception)
            {
            }

            return null;
        }

        private static bool TryParseHexToOffice(string hex, out int office)
        {
            office = 0;
            string t = NormalizeHexOrNull(hex);
            if (t == null || t.Length != 7)
            {
                return false;
            }

            if (!int.TryParse(t.Substring(1, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out int r)
                || !int.TryParse(t.Substring(3, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out int g)
                || !int.TryParse(t.Substring(5, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out int b))
            {
                return false;
            }

            office = r + (g << 8) + (b << 16);
            return true;
        }

        private static string OfficeRgbToHex(int office)
        {
            int r = office & 0xFF;
            int g = (office >> 8) & 0xFF;
            int b = (office >> 16) & 0xFF;
            return "#" + r.ToString("X2", CultureInfo.InvariantCulture)
                + g.ToString("X2", CultureInfo.InvariantCulture)
                + b.ToString("X2", CultureInfo.InvariantCulture);
        }

        private static string NormalizeHexOrNull(string raw)
        {
            if (string.IsNullOrWhiteSpace(raw))
            {
                return null;
            }

            string t = raw.Trim();
            if (t.StartsWith("#", StringComparison.Ordinal) && t.Length == 7)
            {
                return t.ToUpperInvariant();
            }

            return null;
        }

        private static string ParseColorOrNone(string raw)
        {
            if (string.Equals(raw, "none", StringComparison.OrdinalIgnoreCase)
                || string.Equals(raw, "transparent", StringComparison.OrdinalIgnoreCase))
            {
                return "none";
            }

            return NormalizeHexOrNull(raw);
        }

        private static int LegendToXl(string pos)
        {
            switch ((pos ?? "").Trim().ToLowerInvariant())
            {
                case "top":
                    return -4160;
                case "left":
                    return -4131;
                case "right":
                    return -4152;
                default:
                    return -4107;
            }
        }

        private static string LegendFromXl(int pos)
        {
            if (pos == -4160)
            {
                return "top";
            }

            if (pos == -4131)
            {
                return "left";
            }

            if (pos == -4152)
            {
                return "right";
            }

            return "bottom";
        }

        private static string ColLetter(int col)
        {
            string s = "";
            int n = col;
            while (n > 0)
            {
                int m = (n - 1) % 26;
                s = (char)('A' + m) + s;
                n = (n - 1) / 26;
            }

            return s;
        }

        private static bool IsTruthy(object v)
        {
            if (v == null)
            {
                return false;
            }

            if (v is bool b)
            {
                return b;
            }

            try
            {
                int n = Convert.ToInt32(v);
                return n != 0 && n != -4142;
            }
            catch (Exception)
            {
                return false;
            }
        }

        private static bool IsTrue(string raw)
        {
            return string.Equals(raw, "true", StringComparison.OrdinalIgnoreCase)
                || string.Equals(raw, "1", StringComparison.OrdinalIgnoreCase)
                || string.Equals(raw, "yes", StringComparison.OrdinalIgnoreCase);
        }

        private static void Warn(List<string> warnings, string field)
        {
            warnings?.Add("未能套用 " + field);
        }

        private static void WriteAttr(StringBuilder sb, string name, string value)
        {
            if (string.IsNullOrEmpty(value))
            {
                return;
            }

            sb.Append(" ").Append(name).Append("=\"").Append(EscapeAttr(value)).Append("\"");
        }

        private static string GetAttr(XElement el, string name)
        {
            if (el == null)
            {
                return null;
            }

            XAttribute a = el.Attribute(name);
            if (a == null)
            {
                a = el.Attributes().FirstOrDefault(x =>
                    string.Equals(x.Name.LocalName, name, StringComparison.OrdinalIgnoreCase));
            }

            return a == null ? null : a.Value;
        }

        private static string InnerText(XElement el)
        {
            if (el == null)
            {
                return "";
            }

            return string.Concat(el.DescendantNodes().OfType<XText>().Select(t => t.Value)).Trim();
        }

        private static string EscapeAttr(string text)
        {
            if (string.IsNullOrEmpty(text))
            {
                return "";
            }

            return text.Replace("&", "&amp;").Replace("\"", "&quot;").Replace("<", "&lt;").Replace(">", "&gt;");
        }

        private static string EscapeText(string text)
        {
            if (string.IsNullOrEmpty(text))
            {
                return "";
            }

            return text.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;");
        }
    }
}
