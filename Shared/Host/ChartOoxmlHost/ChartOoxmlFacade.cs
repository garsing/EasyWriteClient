using System;
using System.Collections.Generic;
using System.IO;
using System.Xml.Linq;

namespace WordAddIn1.ChartOoxmlHost
{
    /// <summary>文字 + 演示共用入口。F_* / WordReader 只调这里，不写 ZipArchive。</summary>
    internal static class ChartOoxmlFacade
    {
        public static bool TryReadWordChartParts(object document, out List<ChartPartRef> parts, out string error)
        {
            return WordChartPackageHost.TryReadParts(document, out parts, out error);
        }

        public static string BuildDisplayDataTable(ChartPartRef part)
        {
            if (part == null || part.ChartXml == null)
            {
                return "<Chart>\n  <DataTable>\n  </DataTable>\n</Chart>\n";
            }

            if (!ChartPartXml.TryParseGrid(part.ChartXml, out ChartPartXml.SeriesGrid grid, out _))
            {
                return "<Chart>\n  <DataTable>\n  </DataTable>\n</Chart>\n";
            }

            return ChartPartXml.BuildDisplayChartXml(grid);
        }

        public static bool TryBuildChartConfig(
            object document,
            int ordinal,
            string scope,
            out ChartConfig config,
            out string error)
        {
            config = null;
            if (!TryReadWordChartParts(document, out List<ChartPartRef> parts, out error))
            {
                return false;
            }

            if (ordinal < 0 || parts == null || ordinal >= parts.Count)
            {
                error = "抽出：图下标无效";
                return false;
            }

            ChartPartRef part = parts[ordinal];
            if (!ChartPartXml.TryParseGrid(part.ChartXml, out ChartPartXml.SeriesGrid grid, out error))
            {
                return false;
            }

            string normalized = string.IsNullOrEmpty(scope) ? "all" : scope.Trim().ToLowerInvariant();
            var info = new ChartInfo();
            if (normalized == "format" || normalized == "all")
            {
                info.Type = grid.Type;
                info.Title = grid.Title;
                info.LegendPosition = grid.Legend;
                if (grid.YMin.HasValue)
                {
                    info.YAxisMin = grid.YMin.Value;
                }

                if (grid.YMax.HasValue)
                {
                    info.YAxisMax = grid.YMax.Value;
                }

                if (grid.YMajor.HasValue)
                {
                    info.YAxisMajorUnit = grid.YMajor.Value;
                }

                if (grid.AxisY != null)
                {
                    info.AxisY = grid.AxisY;
                }
                else if (!string.IsNullOrEmpty(grid.AxisYTitle))
                {
                    info.AxisY = new AxisYInfo { Title = grid.AxisYTitle };
                }

                if (info.AxisY != null && !string.IsNullOrEmpty(grid.AxisYTitle))
                {
                    info.AxisY.Title = grid.AxisYTitle;
                }

                if (grid.AxisX != null)
                {
                    info.AxisX = grid.AxisX;
                }
                else if (!string.IsNullOrEmpty(grid.AxisXTitle))
                {
                    info.AxisX = new AxisXInfo { Title = grid.AxisXTitle };
                }

                if (info.AxisX != null && !string.IsNullOrEmpty(grid.AxisXTitle))
                {
                    info.AxisX.Title = grid.AxisXTitle;
                }

                info.Gridlines = grid.Gridlines;
                if (grid.ChartStyle > 0)
                {
                    info.ChartStyle = grid.ChartStyle;
                }

                if (grid.GapWidth.HasValue)
                {
                    info.GapWidth = grid.GapWidth.Value;
                }

                if (grid.Overlap.HasValue)
                {
                    info.Overlap = grid.Overlap.Value;
                    info.HasOverlap = true;
                }

                info.AxisY2 = grid.AxisY2;
                info.PlotColor = grid.PlotColor;
                info.ChartAreaColor = grid.ChartAreaColor;
                if (info.AxisY2 != null && !string.IsNullOrEmpty(grid.AxisY2Title))
                {
                    info.AxisY2.Title = grid.AxisY2Title;
                }
                else if (info.AxisY2 == null && !string.IsNullOrEmpty(grid.AxisY2Title))
                {
                    info.AxisY2 = new AxisYInfo { Title = grid.AxisY2Title };
                }
                if (part.ExtentCx > 0)
                {
                    info.Width = (int)Math.Round(part.ExtentCx / 12700.0);
                }

                if (part.ExtentCy > 0)
                {
                    info.Height = (int)Math.Round(part.ExtentCy / 12700.0);
                }
            }

            if (normalized == "data" || normalized == "all")
            {
                info.SeriesCollection = ToSeriesInfos(grid);
            }

            config = new ChartConfig { Chart = info };
            return true;
        }

        public static bool TryRewriteDataAndCopyBack(
            object destDoc,
            int ordinal,
            ChartPartXml.SeriesGrid grid,
            bool pinColors,
            out string error)
        {
            error = null;
            if (destDoc == null || grid == null)
            {
                error = "改包：参数空";
                return false;
            }

            string fullName = WordChartPackageHost.TryReadFullName(destDoc);
            string stamp = "";
            if (WordChartPackageHost.TryFindChartShape(destDoc, ordinal, out object destShape, out _)
                && WordChartPackageHost.TryReadShapeTitle(destShape, out string title)
                && ChartTitleStampHelper.TryParse(title, out string chartId, out _))
            {
                stamp = ChartTitleStampHelper.Format(chartId, "");
                if (ChartTitleStampHelper.TryParse(title, out _, out string suffix))
                {
                    stamp = ChartTitleStampHelper.Format(chartId, suffix);
                }
            }

            if (!WordChartPackageHost.TrySaveUserDoc(destDoc, out error))
            {
                return false;
            }

            if (!string.IsNullOrEmpty(fullName)
                && !string.Equals(fullName, WordChartPackageHost.TryReadFullName(destDoc), StringComparison.OrdinalIgnoreCase))
            {
                error = "Save 后 FullName 被改走";
                return false;
            }

            if (!TryReadWordChartParts(destDoc, out List<ChartPartRef> parts, out error))
            {
                return false;
            }

            if (ordinal < 0 || parts == null || ordinal >= parts.Count)
            {
                error = "改包：图下标无效";
                return false;
            }

            ChartPartRef part = parts[ordinal];
            if (!ChartPartXml.TryRewriteGrid(part.ChartXml, grid, out error))
            {
                return false;
            }

            List<string> colorSnap = ChartPartXml.SnapshotSeriesColors(part.ChartXml);
            var replacements = new Dictionary<string, byte[]>(StringComparer.OrdinalIgnoreCase);
            string zipName = (part.ChartPartName ?? "").TrimStart('/');
            replacements[zipName] = ChartPartXml.XmlPartBytes(part.ChartXml);

            string temp = WordChartPackageHost.NewTempDocx("改包");
            object copy = null;
            try
            {
                string srcPath = WordChartPackageHost.TryReadFullName(destDoc);
                if (string.IsNullOrEmpty(srcPath) || !File.Exists(srcPath))
                {
                    error = "用户稿路径不可拷";
                    return false;
                }

                File.Copy(srcPath, temp, true);
                if (!ChartPartIndex.TryApplyReplacementsToZip(temp, replacements, out error))
                {
                    return false;
                }

                if (!WordChartPackageHost.TryOpenCopy(destDoc, temp, out copy, out error))
                {
                    return false;
                }

                if (!string.IsNullOrEmpty(stamp)
                    && WordChartPackageHost.TryFindChartShape(copy, ordinal, out object sourceShape, out _))
                {
                    WordChartPackageHost.TryWriteShapeTitle(sourceShape, stamp);
                }

                if (!WordChartPackageHost.TryCopyChartOnDoc(copy, ordinal, out error))
                {
                    return false;
                }

                if (!WordChartPackageHost.TryReplaceChartAtOrdinal(destDoc, ordinal, out error))
                {
                    return false;
                }

                if (pinColors && HasExplicitColor(colorSnap))
                {
                    if (!TryReadWordChartParts(destDoc, out List<ChartPartRef> after, out error)
                        || after == null
                        || ordinal < 0
                        || ordinal >= after.Count)
                    {
                        error = string.IsNullOrEmpty(error) ? "贴回后对不上原位置" : error;
                        return false;
                    }

                    List<string> afterColors = ChartPartXml.SnapshotSeriesColors(after[ordinal].ChartXml);
                    if (!ChartPartXml.ColorsMatch(colorSnap, afterColors))
                    {
                        if (!ChartPartXml.TryPinSeriesColors(after[ordinal].ChartXml, colorSnap))
                        {
                            error = ColorPinError(colorSnap, afterColors, ordinal);
                            return false;
                        }

                        var pin = new Dictionary<string, byte[]>(StringComparer.OrdinalIgnoreCase);
                        pin[(after[ordinal].ChartPartName ?? "").TrimStart('/')] =
                            ChartPartXml.XmlPartBytes(after[ordinal].ChartXml);
                        WordChartPackageHost.TryCloseCopy(copy);
                        copy = null;
                        File.Copy(WordChartPackageHost.TryReadFullName(destDoc), temp, true);
                        if (!ChartPartIndex.TryApplyReplacementsToZip(temp, pin, out error))
                        {
                            return false;
                        }

                        if (!WordChartPackageHost.TryOpenCopy(destDoc, temp, out copy, out error))
                        {
                            return false;
                        }

                        if (!WordChartPackageHost.TryCopyChartOnDoc(copy, ordinal, out error)
                            || !WordChartPackageHost.TryReplaceChartAtOrdinal(destDoc, ordinal, out error))
                        {
                            return false;
                        }

                        if (!TryReadWordChartParts(destDoc, out List<ChartPartRef> pinned, out error)
                            || pinned == null
                            || ordinal >= pinned.Count
                            || !ChartPartXml.ColorsMatch(
                                colorSnap,
                                ChartPartXml.SnapshotSeriesColors(pinned[ordinal].ChartXml)))
                        {
                            List<string> pinnedColors = pinned != null && ordinal >= 0 && ordinal < pinned.Count
                                ? ChartPartXml.SnapshotSeriesColors(pinned[ordinal].ChartXml)
                                : new List<string>();
                            error = ColorPinError(colorSnap, pinnedColors, ordinal);
                            return false;
                        }
                    }
                }

                if (!string.IsNullOrEmpty(stamp))
                {
                    System.Diagnostics.Debug.WriteLine("[ApplyChartWps] stamp ordinal=" + ordinal);
                    if (!WordChartPackageHost.TryFindChartShape(destDoc, ordinal, out object next, out error)
                        || !WordChartPackageHost.TryWriteShapeTitle(next, stamp))
                    {
                        error = "贴回后写回 Title 戳失败: 下标=" + ordinal
                            + (string.IsNullOrEmpty(error) ? "" : " " + error);
                        return false;
                    }
                }

                if (!string.IsNullOrEmpty(fullName)
                    && !string.Equals(fullName, WordChartPackageHost.TryReadFullName(destDoc), StringComparison.OrdinalIgnoreCase))
                {
                    error = "贴回后 FullName 被改走";
                    return false;
                }

                return true;
            }
            finally
            {
                WordChartPackageHost.TryCloseCopy(copy);
                WordChartPackageHost.TryDeleteFile(temp);
            }
        }

        private static bool HasExplicitColor(IList<string> colors)
        {
            if (colors == null)
            {
                return false;
            }

            for (int i = 0; i < colors.Count; i++)
            {
                if (!string.IsNullOrEmpty(colors[i]))
                {
                    return true;
                }
            }

            return false;
        }

        private static string ColorPinError(IList<string> snap, IList<string> after, int ordinal)
        {
            string left = snap == null ? "" : string.Join(",", snap);
            string right = after == null ? "" : string.Join(",", after);
            return "贴回后系列色被换且钉不住: 快照=" + left
                + " 贴回=" + right
                + " 下标=" + ordinal;
        }

        /// <summary>
        /// 套用前对齐已有图：format 只换标题、沿用原系列；系列数不一致则失败且不写。点数可以变，按新稿改缓存。
        /// </summary>
        public static bool TryPrepareApplyGrid(
            object document,
            int ordinal,
            ChartInfo incoming,
            string scope,
            out ChartPartXml.SeriesGrid grid,
            out string error)
        {
            grid = null;
            if (!TryBuildChartConfig(document, ordinal, "all", out ChartConfig existing, out error))
            {
                return false;
            }

            ChartPartXml.SeriesGrid have = GridFromChartInfo(existing != null ? existing.Chart : null);
            ChartPartXml.SeriesGrid want = GridFromChartInfo(incoming);
            string normalized = string.IsNullOrEmpty(scope) ? "all" : scope.Trim().ToLowerInvariant();
            int haveSeries = have.Series != null ? have.Series.Count : 0;
            int havePoints = FirstPointCount(have);
            int wantSeries = want.Series != null ? want.Series.Count : 0;

            System.Diagnostics.Debug.WriteLine(
                "[ApplyChartWps] ordinal=" + ordinal
                + " scope=" + normalized
                + " haveSeries=" + haveSeries
                + " havePoints=" + havePoints
                + " wantSeries=" + wantSeries
                + " wantTitle=" + (want.Title ?? ""));

            if (normalized == "format")
            {
                if (wantSeries > 0 && wantSeries != haveSeries)
                {
                    error = MismatchMessage(haveSeries, wantSeries, havePoints, 0);
                    return false;
                }

                grid = have;
                if (!string.IsNullOrEmpty(want.Title))
                {
                    grid.Title = want.Title;
                }

                return true;
            }

            if (wantSeries == 0)
            {
                error = "套用数据 scope 需要 DataTable 系列数据";
                return false;
            }

            int wantPoints = FirstPointCount(want);
            if (wantSeries != haveSeries)
            {
                error = MismatchMessage(haveSeries, wantSeries, havePoints, wantPoints);
                return false;
            }

            if (normalized == "data" || string.IsNullOrEmpty(want.Title))
            {
                want.Title = have.Title;
            }

            grid = want;
            return true;
        }

        private static int FirstPointCount(ChartPartXml.SeriesGrid grid)
        {
            if (grid == null || grid.Series == null || grid.Series.Count == 0 || grid.Series[0].Values == null)
            {
                return 0;
            }

            return grid.Series[0].Values.Count;
        }

        private static string MismatchMessage(int expectedSeries, int actualSeries, int expectedPoints, int actualPoints)
        {
            return "chart_data_structure_mismatch: 期望系列数=" + expectedSeries
                + " 点数=" + expectedPoints
                + "，XML 系列数=" + actualSeries
                + " 点数=" + actualPoints;
        }

        public static ChartPartXml.SeriesGrid GridFromChartInfo(ChartInfo info)
        {
            var grid = new ChartPartXml.SeriesGrid
            {
                Type = info != null ? info.Type : "",
                Title = info != null ? info.Title : "",
                Legend = info != null ? info.LegendPosition : null,
                AxisYTitle = info != null && info.AxisY != null ? info.AxisY.Title : null,
                AxisY2Title = info != null && info.AxisY2 != null ? info.AxisY2.Title : null,
                AxisXTitle = info != null && info.AxisX != null ? info.AxisX.Title : null,
                AxisX = info != null ? info.AxisX : null,
                Gridlines = info != null ? info.Gridlines : null,
                ChartStyle = info != null ? info.ChartStyle : 0,
                AxisY = info != null ? info.AxisY : null,
                AxisY2 = info != null ? info.AxisY2 : null,
                GapWidth = info != null && info.GapWidth > 0 ? (int?)info.GapWidth : null,
                Overlap = info != null && (info.HasOverlap || info.Overlap != 0) ? (int?)info.Overlap : null,
                PlotColor = info != null ? info.PlotColor : null,
                ChartAreaColor = info != null ? info.ChartAreaColor : null,
                Series = new List<ChartPartXml.SeriesCol>()
            };
            if (info != null && (info.YAxisMin != 0 || info.YAxisMax != 0))
            {
                grid.YMin = info.YAxisMin;
                grid.YMax = info.YAxisMax;
            }

            if (info != null && info.YAxisMajorUnit > 0)
            {
                grid.YMajor = info.YAxisMajorUnit;
            }
            if (info == null || info.SeriesCollection == null)
            {
                return grid;
            }

            foreach (SeriesInfo s in info.SeriesCollection)
            {
                var col = new ChartPartXml.SeriesCol
                {
                    Name = s.Name,
                    AxisY = s.AxisY,
                    Color = s.Color,
                    Line = s.Line,
                    LineWeight = s.LineWeight,
                    Marker = s.Marker,
                    ShowDataLabels = s.ShowDataLabels,
                    LabelPosition = s.LabelPosition,
                    ChartType = s.ChartType,
                    Categories = new List<string>(),
                    Values = new List<string>()
                };
                if (s.Points != null)
                {
                    foreach (PointInfo p in s.Points)
                    {
                        col.Categories.Add(p.X ?? "");
                        col.Values.Add(p.Y.ToString(System.Globalization.CultureInfo.InvariantCulture));
                    }
                }

                grid.Series.Add(col);
            }

            return grid;
        }

        private static List<SeriesInfo> ToSeriesInfos(ChartPartXml.SeriesGrid grid)
        {
            var list = new List<SeriesInfo>();
            if (grid == null || grid.Series == null)
            {
                return list;
            }

            foreach (ChartPartXml.SeriesCol col in grid.Series)
            {
                var s = new SeriesInfo
                {
                    Name = col.Name,
                    AxisY = col.AxisY,
                    Color = string.IsNullOrEmpty(col.Color) ? "" : "#" + col.Color.TrimStart('#'),
                    ChartType = col.ChartType,
                    Line = string.IsNullOrEmpty(col.Line) ? null : "#" + col.Line.TrimStart('#'),
                    LineWeight = col.LineWeight,
                    Marker = col.Marker,
                    ShowDataLabels = col.ShowDataLabels,
                    LabelPosition = col.LabelPosition,
                    Points = new List<PointInfo>()
                };
                int n = col.Values != null ? col.Values.Count : 0;
                for (int i = 0; i < n; i++)
                {
                    double y;
                    double.TryParse(
                        col.Values[i],
                        System.Globalization.NumberStyles.Any,
                        System.Globalization.CultureInfo.InvariantCulture,
                        out y);
                    s.Points.Add(new PointInfo
                    {
                        X = col.Categories != null && i < col.Categories.Count ? col.Categories[i] : "",
                        Y = y
                    });
                }

                list.Add(s);
            }

            return list;
        }
    }
}
