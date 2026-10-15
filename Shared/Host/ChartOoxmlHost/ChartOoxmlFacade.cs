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

                if (!WordChartPackageHost.TryCopyChartOnDoc(copy, ordinal, out error))
                {
                    return false;
                }

                if (!WordChartPackageHost.TryDeleteChartOnDoc(destDoc, ordinal, out error))
                {
                    return false;
                }

                if (!WordChartPackageHost.TryPasteToDoc(destDoc, out error))
                {
                    return false;
                }

                if (pinColors)
                {
                    if (!TryReadWordChartParts(destDoc, out List<ChartPartRef> after, out error))
                    {
                        return false;
                    }

                    if (ordinal >= after.Count)
                    {
                        error = "贴回后对不上图";
                        return false;
                    }

                    List<string> afterColors = ChartPartXml.SnapshotSeriesColors(after[ordinal].ChartXml);
                    if (!ChartPartXml.ColorsMatch(colorSnap, afterColors))
                    {
                        if (!ChartPartXml.TryPinSeriesColors(after[ordinal].ChartXml, colorSnap))
                        {
                            error = "贴回后系列色被换且钉不住";
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
                            || !WordChartPackageHost.TryDeleteChartOnDoc(destDoc, ordinal, out error)
                            || !WordChartPackageHost.TryPasteToDoc(destDoc, out error))
                        {
                            return false;
                        }

                        if (!TryReadWordChartParts(destDoc, out List<ChartPartRef> pinned, out error)
                            || ordinal >= pinned.Count
                            || !ChartPartXml.ColorsMatch(
                                colorSnap,
                                ChartPartXml.SnapshotSeriesColors(pinned[ordinal].ChartXml)))
                        {
                            error = "贴回后系列色被换且钉不住";
                            return false;
                        }
                    }
                }

                if (!string.IsNullOrEmpty(stamp))
                {
                    if (!WordChartPackageHost.TryFindChartShape(destDoc, ordinal, out object next, out error)
                        || !WordChartPackageHost.TryWriteShapeTitle(next, stamp))
                    {
                        error = "贴回后写回 Title 戳失败";
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

        public static ChartPartXml.SeriesGrid GridFromChartInfo(ChartInfo info)
        {
            var grid = new ChartPartXml.SeriesGrid
            {
                Type = info != null ? info.Type : "",
                Title = info != null ? info.Title : "",
                Series = new List<ChartPartXml.SeriesCol>()
            };
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
