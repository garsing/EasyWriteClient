using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using WordAddIn1.OpenFiles;

namespace WordAddIn1.ChartOoxmlHost
{
    internal static partial class SpreadsheetChartPackageHost
    {
        private static readonly XNamespace C = "http://schemas.openxmlformats.org/drawingml/2006/chart";

        private static bool TryRewriteChrome(
            string xlsx,
            string sheetName,
            int index,
            ChartInfo format,
            string pinFormulasToFile,
            IList<string> pinColors,
            out string snapshotXml,
            out string error)
        {
            snapshotXml = null;
            error = null;
            if (!ChartPartIndex.TryReadXlsxChart(xlsx, sheetName, index, out string part, out XDocument doc, out error))
            {
                return false;
            }

            ChartPartXml.SeriesGrid grid = ChartOoxmlFacade.GridFromChartInfo(format);
            if (!ChartPartXml.TryApplyChrome(doc, grid, out error))
            {
                return false;
            }

            if (!string.IsNullOrEmpty(pinFormulasToFile))
            {
                PinFormulas(doc, pinFormulasToFile);
            }

            if (pinColors != null)
            {
                ChartPartXml.TryPinSeriesColors(doc, pinColors);
            }

            snapshotXml = doc.ToString(SaveOptions.DisableFormatting);
            byte[] bytes;
            using (var ms = new MemoryStream())
            {
                doc.Save(ms);
                bytes = ms.ToArray();
            }

            var replacements = new Dictionary<string, byte[]>(StringComparer.OrdinalIgnoreCase)
            {
                { part.TrimStart('/'), bytes }
            };
            return ChartPartIndex.TryApplyReplacementsToZip(xlsx, replacements, out error);
        }

        private static bool TryReadFormat(string xlsx, string sheetName, int index, out ChartInfo format, out string error)
        {
            format = null;
            if (!ChartPartIndex.TryReadXlsxChart(xlsx, sheetName, index, out _, out XDocument doc, out error))
            {
                return false;
            }

            if (!ChartPartXml.TryParseGrid(doc, out ChartPartXml.SeriesGrid grid, out error) || grid == null)
            {
                error = error ?? "读 chart 网格失败";
                return false;
            }

            format = GridToChartInfo(doc, grid);
            return true;
        }

        private static ChartInfo GridToChartInfo(XDocument doc, ChartPartXml.SeriesGrid grid)
        {
            bool hasTitle = doc.Descendants(C + "title").Any();
            bool deleted = doc.Descendants(C + "autoTitleDeleted")
                .Any(e => string.Equals(e.Attribute("val")?.Value, "1", StringComparison.Ordinal));
            var info = new ChartInfo
            {
                Type = grid.Type ?? "",
                Title = grid.Title ?? "",
                HasTitleText = hasTitle || deleted,
                ChartStyle = grid.ChartStyle,
                LegendPosition = grid.Legend,
                LegendFontColor = string.IsNullOrEmpty(grid.LegendFontColor)
                    ? grid.LegendFontColor
                    : "#" + grid.LegendFontColor.TrimStart('#'),
                Gridlines = grid.Gridlines,
                ChartAreaColor = string.IsNullOrEmpty(grid.ChartAreaColor) || string.Equals(grid.ChartAreaColor, "none", StringComparison.OrdinalIgnoreCase)
                    ? grid.ChartAreaColor
                    : "#" + grid.ChartAreaColor.TrimStart('#'),
                PlotColor = string.IsNullOrEmpty(grid.PlotColor) || string.Equals(grid.PlotColor, "none", StringComparison.OrdinalIgnoreCase)
                    ? grid.PlotColor
                    : "#" + grid.PlotColor.TrimStart('#'),
                TitleFontColor = string.IsNullOrEmpty(grid.TitleFontColor)
                    ? grid.TitleFontColor
                    : "#" + grid.TitleFontColor.TrimStart('#'),
                TitleFontSize = grid.TitleFontSize,
                TitleFontBold = grid.TitleFontBold,
                HasTitleFontBold = grid.HasTitleFontBold || grid.TitleFontBold,
                AxisX = grid.AxisX,
                AxisY = grid.AxisY,
                AxisY2 = grid.AxisY2,
                SeriesCollection = new List<SeriesInfo>()
            };
            if (deleted && !hasTitle)
            {
                info.Title = "";
            }

            if (grid.Overlap.HasValue)
            {
                info.Overlap = grid.Overlap.Value;
                info.HasOverlap = true;
            }

            if (grid.GapWidth.HasValue)
            {
                info.GapWidth = grid.GapWidth.Value;
            }

            if (grid.Explosion.HasValue)
            {
                info.Explosion = grid.Explosion.Value;
                info.HasExplosion = true;
            }
            else if (IsPieType(grid.Type))
            {
                info.Explosion = 0;
                info.HasExplosion = true;
            }

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

            if (grid.Series == null)
            {
                return info;
            }

            foreach (ChartPartXml.SeriesCol col in grid.Series)
            {
                info.SeriesCollection.Add(new SeriesInfo
                {
                    Name = col.Name,
                    AxisY = col.AxisY,
                    Color = string.Equals(col.Color, "none", StringComparison.OrdinalIgnoreCase)
                        ? "none"
                        : (string.IsNullOrEmpty(col.Color) ? col.Color : "#" + col.Color.TrimStart('#')),
                    ChartType = col.ChartType,
                    Line = col.Line,
                    LineWeight = col.LineWeight,
                    Marker = col.Marker,
                    MarkerSize = col.MarkerSize,
                    MarkerColor = col.MarkerColor,
                    MarkerFill = col.MarkerFill,
                    ShowDataLabels = col.ShowDataLabels,
                    ShowValue = col.ShowValue,
                    ShowPercentage = col.ShowPercentage,
                    LabelPosition = col.LabelPosition,
                    LabelFont = col.LabelFont,
                    LabelSize = col.LabelSize,
                    LabelColor = col.LabelColor,
                    LabelFormat = col.LabelFormat,
                    FillGradient = col.FillGradient,
                    FillAngle = col.FillAngle
                });
            }

            return info;
        }

        private static void PinFormulas(XDocument doc, string fileName)
        {
            if (doc == null || string.IsNullOrEmpty(fileName))
            {
                return;
            }

            foreach (XElement f in doc.Descendants(C + "f").ToList())
            {
                string local = StripBook(f.Value ?? "");
                int bang = local.IndexOf('!');
                if (bang <= 0)
                {
                    f.Value = local;
                    continue;
                }

                string sheet = local.Substring(0, bang).Trim().Trim('\'');
                string rest = local.Substring(bang);
                f.Value = "'[" + fileName + "]" + sheet + "'" + rest;
            }
        }

        private static string StripBook(string formula)
        {
            if (string.IsNullOrEmpty(formula))
            {
                return "";
            }

            return Regex.Replace(formula, @"\[[^\]]+\]", "");
        }

        private static bool TryVerify(
            Excel.Workbook book,
            string sheetName,
            int index,
            string snapshot,
            string forbiddenTemp,
            string allowedBookFile,
            out string error)
        {
            error = null;
            string temp = TempPath();
            try
            {
                book.SaveCopyAs(temp);
                if (!ChartPartIndex.TryReadXlsxChart(temp, sheetName, index, out _, out XDocument doc, out error))
                {
                    return false;
                }

                return FormulaAndSkinOk(doc, snapshot, forbiddenTemp, allowedBookFile, out error);
            }
            catch (Exception ex)
            {
                error = ex.Message;
                return false;
            }
            finally
            {
                TryDelete(temp);
            }
        }

        private static bool TryVerifyEt(
            object book,
            string sheetName,
            int index,
            string snapshot,
            string forbiddenTemp,
            string allowedBookFile,
            out string error)
        {
            string temp = TempPath();
            try
            {
                string full = EtCom.TryReadFullName(book);
                if (!SaveCopyEt(book, full, temp, out error))
                {
                    return false;
                }

                if (!ChartPartIndex.TryReadXlsxChart(temp, sheetName, index, out _, out XDocument doc, out error))
                {
                    return false;
                }

                return FormulaAndSkinOk(doc, snapshot, forbiddenTemp, allowedBookFile, out error);
            }
            finally
            {
                TryDelete(temp);
            }
        }

        private static bool FormulaAndSkinOk(
            XDocument live,
            string snapshot,
            string forbiddenTemp,
            string allowedBookFile,
            out string error)
        {
            error = null;
            string tempName = string.IsNullOrEmpty(forbiddenTemp) ? "" : Path.GetFileName(forbiddenTemp);
            foreach (XElement f in live.Descendants(C + "f"))
            {
                string text = f.Value ?? "";
                int bracket = text.IndexOf('[');
                if (bracket >= 0)
                {
                    int end = text.IndexOf(']', bracket + 1);
                    string book = end > bracket ? text.Substring(bracket + 1, end - bracket - 1) : "";
                    if (string.IsNullOrEmpty(allowedBookFile)
                        || book.IndexOf(allowedBookFile, StringComparison.OrdinalIgnoreCase) < 0)
                    {
                        error = "公式仍指向外部簿";
                        return false;
                    }
                }

                if (!string.IsNullOrEmpty(tempName)
                    && text.IndexOf(tempName, StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    error = "公式含临时文件名";
                    return false;
                }
            }

            if (!string.IsNullOrEmpty(snapshot))
            {
                XDocument want = XDocument.Parse(snapshot);
                string wantTitle = TitleText(want);
                string haveTitle = TitleText(live);
                if (!string.Equals(wantTitle, haveTitle, StringComparison.Ordinal))
                {
                    error = "标题被冲掉";
                    return false;
                }

                if (!ChartPartXml.ColorsMatch(
                        ChartPartXml.SnapshotSeriesColors(want),
                        ChartPartXml.SnapshotSeriesColors(live)))
                {
                    error = "系列色没有钉住";
                    return false;
                }
            }

            return true;
        }

        private static bool IsPieType(string type)
        {
            if (string.IsNullOrEmpty(type))
            {
                return false;
            }

            return type.IndexOf("pie", StringComparison.OrdinalIgnoreCase) >= 0
                || type.IndexOf("doughnut", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private static string TitleText(XDocument doc)
        {
            XElement title = doc.Descendants(C + "title").FirstOrDefault();
            if (title == null)
            {
                return "";
            }

            return string.Concat(title.Descendants().Where(e => e.Name.LocalName == "t").Select(e => e.Value));
        }
    }
}
