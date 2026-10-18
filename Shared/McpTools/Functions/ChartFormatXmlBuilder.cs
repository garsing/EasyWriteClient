using System;
using System.Collections.Generic;
using System.IO;
using System.Xml.Linq;

namespace WordAddIn1
{
    /// <summary>
    /// 将 ChartConfig 序列化为 ChartConfig XML（按 scope 裁剪）。
    /// </summary>
    public static class ChartFormatXmlBuilder
    {
        public static string Build(ChartConfig config, string scope)
        {
            string normalizedScope = string.IsNullOrEmpty(scope) ? "all" : scope.Trim().ToLowerInvariant();
            bool includeFormat = normalizedScope == "all" || normalizedScope == "format";
            bool includeData = normalizedScope == "all" || normalizedScope == "data";

            var root = new XElement("ChartConfig");
            var chartEl = new XElement("Chart");

            if (config?.Chart != null)
            {
                if (includeFormat)
                {
                    AppendFormatElements(chartEl, config.Chart);
                }

                if (includeData)
                {
                    AppendDataTable(chartEl, config.Chart);
                }
            }

            root.Add(chartEl);
            return new XDocument(new XDeclaration("1.0", "utf-8", "yes"), root).ToString();
        }

        private static void AppendFormatElements(XElement chartEl, ChartInfo chart)
        {
            AppendElementIfNotEmpty(chartEl, "Type", chart.Type);
            AppendElementIfNotEmpty(chartEl, "Title", chart.Title);
            AppendElementIfNotEmpty(chartEl, "Theme", chart.Theme);
            if (chart.Width > 0)
            {
                chartEl.Add(new XElement("Width", chart.Width));
            }

            if (chart.Height > 0)
            {
                chartEl.Add(new XElement("Height", chart.Height));
            }

            chartEl.Add(new XElement("ShowDataLabels", chart.ShowDataLabels));
            AppendElementIfNotEmpty(chartEl, "LegendPosition", chart.LegendPosition);
            AppendElementIfNotEmpty(chartEl, "PlotColor", chart.PlotColor);
            if (chart.TitleFontSize > 0)
            {
                chartEl.Add(new XElement("TitleFontSize", chart.TitleFontSize));
            }

            chartEl.Add(new XElement("TitleFontBold", chart.TitleFontBold));
            AppendElementIfNotEmpty(chartEl, "TitleFontColor", chart.TitleFontColor);
            AppendElementIfNotEmpty(chartEl, "Gridlines", chart.Gridlines);
            if (chart.GapWidth > 0)
            {
                chartEl.Add(new XElement("GapWidth", chart.GapWidth));
            }

            if (chart.HasOverlap || chart.Overlap != 0)
            {
                chartEl.Add(new XElement("Overlap", chart.Overlap));
            }

            AppendElementIfNotEmpty(chartEl, "PlotBox", chart.PlotBox);
            AppendElementIfNotEmpty(chartEl, "PlotInside", chart.PlotInside);
            AppendElementIfNotEmpty(chartEl, "ChartAreaColor", chart.ChartAreaColor);
            if (chart.ChartStyle > 0)
            {
                chartEl.Add(new XElement("ChartStyle", chart.ChartStyle));
            }

            AppendElementIfNotEmpty(chartEl, "LegendFontColor", chart.LegendFontColor);

            chartEl.Add(new XElement("DataMarkers", chart.DataMarkers));
            if (chart.MarkerSize > 0)
            {
                chartEl.Add(new XElement("MarkerSize", chart.MarkerSize));
            }

            if (chart.LineWeight > 0)
            {
                chartEl.Add(new XElement("LineWeight", chart.LineWeight));
            }

            if (chart.YAxisMin != 0 || chart.YAxisMax != 0)
            {
                chartEl.Add(new XElement("YAxisMin", chart.YAxisMin));
                chartEl.Add(new XElement("YAxisMax", chart.YAxisMax));
            }

            if (chart.YAxisMajorUnit > 0)
            {
                chartEl.Add(new XElement("YAxisMajorUnit", chart.YAxisMajorUnit));
            }

            if (chart.ShowValue.HasValue)
            {
                chartEl.Add(new XElement("ShowValue", chart.ShowValue.Value ? "true" : "false"));
            }

            if (chart.ShowPercentage.HasValue)
            {
                chartEl.Add(new XElement("ShowPercentage", chart.ShowPercentage.Value ? "true" : "false"));
            }
            if (chart.Explosion > 0)
            {
                chartEl.Add(new XElement("Explosion", chart.Explosion));
            }

            chartEl.Add(new XElement("FillMissingValues", chart.FillMissingValues));

            if (chart.AxisX != null)
            {
                var axisX = new XElement("AxisX");
                AppendElementIfNotEmpty(axisX, "Title", chart.AxisX.Title);
                AppendElementIfNotEmpty(axisX, "Type", chart.AxisX.Type);
                AppendElementIfNotEmpty(axisX, "Format", chart.AxisX.Format);
                AppendElementIfNotEmpty(axisX, "Visible", chart.AxisX.Visible);
                AppendElementIfNotEmpty(axisX, "TickFont", chart.AxisX.TickFont);
                AppendElementIfNotEmpty(axisX, "TickColor", chart.AxisX.TickColor);
                AppendElementIfNotEmpty(axisX, "TickSize", chart.AxisX.TickSize);
                AppendElementIfNotEmpty(axisX, "TickPosition", chart.AxisX.TickPosition);
                AppendElementIfNotEmpty(axisX, "MajorTick", chart.AxisX.MajorTick);
                AppendElementIfNotEmpty(axisX, "MinorTick", chart.AxisX.MinorTick);
                AppendElementIfNotEmpty(axisX, "Grid", chart.AxisX.Grid);
                AppendElementIfNotEmpty(axisX, "Line", chart.AxisX.Line);
                AppendElementIfNotEmpty(axisX, "LineWeight", chart.AxisX.LineWeight);
                AppendElementIfNotEmpty(axisX, "GridColor", chart.AxisX.GridColor);
                if (chart.AxisX.TickLabelCount.HasValue)
                {
                    axisX.Add(new XElement("TickLabelCount", chart.AxisX.TickLabelCount.Value));
                }

                if (chart.AxisX.TickLabelSpacing.HasValue)
                {
                    axisX.Add(new XElement("TickLabelSpacing", chart.AxisX.TickLabelSpacing.Value));
                }

                if (chart.AxisX.BetweenCategories.HasValue)
                {
                    axisX.SetAttributeValue(
                        "BetweenCategories",
                        chart.AxisX.BetweenCategories.Value ? "true" : "false");
                }

                chartEl.Add(axisX);
            }

            if (chart.AxisY != null)
            {
                var axisY = new XElement("AxisY");
                AppendElementIfNotEmpty(axisY, "Title", chart.AxisY.Title);
                AppendElementIfNotEmpty(axisY, "Visible", chart.AxisY.Visible);
                AppendElementIfNotEmpty(axisY, "TickFont", chart.AxisY.TickFont);
                AppendElementIfNotEmpty(axisY, "TickColor", chart.AxisY.TickColor);
                AppendElementIfNotEmpty(axisY, "TickSize", chart.AxisY.TickSize);
                AppendElementIfNotEmpty(axisY, "TickPosition", chart.AxisY.TickPosition);
                AppendElementIfNotEmpty(axisY, "MajorTick", chart.AxisY.MajorTick);
                AppendElementIfNotEmpty(axisY, "MinorTick", chart.AxisY.MinorTick);
                AppendElementIfNotEmpty(axisY, "Grid", chart.AxisY.Grid);
                AppendElementIfNotEmpty(axisY, "Format", chart.AxisY.Format);
                AppendElementIfNotEmpty(axisY, "Line", chart.AxisY.Line);
                AppendElementIfNotEmpty(axisY, "LineWeight", chart.AxisY.LineWeight);
                AppendElementIfNotEmpty(axisY, "GridColor", chart.AxisY.GridColor);
                chartEl.Add(axisY);
            }

            if (chart.AxisY2 != null)
            {
                var axisY2 = new XElement("AxisY2");
                AppendElementIfNotEmpty(axisY2, "Title", chart.AxisY2.Title);
                AppendElementIfNotEmpty(axisY2, "Visible", chart.AxisY2.Visible);
                AppendElementIfNotEmpty(axisY2, "TickFont", chart.AxisY2.TickFont);
                AppendElementIfNotEmpty(axisY2, "TickColor", chart.AxisY2.TickColor);
                AppendElementIfNotEmpty(axisY2, "TickSize", chart.AxisY2.TickSize);
                AppendElementIfNotEmpty(axisY2, "TickPosition", chart.AxisY2.TickPosition);
                AppendElementIfNotEmpty(axisY2, "MajorTick", chart.AxisY2.MajorTick);
                AppendElementIfNotEmpty(axisY2, "MinorTick", chart.AxisY2.MinorTick);
                AppendElementIfNotEmpty(axisY2, "Grid", chart.AxisY2.Grid);
                AppendElementIfNotEmpty(axisY2, "Format", chart.AxisY2.Format);
                AppendElementIfNotEmpty(axisY2, "Line", chart.AxisY2.Line);
                AppendElementIfNotEmpty(axisY2, "LineWeight", chart.AxisY2.LineWeight);
                AppendElementIfNotEmpty(axisY2, "GridColor", chart.AxisY2.GridColor);
                chartEl.Add(axisY2);
            }
        }

        private static void AppendDataTable(XElement chartEl, ChartInfo chart)
        {
            var dataTable = new XElement("DataTable");
            var header = new XElement("Header");
            header.Add(new XElement("Column", new XAttribute("type", "category"), "类别"));

            var seriesList = chart.SeriesCollection ?? new List<SeriesInfo>();
            foreach (var series in seriesList)
            {
                var col = new XElement(
                    "Column",
                    new XAttribute("type", "value"),
                    new XAttribute("axisY", string.IsNullOrEmpty(series.AxisY) ? ChartConfigSyntax.AxisYPrimary : series.AxisY));

                if (!string.IsNullOrEmpty(series.Color))
                {
                    col.Add(new XAttribute("color", series.Color));
                }

                if (!string.IsNullOrEmpty(series.ChartType))
                {
                    col.Add(new XAttribute("chartType", series.ChartType));
                }

                if (series.ShowDataLabels.HasValue)
                {
                    col.Add(new XAttribute("showDataLabels", series.ShowDataLabels.Value ? "true" : "false"));
                }

                ChartSeriesChrome.WriteColumnAttrs(col, series);

                col.Value = series.Name ?? "";
                header.Add(col);
            }

            dataTable.Add(header);

            var rowsEl = new XElement("Rows");
            if (seriesList.Count > 0 && seriesList[0].Points != null && seriesList[0].Points.Count > 0)
            {
                int rowCount = seriesList[0].Points.Count;
                for (int rowIndex = 0; rowIndex < rowCount; rowIndex++)
                {
                    var rowValues = new List<string>();
                    rowValues.Add(seriesList[0].Points[rowIndex].X ?? "");

                    foreach (var series in seriesList)
                    {
                        if (rowIndex < series.Points.Count)
                        {
                            rowValues.Add(series.Points[rowIndex].Y.ToString("F2"));
                        }
                        else
                        {
                            rowValues.Add("0");
                        }
                    }

                    rowsEl.Add(new XElement("Row", EscapeXmlText(string.Join(",", rowValues))));
                }
            }

            dataTable.Add(rowsEl);
            chartEl.Add(dataTable);
        }

        private static void AppendElementIfNotEmpty(XElement parent, string name, string value)
        {
            if (!string.IsNullOrEmpty(value))
            {
                parent.Add(new XElement(name, value));
            }
        }

        private static string EscapeXmlText(string value)
        {
            if (string.IsNullOrEmpty(value))
            {
                return "";
            }

            return value
                .Replace("&", "&amp;")
                .Replace("<", "&lt;")
                .Replace(">", "&gt;");
        }

        public static string GenerateFilename(string scope, string chartId)
        {
            string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss_fff");
            string safeId = SanitizeChartIdForFilename(chartId);
            return $"chart_{scope}_{safeId}_{timestamp}.xml";
        }

        /// <summary>
        /// chart_id 来自 Base64 子串，可能含 / + 等，须净化后才能作为文件名（勿改变 chart_id 本身）。
        /// </summary>
        public static string SanitizeChartIdForFilename(string chartId)
        {
            if (string.IsNullOrEmpty(chartId))
            {
                return "unknown";
            }

            string safe = chartId;
            foreach (char c in Path.GetInvalidFileNameChars())
            {
                safe = safe.Replace(c, '_');
            }

            // Base64 的 + / 在部分环境下可能仍出现在 id 中
            safe = safe.Replace('/', '_').Replace('+', '_');
            return safe;
        }
    }
}
