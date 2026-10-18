using System;
using System.Collections.Generic;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    /// <summary>
    /// 从 Word.Chart COM 提取 ChartConfig（对齐 F_CreateChartFromXmlTool schema）。
    /// </summary>
    public static class ChartExtractHelper
    {
        public static ChartConfig Extract(
            Word.Chart chart,
            Word.InlineShape inlineShape,
            string scope,
            Word.Shape floatingShape = null)
        {
            string normalizedScope = string.IsNullOrEmpty(scope) ? "all" : scope.Trim().ToLowerInvariant();
            var config = new ChartConfig
            {
                Chart = new ChartInfo()
            };

            if (normalizedScope == "data" || normalizedScope == "all")
            {
                ExtractData(chart, config.Chart);
            }

            if (normalizedScope == "format" || normalizedScope == "all")
            {
                ExtractFormat(chart, inlineShape, config.Chart, floatingShape);
            }

            return config;
        }

        private static void ExtractFormat(
            Word.Chart chart,
            Word.InlineShape inlineShape,
            ChartInfo chartInfo,
            Word.Shape floatingShape)
        {
            if (chart == null || chartInfo == null)
            {
                return;
            }

            try
            {
                chartInfo.Type = MapChartType(chart);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[ExtractChart] 读取 ChartType 失败: {ex.Message}");
            }

            try
            {
                chartInfo.Title = chart.HasTitle ? (chart.ChartTitle?.Text ?? "") : "";
                if (chart.HasTitle && chart.ChartTitle != null)
                {
                    try
                    {
                        chartInfo.TitleFontSize = (int)chart.ChartTitle.Format.TextFrame2.TextRange.Font.Size;
                    }
                    catch { }

                    try
                    {
                        chartInfo.TitleFontBold = chart.ChartTitle.Format.TextFrame2.TextRange.Font.Bold
                            == Microsoft.Office.Core.MsoTriState.msoTrue;
                    }
                    catch { }

                    try
                    {
                        int bgr = (int)chart.ChartTitle.Font.Color;
                        if (bgr != 0)
                        {
                            chartInfo.TitleFontColor = TableFormatExtractor.ConvertBgrToRgb(bgr.ToString("X6"));
                        }
                    }
                    catch { }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[ExtractChart] 读取标题失败: {ex.Message}");
            }

            try
            {
                int style = chart.ChartStyle;
                chartInfo.ChartStyle = style;
                chartInfo.Theme = style == 1 ? "office" : "office";
            }
            catch { }

            try
            {
                if (inlineShape != null)
                {
                    chartInfo.Width = (int)Math.Round(inlineShape.Width);
                    chartInfo.Height = (int)Math.Round(inlineShape.Height);
                }
                else if (floatingShape != null)
                {
                    chartInfo.Width = (int)Math.Round(floatingShape.Width);
                    chartInfo.Height = (int)Math.Round(floatingShape.Height);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[ExtractChart] 读取尺寸失败: {ex.Message}");
            }

            try
            {
                if (chart.HasLegend)
                {
                    chartInfo.LegendPosition = MapLegendPosition(chart.Legend.Position);
                    try
                    {
                        chartInfo.LegendFontColor = ChartSeriesChrome.RgbHex(Convert.ToInt32(chart.Legend.Font.Color));
                    }
                    catch
                    {
                    }
                }
                else
                {
                    chartInfo.LegendPosition = ChartConfigSyntax.LegendNone;
                }

                bool anyDataLabels = false;
                dynamic seriesCollection = chart.SeriesCollection();
                int count = seriesCollection.Count;
                for (int i = 1; i <= count; i++)
                {
                    try
                    {
                        if (seriesCollection(i).HasDataLabels)
                        {
                            anyDataLabels = true;
                            break;
                        }
                    }
                    catch { }
                }

                chartInfo.ShowDataLabels = anyDataLabels;
                if (string.Equals(chartInfo.Type, ChartConfigSyntax.TypePie3d, StringComparison.Ordinal)
                    || string.Equals(chartInfo.Type, ChartConfigSyntax.TypePie2d, StringComparison.Ordinal))
                {
                    TryExtractPieLabelFlags(seriesCollection, chartInfo);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[ExtractChart] 读取图例/数据标签失败: {ex.Message}");
            }

            try
            {
                chartInfo.Gridlines = MapGridlines(chart);
            }
            catch { }

            try
            {
                if (ChartConfigSyntax.IsColumn(chartInfo.Type) || ChartConfigSyntax.IsLineOrColumn(chartInfo.Type))
                {
                    dynamic group = chart.ChartGroups(1);
                    if (group != null)
                    {
                        chartInfo.GapWidth = Convert.ToInt32(group.GapWidth);
                        chartInfo.Overlap = Convert.ToInt32(group.Overlap);
                        chartInfo.HasOverlap = true;
                    }
                }
            }
            catch { }

            TryReadAreaColor(chart.PlotArea, value => chartInfo.PlotColor = value);
            TryReadAreaColor(chart.ChartArea, value => chartInfo.ChartAreaColor = value);

            try
            {
                Word.Axis yAxis = chart.Axes(Word.XlAxisType.xlValue);
                chartInfo.YAxisMin = yAxis.MinimumScale;
                chartInfo.YAxisMax = yAxis.MaximumScale;
                chartInfo.YAxisMajorUnit = yAxis.MajorUnit;
            }
            catch { }

            try
            {
                chartInfo.AxisX = new AxisXInfo
                {
                    Title = TryGetAxisTitle(chart, Word.XlAxisType.xlCategory),
                    Type = "category"
                };
                try
                {
                    Word.Axis xAxis = chart.Axes(Word.XlAxisType.xlCategory);
                    ChartSeriesChrome.CaptureAxis(xAxis, (visible, font, color, size, format) =>
                    {
                        chartInfo.AxisX.Visible = visible;
                        chartInfo.AxisX.TickFont = font;
                        chartInfo.AxisX.TickColor = color;
                        chartInfo.AxisX.TickSize = size;
                        chartInfo.AxisX.Format = format;
                    });
                    ChartSeriesChrome.CaptureTicks(xAxis, (pos, major, minor, grid) =>
                    {
                        chartInfo.AxisX.TickPosition = pos;
                        chartInfo.AxisX.MajorTick = major;
                        chartInfo.AxisX.MinorTick = minor;
                        chartInfo.AxisX.Grid = grid;
                    });
                    ChartSeriesChrome.CaptureAxisStroke(xAxis, (line, weight, gridColor) =>
                    {
                        chartInfo.AxisX.Line = line;
                        chartInfo.AxisX.LineWeight = weight;
                        chartInfo.AxisX.GridColor = gridColor;
                    });
                }
                catch { }

                if (ChartConfigSyntax.IsLineOrColumn(chartInfo.Type))
                {
                    try
                    {
                        Word.Axis xAxis = chart.Axes(Word.XlAxisType.xlCategory);
                        chartInfo.AxisX.BetweenCategories = xAxis.AxisBetweenCategories;
                    }
                    catch { }
                }

                chartInfo.AxisY = new AxisYInfo
                {
                    Title = TryGetAxisTitle(chart, Word.XlAxisType.xlValue)
                };
                try
                {
                    Word.Axis yAxisForChrome = chart.Axes(Word.XlAxisType.xlValue);
                    ChartSeriesChrome.CaptureAxis(yAxisForChrome, (visible, font, color, size, format) =>
                    {
                        chartInfo.AxisY.Visible = visible;
                        chartInfo.AxisY.TickFont = font;
                        chartInfo.AxisY.TickColor = color;
                        chartInfo.AxisY.TickSize = size;
                        chartInfo.AxisY.Format = format;
                    });
                    ChartSeriesChrome.CaptureTicks(yAxisForChrome, (pos, major, minor, grid) =>
                    {
                        chartInfo.AxisY.TickPosition = pos;
                        chartInfo.AxisY.MajorTick = major;
                        chartInfo.AxisY.MinorTick = minor;
                        chartInfo.AxisY.Grid = grid;
                    });
                    ChartSeriesChrome.CaptureAxisStroke(yAxisForChrome, (line, weight, gridColor) =>
                    {
                        chartInfo.AxisY.Line = line;
                        chartInfo.AxisY.LineWeight = weight;
                        chartInfo.AxisY.GridColor = gridColor;
                    });
                }
                catch { }

                if (string.Equals(chartInfo.Type, ChartConfigSyntax.TypePie2d, StringComparison.Ordinal)
                    || string.Equals(chartInfo.Type, ChartConfigSyntax.TypePie3d, StringComparison.Ordinal))
                {
                    try
                    {
                        dynamic pie = chart.SeriesCollection(1);
                        chartInfo.Explosion = Convert.ToInt32(pie.Explosion);
                    }
                    catch { }
                }
                try
                {
                    Word.Axis y2Axis = chart.Axes(Word.XlAxisType.xlValue, Word.XlAxisGroup.xlSecondary);
                    if (chartInfo.AxisY2 == null)
                    {
                        chartInfo.AxisY2 = new AxisYInfo();
                    }

                    string y2Title = TryGetAxisTitle(chart, Word.XlAxisType.xlValue, Word.XlAxisGroup.xlSecondary);
                    if (!string.IsNullOrEmpty(y2Title))
                    {
                        chartInfo.AxisY2.Title = y2Title;
                    }

                    ChartSeriesChrome.CaptureAxis(y2Axis, (visible, font, color, size, format) =>
                    {
                        chartInfo.AxisY2.Visible = visible;
                        chartInfo.AxisY2.TickFont = font;
                        chartInfo.AxisY2.TickColor = color;
                        chartInfo.AxisY2.TickSize = size;
                        chartInfo.AxisY2.Format = format;
                    });
                    ChartSeriesChrome.CaptureTicks(y2Axis, (pos, major, minor, grid) =>
                    {
                        chartInfo.AxisY2.TickPosition = pos;
                        chartInfo.AxisY2.MajorTick = major;
                        chartInfo.AxisY2.MinorTick = minor;
                        chartInfo.AxisY2.Grid = grid;
                    });
                    ChartSeriesChrome.CaptureAxisStroke(y2Axis, (line, weight, gridColor) =>
                    {
                        chartInfo.AxisY2.Line = line;
                        chartInfo.AxisY2.LineWeight = weight;
                        chartInfo.AxisY2.GridColor = gridColor;
                    });
                }
                catch { }
            }
            catch { }
        }

        private static void ExtractData(Word.Chart chart, ChartInfo chartInfo)
        {
            if (chart == null || chartInfo == null)
            {
                return;
            }

            chartInfo.SeriesCollection = new List<SeriesInfo>();

            try
            {
                dynamic seriesCollection = chart.SeriesCollection();
                int seriesCount = seriesCollection.Count;
                if (seriesCount == 0)
                {
                    return;
                }

                for (int i = 1; i <= seriesCount; i++)
                {
                    try
                    {
                        dynamic series = seriesCollection(i);
                        string chartType = MapChartType(chart);
                        var seriesInfo = new SeriesInfo
                        {
                            Name = series.Name ?? $"系列{i}",
                            AxisY = ChartConfigSyntax.AxisYPrimary,
                            Color = "",
                            ChartType = "",
                            ShowDataLabels = null,
                            Points = new List<PointInfo>()
                        };

                        try
                        {
                            if (IsSecondaryAxisGroup(series.AxisGroup))
                            {
                                seriesInfo.AxisY = ChartConfigSyntax.AxisYSecondary;
                            }
                        }
                        catch { }

                        seriesInfo.Color = ExtractSeriesColorHex(series, chartType);
                        ChartSeriesChrome.Capture(series, seriesInfo);

                        try
                        {
                            string seriesType = MapXlChartType((int)series.ChartType);
                            if (!string.IsNullOrEmpty(seriesType))
                            {
                                seriesInfo.ChartType = seriesType;
                            }
                        }
                        catch { }

                        try
                        {
                            seriesInfo.ShowDataLabels = series.HasDataLabels;
                        }
                        catch { }

                        object valuesObj = null;
                        object xValuesObj = null;
                        try
                        {
                            var seriesObj = (object)series;
                            var t = seriesObj.GetType();
                            valuesObj = t.InvokeMember(
                                "Values",
                                System.Reflection.BindingFlags.GetProperty
                                    | System.Reflection.BindingFlags.Public
                                    | System.Reflection.BindingFlags.Instance,
                                null,
                                seriesObj,
                                null);
                            xValuesObj = t.InvokeMember(
                                "XValues",
                                System.Reflection.BindingFlags.GetProperty
                                    | System.Reflection.BindingFlags.Public
                                    | System.Reflection.BindingFlags.Instance,
                                null,
                                seriesObj,
                                null);
                        }
                        catch (Exception ex)
                        {
                            System.Diagnostics.Debug.WriteLine(
                                $"[ExtractChart] 反射获取 Values/XValues 失败: {ex.Message}");
                        }

                        Array valuesArray = valuesObj as Array;
                        Array xValuesArray = xValuesObj as Array;
                        int pointCount = 0;
                        if (valuesArray != null)
                        {
                            pointCount = valuesArray.GetUpperBound(0) - valuesArray.GetLowerBound(0) + 1;
                        }

                        int valuesStartIndex = valuesArray?.GetLowerBound(0) ?? 0;
                        int xValuesStartIndex = xValuesArray?.GetLowerBound(0) ?? 0;

                        for (int j = 0; j < pointCount; j++)
                        {
                            string xValue = "";
                            double yValue = 0;

                            if (xValuesArray != null)
                            {
                                int xIndex = xValuesStartIndex + j;
                                if (xIndex <= xValuesArray.GetUpperBound(0))
                                {
                                    object xObj = xValuesArray.GetValue(xIndex);
                                    xValue = xObj?.ToString() ?? "";
                                }
                            }

                            if (valuesArray != null)
                            {
                                int yIndex = valuesStartIndex + j;
                                if (yIndex <= valuesArray.GetUpperBound(0))
                                {
                                    object yObj = valuesArray.GetValue(yIndex);
                                    if (yObj != null && double.TryParse(yObj.ToString(), out double parsed))
                                    {
                                        yValue = parsed;
                                    }
                                }
                            }

                            seriesInfo.Points.Add(new PointInfo
                            {
                                X = xValue,
                                Y = yValue
                            });
                        }

                        chartInfo.SeriesCollection.Add(seriesInfo);
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"[ExtractChart] 提取系列{i}失败: {ex.Message}");
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[ExtractChart] 提取数据失败: {ex.Message}");
            }
        }

        /// <summary>
        /// 系列 AxisGroup 从 Excel 回来常常是整数或另一套枚举，和 Word.XlAxisGroup 直接比较会失败，次轴就被记成 y。
        /// </summary>
        private static bool IsSecondaryAxisGroup(object axisGroup)
        {
            if (axisGroup == null)
            {
                return false;
            }

            try
            {
                return Convert.ToInt32(axisGroup) == (int)Word.XlAxisGroup.xlSecondary;
            }
            catch
            {
                return false;
            }
        }

        private static string MapChartType(Word.Chart chart)
        {
            int chartType = (int)chart.ChartType;
            string mapped = MapXlChartType(chartType);
            if (!string.IsNullOrEmpty(mapped))
            {
                return mapped;
            }

            if (chartType == -4111)
            {
                try
                {
                    dynamic first = chart.SeriesCollection(1);
                    string fromSeries = MapXlChartType((int)first.ChartType);
                    if (!string.IsNullOrEmpty(fromSeries))
                    {
                        return fromSeries;
                    }
                }
                catch
                {
                }

                return ChartConfigSyntax.TypeColumn;
            }

            string raw = chart.ChartType.ToString().ToLowerInvariant();
            if (raw.IndexOf("pie", StringComparison.Ordinal) >= 0
                && raw.IndexOf("3d", StringComparison.Ordinal) < 0)
            {
                return ChartConfigSyntax.TypePie2d;
            }

            return raw;
        }

        private static string MapXlChartType(int chartType)
        {
            if (chartType == (int)Excel.XlChartType.xlColumnClustered
                || chartType == (int)Excel.XlChartType.xlColumnStacked
                || chartType == (int)Excel.XlChartType.xlColumnStacked100)
            {
                return ChartConfigSyntax.TypeColumn;
            }

            if (chartType == (int)Excel.XlChartType.xlBarClustered
                || chartType == (int)Excel.XlChartType.xlBarStacked
                || chartType == (int)Excel.XlChartType.xlBarStacked100)
            {
                return ChartConfigSyntax.TypeBar;
            }

            if (chartType == (int)Excel.XlChartType.xlLine
                || chartType == (int)Excel.XlChartType.xlLineMarkers
                || chartType == (int)Excel.XlChartType.xlLineStacked
                || chartType == (int)Excel.XlChartType.xlLineMarkersStacked)
            {
                return ChartConfigSyntax.TypeLine;
            }

            if (chartType == (int)Excel.XlChartType.xlPie)
            {
                return ChartConfigSyntax.TypePie2d;
            }

            if (chartType == (int)Excel.XlChartType.xl3DPie)
            {
                return ChartConfigSyntax.TypePie3d;
            }

            return null;
        }

        private static void TryReadAreaColor(dynamic area, Action<string> assign)
        {
            try
            {
                if (area == null)
                {
                    return;
                }

                dynamic fill = area.Format.Fill;
                int visible = Convert.ToInt32(fill.Visible);
                if (visible == (int)Microsoft.Office.Core.MsoTriState.msoFalse)
                {
                    assign("none");
                    return;
                }

                try
                {
                    assign(ConvertRgbToHex(Convert.ToInt32(fill.ForeColor.RGB)));
                    return;
                }
                catch
                {
                }

                int interior = Convert.ToInt32(area.Interior.Color);
                assign(ConvertRgbToHex(interior));
            }
            catch
            {
            }
        }

        private static string MapLegendPosition(Word.XlLegendPosition position)
        {
            if (position == Word.XlLegendPosition.xlLegendPositionTop)
            {
                return "top";
            }

            if (position == Word.XlLegendPosition.xlLegendPositionBottom)
            {
                return "bottom";
            }

            if (position == Word.XlLegendPosition.xlLegendPositionLeft)
            {
                return "left";
            }

            if (position == Word.XlLegendPosition.xlLegendPositionRight)
            {
                return "right";
            }

            return "bottom";
        }

        private static string MapGridlines(Word.Chart chart)
        {
            try
            {
                Word.Axis valueAxis = chart.Axes(Word.XlAxisType.xlValue);
                return valueAxis.HasMajorGridlines ? ChartConfigSyntax.GridOn : ChartConfigSyntax.GridOff;
            }
            catch
            {
                return ChartConfigSyntax.GridOff;
            }
        }

        private static string TryGetAxisTitle(
            Word.Chart chart,
            Word.XlAxisType axisType,
            Word.XlAxisGroup group = Word.XlAxisGroup.xlPrimary)
        {
            try
            {
                Word.Axis axis = chart.Axes(axisType, group);
                if (axis.HasTitle)
                {
                    return axis.AxisTitle?.Text ?? "";
                }
            }
            catch { }

            return "";
        }

        private static void TryExtractPieLabelFlags(dynamic seriesCollection, ChartInfo chartInfo)
        {
            try
            {
                dynamic first = seriesCollection(1);
                if (first == null || !first.HasDataLabels)
                {
                    return;
                }

                dynamic labels = first.DataLabels;
                try
                {
                    chartInfo.ShowValue = (bool)labels.ShowValue;
                }
                catch
                {
                }

                try
                {
                    chartInfo.ShowPercentage = (bool)labels.ShowPercentage;
                }
                catch
                {
                }
            }
            catch
            {
            }
        }

        private static string ExtractSeriesColorHex(dynamic series, string chartType)
        {
            bool preferFill = !string.Equals(chartType, "line", StringComparison.OrdinalIgnoreCase);

            if (preferFill)
            {
                string fillColor = TryReadSeriesColorRgb(series, useFill: true);
                if (!string.IsNullOrEmpty(fillColor))
                {
                    return fillColor;
                }

                return TryReadSeriesColorRgb(series, useFill: false);
            }

            string lineColor = TryReadSeriesColorRgb(series, useFill: false);
            if (!string.IsNullOrEmpty(lineColor))
            {
                return lineColor;
            }

            return TryReadSeriesColorRgb(series, useFill: true);
        }

        private static string TryReadSeriesColorRgb(dynamic series, bool useFill)
        {
            try
            {
                int rgbValue = useFill
                    ? (int)series.Format.Fill.ForeColor.RGB
                    : (int)series.Format.Line.ForeColor.RGB;
                return ConvertRgbToHex(rgbValue);
            }
            catch
            {
                return "";
            }
        }

        private static string ConvertRgbToHex(int rgbValue)
        {
            int r = rgbValue & 0xFF;
            int g = (rgbValue >> 8) & 0xFF;
            int b = (rgbValue >> 16) & 0xFF;
            return string.Format("{0:X2}{1:X2}{2:X2}", r, g, b);
        }
    }
}
