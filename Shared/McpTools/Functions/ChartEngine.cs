using System;
using System.Collections.Generic;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    /// <summary>
    /// 图表 create/apply 共用 COM 引擎（非 MCP 工具层）。
    /// </summary>
    public static class ChartEngine
    {
        public static Excel.XlChartType GetChartType(string type)
        {
            switch (type.ToLower())
            {
                case "column_clustered":
                    return Excel.XlChartType.xlColumnClustered;
                case "line":
                    return Excel.XlChartType.xlLine;
                case "pie3d":
                    return Excel.XlChartType.xl3DPie;
                default:
                    return Excel.XlChartType.xlColumnClustered;
            }
        }

        /// <summary>
        /// 获取图表样式
        /// </summary>
        public static int GetChartStyle(string theme)
        {
            switch (theme.ToLower())
            {
                case "office":
                    return 1; // Office style
                default:
                    return 1;
            }
        }

        /// <summary>
        /// 获取图例位置
        /// </summary>
        public static int GetLegendPosition(string position)
        {
            switch (position.ToLower())
            {
                case "top":
                    return -4160; // xlLegendPositionTop
                case "bottom":
                    return -4107; // xlLegendPositionBottom
                case "left":
                    return -4131; // xlLegendPositionLeft
                case "right":
                    return -4152; // xlLegendPositionRight
                default:
                    return -4107; // xlLegendPositionBottom
            }
        }

        /// <summary>
        /// 解析类别轴刻度位置：XML 显式值优先，否则按图表类型默认（折线=false，柱形=true）。
        /// </summary>
        public static bool ResolveAxisBetweenCategories(ChartInfo chartInfo)
        {
            if (chartInfo?.AxisX?.BetweenCategories.HasValue == true)
            {
                return chartInfo.AxisX.BetweenCategories.Value;
            }

            return string.Equals(chartInfo?.Type, "line", StringComparison.OrdinalIgnoreCase) ? false : true;
        }

        private static void ApplyAxisTitles(Word.Chart chart, ChartInfo chartInfo, List<string> errors)
        {
            if (chart == null || chartInfo == null)
            {
                return;
            }

            if (chartInfo.AxisX != null && !string.IsNullOrWhiteSpace(chartInfo.AxisX.Title))
            {
                try
                {
                    Word.Axis xAxis = chart.Axes(Word.XlAxisType.xlCategory);
                    xAxis.HasTitle = true;
                    xAxis.AxisTitle.Text = chartInfo.AxisX.Title.Trim();
                    System.Diagnostics.Debug.WriteLine($"[DEBUG] 已设置X轴标题: {chartInfo.AxisX.Title}");
                }
                catch (Exception ex)
                {
                    string errorMsg = $"设置X轴标题失败: {ex.Message}";
                    errors.Add(errorMsg);
                    System.Diagnostics.Debug.WriteLine($"[DEBUG] {errorMsg}");
                }
            }

            if (chartInfo.AxisY != null && !string.IsNullOrWhiteSpace(chartInfo.AxisY.PrimaryTitle))
            {
                try
                {
                    Word.Axis primaryYAxis = chart.Axes(Word.XlAxisType.xlValue);
                    primaryYAxis.HasTitle = true;
                    primaryYAxis.AxisTitle.Text = chartInfo.AxisY.PrimaryTitle.Trim();
                    System.Diagnostics.Debug.WriteLine($"[DEBUG] 已设置Y轴标题: {chartInfo.AxisY.PrimaryTitle}");
                }
                catch (Exception ex)
                {
                    string errorMsg = $"设置Y轴标题失败: {ex.Message}";
                    errors.Add(errorMsg);
                    System.Diagnostics.Debug.WriteLine($"[DEBUG] {errorMsg}");
                }
            }

            if (chartInfo.AxisY != null && !string.IsNullOrWhiteSpace(chartInfo.AxisY.SecondaryTitle))
            {
                try
                {
                    Word.Axis secondaryYAxis = chart.Axes(
                        Word.XlAxisType.xlValue,
                        Word.XlAxisGroup.xlSecondary);
                    secondaryYAxis.HasTitle = true;
                    secondaryYAxis.AxisTitle.Text = chartInfo.AxisY.SecondaryTitle.Trim();
                    System.Diagnostics.Debug.WriteLine($"[DEBUG] 已设置次Y轴标题: {chartInfo.AxisY.SecondaryTitle}");
                }
                catch (Exception ex)
                {
                    string errorMsg = $"设置次Y轴标题失败: {ex.Message}";
                    errors.Add(errorMsg);
                    System.Diagnostics.Debug.WriteLine($"[DEBUG] {errorMsg}");
                }
            }
        }

        /// <summary>
        /// 填充图表数据（按Series填充，支持双Y轴和混合图表类型）
        /// </summary>
        public static ChartOperationResult FillChartData(Word.Chart chart, ChartInfo chartInfo)
        {
            var errors = new List<string>();

            try
            {
                if (chartInfo.SeriesCollection == null || chartInfo.SeriesCollection.Count == 0)
                {
                    return new ChartOperationResult { Success = false, Error = "SeriesCollection为空或没有Series" };
                }

                // 收集所有Series的X轴标签（从第一个Series提取，假设所有Series的X值一致）
                var xAxisLabels = new List<string>();
                if (chartInfo.SeriesCollection[0].Points != null && chartInfo.SeriesCollection[0].Points.Count > 0)
                {
                    foreach (var point in chartInfo.SeriesCollection[0].Points)
                    {
                        xAxisLabels.Add(point.X ?? "");
                    }
                }

                // 遍历每个Series，设置数据
                int validSeriesIndex = 0;
                for (int i = 0; i < chartInfo.SeriesCollection.Count; i++)
                {
                    var seriesInfo = chartInfo.SeriesCollection[i];
                    if (seriesInfo == null || seriesInfo.Points == null || seriesInfo.Points.Count == 0)
                    {
                        continue;
                    }

                    validSeriesIndex++;
                    try
                    {
                        // 获取图表系列（从1开始索引）
                        dynamic series = chart.SeriesCollection(validSeriesIndex);
                        if (series == null)
                        {
                            errors.Add($"无法访问Series{validSeriesIndex}");
                            continue;
                        }

                        // 收集Y值
                        var yValues = new List<double>();
                        var labelsToUse = new List<string>();
                        foreach (var point in seriesInfo.Points)
                        {
                            yValues.Add(point.Y);
                            // 使用对应的X轴标签
                            int pointIndex = seriesInfo.Points.IndexOf(point);
                            if (pointIndex < xAxisLabels.Count)
                            {
                                labelsToUse.Add(xAxisLabels[pointIndex]);
                            }
                            else
                            {
                                labelsToUse.Add(point.X ?? "");
                            }
                        }

                        // 设置系列值
                        if (yValues.Count > 0)
                        {
                            series.Values = yValues.ToArray();
                        }

                        // 设置X轴标签
                        if (labelsToUse.Count > 0)
                        {
                            series.XValues = labelsToUse.ToArray();
                        }

                        // 设置系列名称
                        if (!string.IsNullOrEmpty(seriesInfo.Name))
                        {
                            series.Name = seriesInfo.Name;
                        }

                        // 设置Y轴（primary或secondary）
                        if (!string.IsNullOrEmpty(seriesInfo.AxisY) && seriesInfo.AxisY == "secondary")
                        {
                            try
                            {
                                series.AxisGroup = Word.XlAxisGroup.xlSecondary;
                            }
                            catch (Exception ex)
                            {
                                errors.Add($"设置Series{i + 1}为secondary轴失败: {ex.Message}");
                            }
                        }

                        // 设置系列颜色（如果指定）
                        if (!string.IsNullOrEmpty(seriesInfo.Color))
                        {
                            try
                            {
                                string colorStr = seriesInfo.Color.Trim();
                                int colorValue;
                                if (colorStr.Length == 6)
                                {
                                    // RGB格式，转换为BGR（Word API使用BGR）
                                    string r = colorStr.Substring(0, 2);
                                    string g = colorStr.Substring(2, 2);
                                    string b = colorStr.Substring(4, 2);
                                    string bgrStr = b + g + r; // BGR格式
                                    colorValue = Convert.ToInt32(bgrStr, 16);
                                }
                                else
                                {
                                    colorValue = Convert.ToInt32(colorStr, 16);
                                }

                                // 根据图表类型设置颜色
                                if (chartInfo.Type == "line")
                                {
                                    series.Format.Line.ForeColor.RGB = colorValue;
                                }
                                else
                                {
                                    series.Format.Fill.ForeColor.RGB = colorValue;
                                }
                            }
                            catch (Exception ex)
                            {
                                errors.Add($"设置Series{i + 1}颜色失败: {ex.Message}");
                            }
                        }

                        // 设置混合图表类型（如果Series指定了ChartType）
                        if (!string.IsNullOrEmpty(seriesInfo.ChartType))
                        {
                            try
                            {
                                Excel.XlChartType chartType = GetChartType(seriesInfo.ChartType);
                                series.ChartType = chartType;
                            }
                            catch (Exception ex)
                            {
                                errors.Add($"设置Series{i + 1}图表类型失败: {ex.Message}");
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        string errorMsg = $"设置Series{i + 1}数据失败: {ex.Message}";
                        errors.Add(errorMsg);
                        System.Diagnostics.Debug.WriteLine($"[DEBUG] {errorMsg}");
                    }
                }

                // 删除多余的系列
                try
                {
                    int totalSeries = chart.SeriesCollection().Count;
                    System.Diagnostics.Debug.WriteLine($"[DEBUG] 填充了 {validSeriesIndex} 个有效系列，图表共有 {totalSeries} 个系列");

                    for (int i = totalSeries; i > validSeriesIndex; i--)
                    {
                        try
                        {
                            chart.SeriesCollection(i).Delete();
                            System.Diagnostics.Debug.WriteLine($"[DEBUG] 已删除多余系列{i}");
                        }
                        catch (Exception ex)
                        {
                            string errorMsg = $"删除多余系列{i}失败: {ex.Message}";
                            errors.Add(errorMsg);
                            System.Diagnostics.Debug.WriteLine($"[DEBUG] {errorMsg}");
                        }
                    }
                }
                catch (Exception ex)
                {
                    string errorMsg = $"删除多余系列异常: {ex.Message}";
                    errors.Add(errorMsg);
                    System.Diagnostics.Debug.WriteLine($"[DEBUG] {errorMsg}");
                }

                if (errors.Count > 0)
                {
                    return new ChartOperationResult { Success = true, Error = string.Join("; ", errors) };
                }

                return new ChartOperationResult { Success = true, Error = null };
            }
            catch (Exception ex)
            {
                string errorMsg = $"填充图表数据异常: {ex.Message}";
                System.Diagnostics.Debug.WriteLine($"[DEBUG] {errorMsg}");
                return new ChartOperationResult { Success = false, Error = errorMsg };
            }
        }


        /// <summary>
        /// 设置网格线
        /// </summary>
        public static ChartOperationResult SetGridlines(Word.Chart chart, string gridlines)
        {
            try
            {
                Word.Axis valueAxis = chart.Axes(Word.XlAxisType.xlValue);
                
                switch (gridlines.ToLower())
                {
                    case "none":
                        valueAxis.HasMajorGridlines = false;
                        valueAxis.HasMinorGridlines = false;
                        break;
                    case "major":
                        valueAxis.HasMajorGridlines = true;
                        valueAxis.HasMinorGridlines = false;
                        // 设置主要网格线为虚线
                        try
                        {
                            valueAxis.MajorGridlines.Format.Line.DashStyle = Microsoft.Office.Core.MsoLineDashStyle.msoLineDash;
                        }
                        catch (Exception ex)
                        {
                            System.Diagnostics.Debug.WriteLine($"[DEBUG] 设置主要网格线虚线样式失败: {ex.Message}");
                        }
                        break;
                    case "minor":
                        valueAxis.HasMajorGridlines = false;
                        valueAxis.HasMinorGridlines = true;
                        // 设置次要网格线为虚线
                        try
                        {
                            valueAxis.MinorGridlines.Format.Line.DashStyle = Microsoft.Office.Core.MsoLineDashStyle.msoLineDash;
                        }
                        catch (Exception ex)
                        {
                            System.Diagnostics.Debug.WriteLine($"[DEBUG] 设置次要网格线虚线样式失败: {ex.Message}");
                        }
                        break;
                    case "both":
                        valueAxis.HasMajorGridlines = true;
                        valueAxis.HasMinorGridlines = true;
                        // 设置主要和次要网格线为虚线
                        try
                        {
                            valueAxis.MajorGridlines.Format.Line.DashStyle = Microsoft.Office.Core.MsoLineDashStyle.msoLineDash;
                            valueAxis.MinorGridlines.Format.Line.DashStyle = Microsoft.Office.Core.MsoLineDashStyle.msoLineDash;
                        }
                        catch (Exception ex)
                        {
                            System.Diagnostics.Debug.WriteLine($"[DEBUG] 设置网格线虚线样式失败: {ex.Message}");
                        }
                        break;
                    default:
                        return new ChartOperationResult { Success = false, Error = $"Gridlines值无效: '{gridlines}'，有效值为: none, major, minor, both" };
                }

                return new ChartOperationResult { Success = true, Error = null };
            }
            catch (Exception ex)
            {
                string errorMsg = $"设置网格线失败: {ex.Message}";
                System.Diagnostics.Debug.WriteLine($"[DEBUG] {errorMsg}");
                return new ChartOperationResult { Success = false, Error = errorMsg };
            }
        }

        /// <summary>
        /// Y轴范围计算结果
        /// </summary>
        internal class YAxisRange
        {
            public double Min { get; set; }
            public double Max { get; set; }
            public double MajorUnit { get; set; }
        }

        /// <summary>
        /// 获取默认颜色数组（BGR格式）
        /// 颜色顺序：粉色、蓝色、橙色、绿色
        /// </summary>
        public static int[] GetDefaultColors()
        {
            // RGB颜色转换为BGR格式（Word API使用BGR）
            // (216,104,164) 粉色 -> (164,104,216) -> A468D8
            // (131,165,215) 蓝色 -> (215,165,131) -> D7A583
            // (225,165,102) 橙色 -> (102,165,225) -> 66A5E1
            // (118,182,204) 绿色 -> (204,182,118) -> CCB676
            return new int[]
            {
                0xA468D8, // 粉色 (216,104,164)
                0xD7A583, // 蓝色 (131,165,215)
                0x66A5E1, // 橙色 (225,165,102)
                0xCCB676  // 绿色 (118,182,204)
            };
        }

        /// <summary>
        /// 计算合适的网格线间隔（major unit）
        /// 目标是让图表上有4-6条网格线，间隔约为跨度的1/4到1/6
        /// </summary>
        public static double CalculateMajorUnit(double span)
        {
            if (span <= 0) return 10;
            
            // 计算数量级（10的幂次）
            double magnitude = Math.Pow(10, Math.Floor(Math.Log10(span)));
            
            // 将跨度归一化到1-10之间
            double normalized = span / magnitude;
            
            // 根据归一化的值选择合适的间隔，确保网格线数量适中（4-6条）
            double unit;
            if (normalized <= 1.5)
            {
                unit = 0.2; // 间隔为0.2 × magnitude，适合跨度较小的情况
            }
            else if (normalized <= 3)
            {
                unit = 0.5; // 间隔为0.5 × magnitude
            }
            else if (normalized <= 6)
            {
                unit = 1; // 间隔为1 × magnitude
            }
            else if (normalized <= 10)
            {
                unit = 2; // 间隔为2 × magnitude
            }
            else
            {
                // 如果跨度超过10倍数量级，使用5倍数量级
                unit = 5;
            }
            
            return unit * magnitude;
        }

        /// <summary>
        /// 自动计算X轴刻度标签间隔
        /// 当标签总数超过10个时，自动计算合适的间隔，使得显示的标签数量在5-10个之间
        /// </summary>
        /// <param name="totalLabelCount">X轴标签总数</param>
        /// <returns>刻度标签间隔（隔多少格显示一个标签）</returns>
        public static int CalculateAutoTickLabelSpacing(int totalLabelCount)
        {
            if (totalLabelCount <= 10)
            {
                return 1; // 10个以内全部显示
            }

            // 目标：显示的标签数量在5-10个之间
            // 计算间隔：间隔 = (总数 - 1) / (目标数量 - 1)
            // 优先选择能显示约8个标签的间隔
            
            int targetDisplayCount = 8; // 目标显示8个标签
            int spacing = (int)Math.Ceiling((double)(totalLabelCount - 1) / (targetDisplayCount - 1));
            
            // 验证：使用这个间隔实际会显示多少个标签
            int actualDisplayCount = (totalLabelCount - 1) / spacing + 1;
            
            // 如果显示的标签数量太少（少于5个），减小间隔
            if (actualDisplayCount < 5)
            {
                spacing = (int)Math.Ceiling((double)(totalLabelCount - 1) / 4); // 至少显示5个
            }
            // 如果显示的标签数量太多（超过10个），增大间隔
            else if (actualDisplayCount > 10)
            {
                spacing = (int)Math.Ceiling((double)(totalLabelCount - 1) / 9); // 最多显示10个
            }
            
            // 确保间隔至少为1
            if (spacing < 1) spacing = 1;
            
            // 重新计算实际显示数量
            actualDisplayCount = (totalLabelCount - 1) / spacing + 1;
            System.Diagnostics.Debug.WriteLine($"[DEBUG] 自动计算刻度标签间隔：总数={totalLabelCount}，间隔={spacing}，将显示{actualDisplayCount}个标签");
            
            return spacing;
        }

        /// <summary>
        /// 根据数据自动计算Y轴范围（针对指定轴）
        /// </summary>
#pragma warning disable CS8632 // 禁用可空引用类型警告
        private static YAxisRange? CalculateYAxisRange(ChartInfo chartInfo, string axisY = "primary")
#pragma warning restore CS8632
        {
            if (chartInfo.SeriesCollection == null || chartInfo.SeriesCollection.Count == 0)
            {
                return null;
            }

            double minValue = double.MaxValue;
            double maxValue = double.MinValue;
            bool hasValidData = false;

            // 遍历指定轴的Series，收集所有Y值
            foreach (var series in chartInfo.SeriesCollection)
            {
                if (series == null || series.Points == null) continue;

                // 只处理指定轴的Series
                string seriesAxisY = string.IsNullOrEmpty(series.AxisY) ? "primary" : series.AxisY;
                if (seriesAxisY != axisY) continue;

                foreach (var point in series.Points)
                {
                    if (point != null)
                    {
                        if (point.Y < minValue) minValue = point.Y;
                        if (point.Y > maxValue) maxValue = point.Y;
                        hasValidData = true;
                    }
                }
            }

            if (!hasValidData)
            {
                return null;
            }

            // 计算跨度
            double span = maxValue - minValue;
            
            // 计算合适的网格线间隔（major unit）
            double majorUnit;
            if (span == 0)
            {
                // 如果跨度为0，使用最小值的绝对值来计算间隔
                majorUnit = CalculateMajorUnit(Math.Abs(minValue));
            }
            else
            {
                majorUnit = CalculateMajorUnit(span);
            }
            
            // 先从数据的最小值和最大值加上0.5个majorUnit的边距
            double paddedMin = minValue - majorUnit * 0.5;
            double paddedMax = maxValue + majorUnit * 0.5;
            
            // 如果跨度为0（所有值相同），设置一个默认范围
            if (span == 0)
            {
                if (minValue == 0)
                {
                    return new YAxisRange { Min = 0, Max = 10, MajorUnit = 10 };
                }
                else
                {
                    // 上下各预留0.5个网格线刻度，然后对齐到网格线
                    double roundedMin = Math.Floor(paddedMin / majorUnit) * majorUnit;
                    double roundedMax = Math.Ceiling(paddedMax / majorUnit) * majorUnit;
                    return new YAxisRange { Min = roundedMin, Max = roundedMax, MajorUnit = majorUnit };
                }
            }
            
            // 将对齐后的范围对齐到网格线刻度（向下取整和向上取整）
            double calculatedMin = Math.Floor(paddedMin / majorUnit) * majorUnit;
            double calculatedMax = Math.Ceiling(paddedMax / majorUnit) * majorUnit;
            
            // 检查对齐后的范围是否超出数据范围太多（超过1个majorUnit）
            // 如果超出太多，使用向上/向下对齐到更接近数据范围的网格线，避免显示完整的空白网格线
            double minPadding = minValue - calculatedMin;
            double maxPadding = calculatedMax - maxValue;
            
            if (minPadding > majorUnit)
            {
                // 下方空白超过1个majorUnit，向上对齐到最近的网格线（更接近数据最小值）
                calculatedMin = Math.Ceiling((minValue - majorUnit * 0.5) / majorUnit) * majorUnit;
            }
            
            if (maxPadding > majorUnit)
            {
                // 上方空白超过1个majorUnit，向下对齐到最近的网格线（更接近数据最大值）
                calculatedMax = Math.Floor((maxValue + majorUnit * 0.5) / majorUnit) * majorUnit;
            }
            
            // 如果数据都是非负数，确保Y轴最小值不小于0
            if (minValue >= 0 && calculatedMin < 0)
            {
                // 将下方空隙移到上方
                double extraPadding = Math.Abs(calculatedMin);
                calculatedMin = 0;
                calculatedMax = Math.Ceiling((maxValue + majorUnit * 0.5 + extraPadding) / majorUnit) * majorUnit;
            }
            
            return new YAxisRange { Min = calculatedMin, Max = calculatedMax, MajorUnit = majorUnit };
        }

        public static SeriesInfo TryGetSeriesInfoFromXml(ChartInfo chartInfo, int seriesIndexOneBased)
        {
            if (chartInfo?.SeriesCollection == null || seriesIndexOneBased < 1)
            {
                return null;
            }

            int idx = seriesIndexOneBased - 1;
            if (idx >= chartInfo.SeriesCollection.Count)
            {
                return null;
            }

            return chartInfo.SeriesCollection[idx];
        }

        /// <summary>
        /// 应用图表特定格式
        /// </summary>
        public static ChartOperationResult ApplyChartSpecificFormatting(Word.Chart chart, ChartInfo chartInfo)
        {
            var errors = new List<string>();

            try
            {
                // 默认隐藏Y轴线，但保留左侧数值刻度标签（适用于折线图和柱状图）
                if (chartInfo.Type == "line" || chartInfo.Type == "column_clustered")
                {
                    try
                    {
                        Word.Axis yAxis = chart.Axes(Word.XlAxisType.xlValue);
                        // 将Y轴标签位置设置为左侧（保留刻度数值）
                        yAxis.TickLabelPosition = Word.XlTickLabelPosition.xlTickLabelPositionLow;
                        // 隐藏Y轴刻度线（不显示刻度标记）
                        yAxis.MajorTickMark = Word.XlTickMark.xlTickMarkNone;
                        yAxis.MinorTickMark = Word.XlTickMark.xlTickMarkNone;
                        // 隐藏Y轴线本身（不显示轴线）
                        yAxis.Format.Line.Visible = Microsoft.Office.Core.MsoTriState.msoFalse;
                        
                        // 设置Y轴标签字体大小（默认9号）
                        try
                        {
                            yAxis.TickLabels.Font.Size = 9;
                            System.Diagnostics.Debug.WriteLine("[DEBUG] 已设置Y轴标签字体大小为9号");
                        }
                        catch (Exception ex)
                        {
                            System.Diagnostics.Debug.WriteLine($"[DEBUG] 设置Y轴标签字体大小失败: {ex.Message}");
                        }
                        
                        System.Diagnostics.Debug.WriteLine("[DEBUG] 已隐藏Y轴线，保留左侧刻度标签");
                    }
                    catch (Exception ex)
                    {
                        string errorMsg = $"设置Y轴格式失败: {ex.Message}";
                        errors.Add(errorMsg);
                        System.Diagnostics.Debug.WriteLine($"[DEBUG] {errorMsg}");
                    }
                }

                // 设置数据标签
                // 如果图表级别启用了数据标签，则检查每个系列的设置
                if (chartInfo.ShowDataLabels)
                {
                    try
                    {
                        // 遍历所有系列设置数据标签
                        int seriesCount = chart.SeriesCollection().Count;
                        for (int i = 1; i <= seriesCount; i++)
                        {
                            dynamic series = chart.SeriesCollection(i);
                            
                            // 检查该系列是否应该显示数据标签
                            // 如果Series的ShowDataLabels为null，继承图表级别设置（显示）
                            // 如果Series的ShowDataLabels为false，不显示
                            // 如果Series的ShowDataLabels为true，显示
                            bool shouldShowLabels = true;
                            var seriesInfo = TryGetSeriesInfoFromXml(chartInfo, i);
                            if (seriesInfo != null && seriesInfo.ShowDataLabels.HasValue)
                            {
                                shouldShowLabels = seriesInfo.ShowDataLabels.Value;
                            }
                            
                            if (shouldShowLabels)
                            {
                                series.HasDataLabels = true;

                            // 对于折线图和柱状图，设置数据标签位置为上方，避免与线条重叠
                            if (chartInfo.Type == "line" || chartInfo.Type == "column_clustered")
                            {
                                try
                                {
                                    // 设置数据标签位置为上方（Above）
                                    // xlLabelPositionAbove = -4128
                                    series.DataLabels.Position = Word.XlDataLabelPosition.xlLabelPositionAbove;
                                    
                                    // 设置数据标签格式：只显示数值，不显示其他信息
                                    series.DataLabels.ShowValue = true;
                                    series.DataLabels.ShowCategoryName = false;
                                    series.DataLabels.ShowSeriesName = false;
                                    series.DataLabels.ShowPercentage = false;
                                    series.DataLabels.ShowLegendKey = false;
                                    
                                    System.Diagnostics.Debug.WriteLine($"[DEBUG] 已设置系列{i}的数据标签位置为上方");
                                }
                                catch (Exception ex)
                                {
                                    System.Diagnostics.Debug.WriteLine($"[DEBUG] 设置系列{i}数据标签位置失败: {ex.Message}");
                                }
                            }

                            // 设置数据标签类型（针对饼图）
                            if (chartInfo.Type == "pie3d" && !string.IsNullOrEmpty(chartInfo.DataLabelType))
                            {
                                try
                                {
                                    switch (chartInfo.DataLabelType.ToLower())
                                    {
                                        case "percent":
                                            series.DataLabels.ShowPercentage = true;
                                            series.DataLabels.ShowValue = false;
                                            break;
                                        case "value":
                                            series.DataLabels.ShowPercentage = false;
                                            series.DataLabels.ShowValue = true;
                                            break;
                                        case "both":
                                            series.DataLabels.ShowPercentage = true;
                                            series.DataLabels.ShowValue = true;
                                            break;
                                        default:
                                            errors.Add($"DataLabelType值无效: '{chartInfo.DataLabelType}'，有效值为: percent, value, both");
                                            break;
                                    }
                                }
                                catch (Exception ex)
                                {
                                    string errorMsg = $"设置数据标签类型失败: {ex.Message}";
                                    errors.Add(errorMsg);
                                    System.Diagnostics.Debug.WriteLine($"[DEBUG] {errorMsg}");
                                }
                            }
                            }
                            else
                            {
                                // 该系列不显示数据标签
                                series.HasDataLabels = false;
                                System.Diagnostics.Debug.WriteLine($"[DEBUG] 系列{i}不显示数据标签（Series.ShowDataLabels=false）");
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        string errorMsg = $"设置数据标签失败: {ex.Message}";
                        errors.Add(errorMsg);
                        System.Diagnostics.Debug.WriteLine($"[DEBUG] {errorMsg}");
                    }
                }
                else
                {
                    // 图表级别未启用数据标签，但需要检查是否有Series单独启用了
                    try
                    {
                        int seriesCount = chart.SeriesCollection().Count;
                        for (int i = 1; i <= seriesCount; i++)
                        {
                            dynamic series = chart.SeriesCollection(i);
                            
                            // 如果Series明确设置为true，则显示（即使图表级别为false）
                            bool shouldShowLabels = false;
                            var seriesInfo = TryGetSeriesInfoFromXml(chartInfo, i);
                            if (seriesInfo != null
                                && seriesInfo.ShowDataLabels.HasValue
                                && seriesInfo.ShowDataLabels.Value)
                            {
                                shouldShowLabels = true;
                            }
                            
                            if (shouldShowLabels)
                            {
                                series.HasDataLabels = true;
                                
                                // 对于折线图和柱状图，设置数据标签位置为上方
                                if (chartInfo.Type == "line" || chartInfo.Type == "column_clustered")
                                {
                                    try
                                    {
                                        series.DataLabels.Position = Word.XlDataLabelPosition.xlLabelPositionAbove;
                                        series.DataLabels.ShowValue = true;
                                        series.DataLabels.ShowCategoryName = false;
                                        series.DataLabels.ShowSeriesName = false;
                                        series.DataLabels.ShowPercentage = false;
                                        series.DataLabels.ShowLegendKey = false;
                                        System.Diagnostics.Debug.WriteLine($"[DEBUG] 系列{i}单独启用了数据标签");
                                    }
                                    catch (Exception ex)
                                    {
                                        System.Diagnostics.Debug.WriteLine($"[DEBUG] 设置系列{i}数据标签失败: {ex.Message}");
                                    }
                                }
                            }
                            else
                            {
                                series.HasDataLabels = false;
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        string errorMsg = $"设置数据标签失败: {ex.Message}";
                        errors.Add(errorMsg);
                        System.Diagnostics.Debug.WriteLine($"[DEBUG] {errorMsg}");
                    }
                }

                // 设置图表颜色（仅对非折线图，折线图的颜色在折线图特定设置中处理）
                if (chartInfo.Type != "line")
                {
                    try
                    {
                        int seriesCount = chart.SeriesCollection().Count;
                        int[] defaultColors = GetDefaultColors();
                        
                        for (int i = 1; i <= seriesCount; i++)
                        {
                            dynamic series = chart.SeriesCollection(i);
                            int colorValue;
                            
                            if (!string.IsNullOrEmpty(chartInfo.PlotColor))
                            {
                                // 用户指定了颜色，使用用户指定的颜色
                                string colorStr = chartInfo.PlotColor.Trim();
                                if (colorStr.Length == 6)
                                {
                                    // 将RGB转换为BGR格式（交换R和B的位置）
                                    string r = colorStr.Substring(0, 2);
                                    string g = colorStr.Substring(2, 2);
                                    string b = colorStr.Substring(4, 2);
                                    string bgrStr = b + g + r; // BGR格式
                                    colorValue = Convert.ToInt32(bgrStr, 16);
                                }
                                else
                                {
                                    colorValue = Convert.ToInt32(colorStr, 16);
                                }
                            }
                            else
                            {
                                // 用户没有指定颜色，使用默认颜色（按系列顺序）
                                int colorIndex = (i - 1) % defaultColors.Length; // 循环使用颜色
                                colorValue = defaultColors[colorIndex];
                            }
                            
                            series.Format.Fill.ForeColor.RGB = colorValue;
                        }
                    }
                    catch (Exception ex)
                    {
                        string errorMsg = $"设置颜色失败: {ex.Message}（颜色值格式可能不正确，应为16进制字符串，如'FF0000'）";
                        errors.Add(errorMsg);
                        System.Diagnostics.Debug.WriteLine($"[DEBUG] {errorMsg}");
                    }
                }

                // 柱状图特定设置
                if (chartInfo.Type == "column_clustered" && chartInfo.GapWidth > 0)
                {
                    try
                    {
                        dynamic series1 = chart.SeriesCollection(1);
                        series1.GapWidth = chartInfo.GapWidth;
                    }
                    catch (Exception ex)
                    {
                        string errorMsg = $"设置柱间隙失败: {ex.Message}";
                        errors.Add(errorMsg);
                        System.Diagnostics.Debug.WriteLine($"[DEBUG] {errorMsg}");
                    }
                }

                // 折线图特定设置
                if (chartInfo.Type == "line")
                {
                    try
                    {
                        // 遍历所有系列，设置数据点标记（默认不显示）和线条颜色
                        int seriesCount = chart.SeriesCollection().Count;
                        for (int i = 1; i <= seriesCount; i++)
                        {
                            dynamic series = chart.SeriesCollection(i);
                            
                            // 确保线条可见（必须先设置）
                            try
                            {
                                series.Format.Line.Visible = Microsoft.Office.Core.MsoTriState.msoTrue;
                            }
                            catch (Exception ex)
                            {
                                System.Diagnostics.Debug.WriteLine($"[DEBUG] 设置线条可见性失败: {ex.Message}");
                            }
                            
                            // 设置折线图线条颜色（仅当Series未指定颜色时）
                            try
                            {
                                // 检查当前Series是否已经在前面设置了颜色
                                bool seriesColorSet = false;
                                var seriesInfo = TryGetSeriesInfoFromXml(chartInfo, i);
                                if (seriesInfo != null && !string.IsNullOrEmpty(seriesInfo.Color))
                                {
                                    // Series已经指定了颜色，跳过设置（颜色已在前面设置）
                                    seriesColorSet = true;
                                }
                                
                                // 只有当Series没有指定颜色时，才设置默认颜色
                                if (!seriesColorSet)
                                {
                                    int colorValue;
                                    // 使用默认颜色（按系列顺序）
                                    int[] defaultColors = GetDefaultColors();
                                    int colorIndex = (i - 1) % defaultColors.Length; // 循环使用颜色
                                    colorValue = defaultColors[colorIndex];
                                    series.Format.Line.ForeColor.RGB = colorValue;
                                }
                            }
                            catch (Exception ex)
                            {
                                string errorMsg = $"设置折线图线条颜色失败: {ex.Message}";
                                errors.Add(errorMsg);
                                System.Diagnostics.Debug.WriteLine($"[DEBUG] {errorMsg}");
                            }
                            
                            // 默认不显示数据点标记（节点）
                            if (chartInfo.DataMarkers)
                            {
                                series.MarkerStyle = Word.XlMarkerStyle.xlMarkerStyleCircle;
                                if (chartInfo.MarkerSize > 0)
                                {
                                    series.MarkerSize = chartInfo.MarkerSize;
                                }
                            }
                            else
                            {
                                // 明确设置为无标记
                                series.MarkerStyle = Word.XlMarkerStyle.xlMarkerStyleNone;
                            }

                            if (chartInfo.LineWeight > 0)
                            {
                                series.Format.Line.Weight = (float)chartInfo.LineWeight;
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        string errorMsg = $"设置折线图属性失败: {ex.Message}";
                        errors.Add(errorMsg);
                        System.Diagnostics.Debug.WriteLine($"[DEBUG] {errorMsg}");
                    }
                }

                // 设置Y轴范围（适用于折线图和柱状图）
                if (chartInfo.Type == "line" || chartInfo.Type == "column_clustered")
                {
                    try
                    {
                        // 检查是否有secondary轴的Series
                        bool hasSecondaryAxis = false;
                        if (chartInfo.SeriesCollection != null)
                        {
                            foreach (var series in chartInfo.SeriesCollection)
                            {
                                if (series != null && series.AxisY == "secondary")
                                {
                                    hasSecondaryAxis = true;
                                    break;
                                }
                            }
                        }

                        // 设置主Y轴
                        Word.Axis primaryYAxis = chart.Axes(Word.XlAxisType.xlValue);
                        
                        // 如果用户没有设置Y轴范围，自动根据数据计算
                        if (chartInfo.YAxisMin == 0 && chartInfo.YAxisMax == 0)
                        {
                            var calculatedRange = CalculateYAxisRange(chartInfo, "primary");
                            if (calculatedRange != null)
                            {
                                primaryYAxis.MinimumScale = calculatedRange.Min;
                                primaryYAxis.MaximumScale = calculatedRange.Max;
                                // 设置计算出的主要单位，确保刻度是整数
                                primaryYAxis.MajorUnit = calculatedRange.MajorUnit;
                                System.Diagnostics.Debug.WriteLine($"[DEBUG] 自动计算主Y轴范围: Min={calculatedRange.Min}, Max={calculatedRange.Max}, MajorUnit={calculatedRange.MajorUnit}");
                            }
                        }
                        else
                        {
                            // 用户明确设置了Y轴范围（仅适用于主Y轴）
                            if (chartInfo.YAxisMin != 0) primaryYAxis.MinimumScale = chartInfo.YAxisMin;
                            if (chartInfo.YAxisMax != 0) primaryYAxis.MaximumScale = chartInfo.YAxisMax;
                        }
                        
                        // 设置主要单位（如果用户没有设置，且自动计算也没有设置，则使用默认值）
                        if (chartInfo.YAxisMajorUnit > 0)
                        {
                            primaryYAxis.MajorUnit = chartInfo.YAxisMajorUnit;
                        }
                        else if (chartInfo.YAxisMin == 0 && chartInfo.YAxisMax == 0)
                        {
                            // 如果自动计算了范围，MajorUnit已经在上面设置了
                            // 这里不需要额外设置
                        }

                        // 如果有secondary轴，设置次Y轴
                        if (hasSecondaryAxis)
                        {
                            Word.Axis secondaryYAxis = chart.Axes(Word.XlAxisType.xlValue, Word.XlAxisGroup.xlSecondary);
                            
                            // 自动计算次Y轴范围
                            var secondaryRange = CalculateYAxisRange(chartInfo, "secondary");
                            if (secondaryRange != null)
                            {
                                secondaryYAxis.MinimumScale = secondaryRange.Min;
                                secondaryYAxis.MaximumScale = secondaryRange.Max;
                                secondaryYAxis.MajorUnit = secondaryRange.MajorUnit;
                                System.Diagnostics.Debug.WriteLine($"[DEBUG] 自动计算次Y轴范围: Min={secondaryRange.Min}, Max={secondaryRange.Max}, MajorUnit={secondaryRange.MajorUnit}");
                            }
                            
                            // 设置次Y轴标签字体大小（默认9号）
                            try
                            {
                                secondaryYAxis.TickLabels.Font.Size = 9;
                                System.Diagnostics.Debug.WriteLine("[DEBUG] 已设置次Y轴标签字体大小为9号");
                            }
                            catch (Exception ex)
                            {
                                System.Diagnostics.Debug.WriteLine($"[DEBUG] 设置次Y轴标签字体大小失败: {ex.Message}");
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        string errorMsg = $"设置Y轴范围失败: {ex.Message}";
                        errors.Add(errorMsg);
                        System.Diagnostics.Debug.WriteLine($"[DEBUG] {errorMsg}");
                    }
                }

                // 设置X轴刻度线类型和刻度标签显示（适用于折线图和柱状图）
                if (chartInfo.Type == "line" || chartInfo.Type == "column_clustered")
                {
                    try
                    {
                        Word.Axis xAxis = chart.Axes(Word.XlAxisType.xlCategory);
                        
                        // 设置主刻度线类型为"外部"（Outside）- 显示主刻度线，标签对齐到主刻度线
                        xAxis.MajorTickMark = Word.XlTickMark.xlTickMarkOutside;
                        
                        // 设置次刻度线类型为"无"（None）- 隐藏次刻度线
                        xAxis.MinorTickMark = Word.XlTickMark.xlTickMarkNone;
                        
                        // 获取X轴标签的总数
                        int totalLabelCount = 0;
                        if (chartInfo.SeriesCollection != null && chartInfo.SeriesCollection.Count > 0)
                        {
                            var firstSeries = chartInfo.SeriesCollection[0];
                            if (firstSeries != null && firstSeries.Points != null)
                            {
                                totalLabelCount = firstSeries.Points.Count;
                            }
                        }
                        
                        // 计算刻度标签间隔
                        int tickLabelSpacing = 1; // 默认值
                        
                        if (chartInfo.AxisX != null)
                        {
                            // 优先级1：如果指定了TickLabelCount（显示多少个刻度标签）
                            if (chartInfo.AxisX.TickLabelCount.HasValue && chartInfo.AxisX.TickLabelCount.Value > 0)
                            {
                                int targetCount = chartInfo.AxisX.TickLabelCount.Value;
                                if (totalLabelCount > 0 && targetCount < totalLabelCount)
                                {
                                    // 计算间隔：如果总共有N个标签，要显示M个，则间隔为 N/(M-1) 向上取整
                                    tickLabelSpacing = (int)Math.Ceiling((double)(totalLabelCount - 1) / (targetCount - 1));
                                    if (tickLabelSpacing < 1) tickLabelSpacing = 1;
                                    System.Diagnostics.Debug.WriteLine($"[DEBUG] 根据TickLabelCount={targetCount}计算刻度标签间隔: {tickLabelSpacing}");
                                }
                                else
                                {
                                    tickLabelSpacing = 1; // 如果要显示的标签数大于等于总数，则全部显示
                                    System.Diagnostics.Debug.WriteLine($"[DEBUG] TickLabelCount={targetCount}大于等于总数{totalLabelCount}，全部显示");
                                }
                            }
                            // 优先级2：如果指定了TickLabelSpacing（隔多少格显示一个）
                            else if (chartInfo.AxisX.TickLabelSpacing.HasValue && chartInfo.AxisX.TickLabelSpacing.Value > 0)
                            {
                                tickLabelSpacing = chartInfo.AxisX.TickLabelSpacing.Value;
                                System.Diagnostics.Debug.WriteLine($"[DEBUG] 使用指定的TickLabelSpacing: {tickLabelSpacing}");
                            }
                            // 优先级3：默认行为（当超过10个刻度时自动计算）
                            else if (totalLabelCount > 10)
                            {
                                tickLabelSpacing = CalculateAutoTickLabelSpacing(totalLabelCount);
                                System.Diagnostics.Debug.WriteLine($"[DEBUG] 自动计算刻度标签间隔（总数={totalLabelCount}）: {tickLabelSpacing}");
                            }
                            else
                            {
                                tickLabelSpacing = 1; // 默认全部显示
                                System.Diagnostics.Debug.WriteLine($"[DEBUG] 使用默认刻度标签间隔: {tickLabelSpacing}");
                            }
                        }
                        else if (totalLabelCount > 10)
                        {
                            // 如果没有AxisX配置，但超过10个刻度，也自动计算
                            tickLabelSpacing = CalculateAutoTickLabelSpacing(totalLabelCount);
                            System.Diagnostics.Debug.WriteLine($"[DEBUG] 自动计算刻度标签间隔（总数={totalLabelCount}，无AxisX配置）: {tickLabelSpacing}");
                        }
                        
                        // 设置刻度标签间隔
                        xAxis.TickLabelSpacing = tickLabelSpacing;
                        
                        // 设置刻度线间隔（与标签间隔保持一致）
                        xAxis.TickMarkSpacing = tickLabelSpacing;
                        
                        // 刻度在类别之间（柱形默认 true）或刻度线上（折线默认 false）；XML BetweenCategories 可覆盖
                        bool betweenCategories = ResolveAxisBetweenCategories(chartInfo);
                        try
                        {
                            xAxis.AxisBetweenCategories = betweenCategories;
                            System.Diagnostics.Debug.WriteLine(
                                $"[DEBUG] 设置标签位置：AxisBetweenCategories={betweenCategories}（type={chartInfo.Type}）");
                        }
                        catch (Exception ex)
                        {
                            System.Diagnostics.Debug.WriteLine($"[DEBUG] 设置AxisBetweenCategories失败: {ex.Message}");
                        }
                        
                        // 设置X轴标签字体大小（默认9号）
                        try
                        {
                            xAxis.TickLabels.Font.Size = 9;
                            System.Diagnostics.Debug.WriteLine("[DEBUG] 已设置X轴标签字体大小为9号");
                        }
                        catch (Exception ex)
                        {
                            System.Diagnostics.Debug.WriteLine($"[DEBUG] 设置X轴标签字体大小失败: {ex.Message}");
                        }
                        
                        System.Diagnostics.Debug.WriteLine($"[DEBUG] 设置X轴刻度线：主刻度线=外部（显示），次刻度线=无（隐藏），标签间隔={tickLabelSpacing}，刻度线间隔={tickLabelSpacing}，AxisBetweenCategories={ResolveAxisBetweenCategories(chartInfo)}");
                    }
                    catch (Exception ex)
                    {
                        string errorMsg = $"设置X轴刻度线失败: {ex.Message}";
                        errors.Add(errorMsg);
                        System.Diagnostics.Debug.WriteLine($"[DEBUG] {errorMsg}");
                    }
                }

                ApplyAxisTitles(chart, chartInfo, errors);

                // 饼图特定设置
                if (chartInfo.Type == "pie3d" && chartInfo.Explosion > 0)
                {
                    try
                    {
                        dynamic series1 = chart.SeriesCollection(1);
                        series1.Explosion = chartInfo.Explosion;
                    }
                    catch (Exception ex)
                    {
                        string errorMsg = $"设置分离程度失败: {ex.Message}";
                        errors.Add(errorMsg);
                        System.Diagnostics.Debug.WriteLine($"[DEBUG] {errorMsg}");
                    }
                }

                if (errors.Count > 0)
                {
                    return new ChartOperationResult { Success = true, Error = string.Join("; ", errors) };
                }

                return new ChartOperationResult { Success = true, Error = null };
            }
            catch (Exception ex)
            {
                string errorMsg = $"应用图表格式异常: {ex.Message}";
                System.Diagnostics.Debug.WriteLine($"[DEBUG] {errorMsg}");
                return new ChartOperationResult { Success = false, Error = errorMsg };
            }
        }

        /// <summary>
        /// 对已有图表套用 ChartConfig（PR-3 apply 工具入口）。
        /// </summary>
        public static ChartApplyResult ApplyToExistingChart(
            Word.Chart chart,
            Word.InlineShape inlineShape,
            ChartInfo chartInfo,
            string scope)
        {
            var result = new ChartApplyResult();
            if (chart == null || chartInfo == null)
            {
                result.Success = false;
                result.Error = "图表或配置不可用";
                return result;
            }

            Word.Application app = null;
            try
            {
                app = inlineShape?.Application ?? chart.Application;
            }
            catch
            {
                // ignore
            }

            ChartComUiState uiState = ChartComUiHelper.EnterBatch(app, chart);

            try
            {
                return ApplyToExistingChartCore(chart, inlineShape, chartInfo, scope);
            }
            finally
            {
                ChartComUiHelper.ExitBatch(app, uiState, chart);
            }
        }

        public static ChartApplyResult ApplyToExistingChartCore(
            Word.Chart chart,
            Word.InlineShape inlineShape,
            ChartInfo chartInfo,
            string scope)
        {
            var result = new ChartApplyResult();
            if (chart == null || chartInfo == null)
            {
                result.Success = false;
                result.Error = "图表或配置不可用";
                return result;
            }

            string normalizedScope = string.IsNullOrEmpty(scope) ? "all" : scope.Trim().ToLowerInvariant();
            var warnings = new List<string>();

            try
            {
                if (normalizedScope == "data" || normalizedScope == "all")
                {
                    var fillResult = FillChartData(chart, chartInfo);
                    if (!fillResult.Success)
                    {
                        result.Success = false;
                        result.Error = fillResult.Error;
                        return result;
                    }

                    if (!string.IsNullOrEmpty(fillResult.Error))
                    {
                        warnings.Add(fillResult.Error);
                    }

                    result.SeriesUpdated = chartInfo.SeriesCollection?.Count ?? 0;
                    result.PointsUpdated = chartInfo.SeriesCollection != null
                        && chartInfo.SeriesCollection.Count > 0
                        && chartInfo.SeriesCollection[0].Points != null
                        ? chartInfo.SeriesCollection[0].Points.Count
                        : 0;
                }

                if (normalizedScope == "format" || normalizedScope == "all")
                {
                    var formatWarnings = ApplyFormatToExistingChart(chart, inlineShape, chartInfo);
                    warnings.AddRange(formatWarnings);
                }

                result.Success = true;
                result.Warnings = warnings;
                System.Diagnostics.Debug.WriteLine(
                    $"[ApplyChart] 套用完成 scope={normalizedScope} warnings={warnings.Count}");
                return result;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.Error = $"套用图表失败: {ex.Message}";
                return result;
            }
        }

        public static void ApplyChartTitleFormat(
            Word.Chart chart,
            ChartInfo chartInfo,
            ICollection<string> messages)
        {
            if (chart == null || chartInfo == null || !chart.HasTitle)
            {
                return;
            }

            try
            {
                int titleFontSize = chartInfo.TitleFontSize > 0 ? chartInfo.TitleFontSize : 11;
                chart.ChartTitle.Format.TextFrame2.TextRange.Font.Size = titleFontSize;

                if (chartInfo.TitleFontBold)
                {
                    chart.ChartTitle.Format.TextFrame2.TextRange.Font.Bold =
                        Microsoft.Office.Core.MsoTriState.msoTrue;
                }

                if (!string.IsNullOrEmpty(chartInfo.TitleFontColor))
                {
                    Word.WdColor titleColor = TableFormatExtractor.ConvertRgbHexToWdColor(chartInfo.TitleFontColor);
                    if (titleColor == Word.WdColor.wdColorAutomatic)
                    {
                        messages.Add($"标题颜色无效: {chartInfo.TitleFontColor}（须为 RGB 十六进制，如 FF0000）");
                    }
                    else
                    {
                        chart.ChartTitle.Font.Color = titleColor;
                        System.Diagnostics.Debug.WriteLine(
                            $"[DEBUG] 已设置标题颜色 TitleFontColor={chartInfo.TitleFontColor}");
                    }
                }
            }
            catch (Exception ex)
            {
                string errorMsg = $"设置标题字体/颜色失败: {ex.Message}";
                if (!string.IsNullOrEmpty(chartInfo.TitleFontColor))
                {
                    errorMsg += "（TitleFontColor 须为十六进制 RGB，如 FF0000）";
                }

                messages.Add(errorMsg);
                System.Diagnostics.Debug.WriteLine($"[DEBUG] {errorMsg}");
            }
        }

        public static List<string> ApplyFormatToExistingChart(
            Word.Chart chart,
            Word.InlineShape inlineShape,
            ChartInfo chartInfo)
        {
            var warnings = new List<string>();

            if (!string.IsNullOrEmpty(chartInfo.Type))
            {
                try
                {
                    chart.ChartType = (Microsoft.Office.Core.XlChartType)GetChartType(chartInfo.Type);
                }
                catch (Exception ex)
                {
                    warnings.Add($"设置图表类型失败: {ex.Message}");
                }
            }

            if (inlineShape != null)
            {
                try
                {
                    int configuredWidth = chartInfo.Width > 0 ? chartInfo.Width : (int)inlineShape.Width;
                    int configuredHeight = chartInfo.Height > 0 ? chartInfo.Height : (int)inlineShape.Height;
                    if (configuredWidth > 400)
                    {
                        double scaleRatio = 400.0 / configuredWidth;
                        inlineShape.Width = 400;
                        inlineShape.Height = (int)Math.Round(configuredHeight * scaleRatio);
                    }
                    else if (chartInfo.Width > 0 || chartInfo.Height > 0)
                    {
                        inlineShape.Width = configuredWidth;
                        inlineShape.Height = configuredHeight;
                    }
                }
                catch (Exception ex)
                {
                    warnings.Add($"设置图表尺寸失败: {ex.Message}");
                }
            }

            if (!string.IsNullOrEmpty(chartInfo.Theme))
            {
                try
                {
                    chart.ChartStyle = GetChartStyle(chartInfo.Theme);
                }
                catch (Exception ex)
                {
                    warnings.Add($"设置主题失败: {ex.Message}");
                }
            }

            if (!string.IsNullOrEmpty(chartInfo.Title))
            {
                try
                {
                    chart.HasTitle = true;
                    chart.ChartTitle.Text = chartInfo.Title;
                    ApplyChartTitleFormat(chart, chartInfo, warnings);
                }
                catch (Exception ex)
                {
                    warnings.Add($"设置标题失败: {ex.Message}");
                }
            }
            else if (chartInfo.TitleFontSize > 0
                || chartInfo.TitleFontBold
                || !string.IsNullOrEmpty(chartInfo.TitleFontColor))
            {
                try
                {
                    if (chart.HasTitle)
                    {
                        ApplyChartTitleFormat(chart, chartInfo, warnings);
                    }
                    else
                    {
                        warnings.Add("TitleFont* 需要图表已有标题或 XML 中提供 Title");
                    }
                }
                catch (Exception ex)
                {
                    warnings.Add($"设置标题字体失败: {ex.Message}");
                }
            }
            else
            {
                try
                {
                    chart.HasTitle = false;
                }
                catch (Exception ex)
                {
                    warnings.Add($"隐藏标题失败: {ex.Message}");
                }
            }

            try
            {
                chart.HasLegend = true;
                string legendPosition = !string.IsNullOrEmpty(chartInfo.LegendPosition)
                    ? chartInfo.LegendPosition
                    : "top";
                chart.Legend.Position = (Word.XlLegendPosition)GetLegendPosition(legendPosition);
            }
            catch (Exception ex)
            {
                warnings.Add($"设置图例失败: {ex.Message}");
            }

            if (!string.IsNullOrEmpty(chartInfo.Gridlines))
            {
                var gridResult = SetGridlines(chart, chartInfo.Gridlines);
                if (!gridResult.Success)
                {
                    warnings.Add($"设置网格线失败: {gridResult.Error}");
                }
                else if (!string.IsNullOrEmpty(gridResult.Error))
                {
                    warnings.Add(gridResult.Error);
                }
            }

            var formatResult = ApplyChartSpecificFormatting(chart, chartInfo);
            if (!formatResult.Success)
            {
                warnings.Add($"应用图表格式失败: {formatResult.Error}");
            }
            else if (!string.IsNullOrEmpty(formatResult.Error))
            {
                warnings.Add(formatResult.Error);
            }

            return warnings;
        }
    }
}
