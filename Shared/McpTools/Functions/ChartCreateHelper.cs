using System;
using System.Collections.Generic;
using System.Linq;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    public static class ChartCreateHelper
    {
        /// <summary>
        /// 验证配置
        /// </summary>
        public static ChartValidationResult ValidateConfig(ChartConfig config)
        {
            var errors = new List<string>();

            if (config == null)
            {
                return new ChartValidationResult { IsValid = false, Error = "配置对象为空" };
            }

            if (config.Chart == null)
            {
                errors.Add("缺少Chart配置");
            }
            else
            {
                if (string.IsNullOrEmpty(config.Chart.Type))
                {
                    errors.Add("Chart.Type字段为空或缺失");
                }
                else
                {
                    // 验证图表类型
                    string[] validTypes = { "column_clustered", "line", "pie3d" };
                    if (!validTypes.Contains(config.Chart.Type))
                    {
                        errors.Add($"Chart.Type值无效: '{config.Chart.Type}'，有效值为: {string.Join(", ", validTypes)}");
                    }
                }

                // 标题现在是可选的，不再验证

                // 验证SeriesCollection
                if (config.Chart.SeriesCollection == null)
                {
                    errors.Add("Chart.SeriesCollection字段为空或缺失");
                }
                else if (config.Chart.SeriesCollection.Count == 0)
                {
                    errors.Add("Chart.SeriesCollection至少需要1个Series");
                }
                else
                {
                    // 收集所有Series的X轴标签，用于验证一致性
                    var allXLabels = new HashSet<string>();
                    bool hasXLabels = false;

                    for (int i = 0; i < config.Chart.SeriesCollection.Count; i++)
                    {
                        var series = config.Chart.SeriesCollection[i];
                        if (series == null)
                        {
                            errors.Add($"SeriesCollection第{i + 1}个Series为空");
                            continue;
                        }

                        // 验证Series名称
                        if (string.IsNullOrEmpty(series.Name))
                        {
                            errors.Add($"SeriesCollection第{i + 1}个Series的Name属性为空");
                        }

                        // 验证AxisY属性
                        if (!string.IsNullOrEmpty(series.AxisY) && series.AxisY != "primary" && series.AxisY != "secondary")
                        {
                            errors.Add($"SeriesCollection第{i + 1}个Series的AxisY属性无效: '{series.AxisY}'，有效值为: primary, secondary");
                        }

                        // 验证Points
                        if (series.Points == null)
                        {
                            errors.Add($"SeriesCollection第{i + 1}个Series的Points字段为空");
                        }
                        else if (series.Points.Count == 0)
                        {
                            errors.Add($"SeriesCollection第{i + 1}个Series至少需要1个Point");
                        }
                        else
                        {
                            // 验证每个Point
                            for (int j = 0; j < series.Points.Count; j++)
                            {
                                var point = series.Points[j];
                                if (point == null)
                                {
                                    errors.Add($"SeriesCollection第{i + 1}个Series的第{j + 1}个Point为空");
                                }
                                else
                                {
                                    // 验证X轴标签
                                    if (!string.IsNullOrEmpty(point.X))
                                    {
                                        allXLabels.Add(point.X);
                                        hasXLabels = true;
                                    }

                                    // Y值可以为0，不强制验证
                                }
                            }
                        }
                    }

                    // 验证X轴标签一致性：所有Series的Point应该有相同的X值集合
                    if (hasXLabels && config.Chart.SeriesCollection.Count > 1)
                    {
                        var firstSeriesXLabels = new HashSet<string>();
                        foreach (var point in config.Chart.SeriesCollection[0].Points)
                        {
                            if (!string.IsNullOrEmpty(point.X))
                            {
                                firstSeriesXLabels.Add(point.X);
                            }
                        }

                        for (int i = 1; i < config.Chart.SeriesCollection.Count; i++)
                        {
                            var seriesXLabels = new HashSet<string>();
                            foreach (var point in config.Chart.SeriesCollection[i].Points)
                            {
                                if (!string.IsNullOrEmpty(point.X))
                                {
                                    seriesXLabels.Add(point.X);
                                }
                            }

                            if (!firstSeriesXLabels.SetEquals(seriesXLabels))
                            {
                                errors.Add($"SeriesCollection第{i + 1}个Series的X轴标签与第1个Series不一致");
                            }
                        }
                    }
                }
            }

            if (errors.Count > 0)
            {
                return new ChartValidationResult { IsValid = false, Error = string.Join("; ", errors) };
            }

            return new ChartValidationResult { IsValid = true, Error = null };
        }

        /// <summary>
        /// 创建图表
        /// </summary>
        public static ChartCreateResult CreateAtLocation(Word.Document document, ChartConfig config, int? position = null)
        {
            var errors = new List<string>();
            Word.Application app = document?.Application;
            ChartComUiState uiState = ChartComUiHelper.EnterBatch(app);
            Word.Chart chartForUi = null;

            try
            {
                return CreateAtLocationCore(document, config, position, uiState, ref chartForUi, errors);
            }
            finally
            {
                ChartComUiHelper.ExitBatch(app, uiState, chartForUi);
            }
        }

        private static ChartCreateResult CreateAtLocationCore(
            Word.Document document,
            ChartConfig config,
            int? position,
            ChartComUiState uiState,
            ref Word.Chart chartForUi,
            List<string> errors)
        {
            errors.Clear();

            try
            {
                Word.Range insertRange;

                if (position.HasValue)
                {
                    // 在指定位置创建图表
                    try
                    {
                        if (position.Value < 0)
                        {
                            return new ChartCreateResult { Success = false, Error = $"位置参数无效: {position.Value}，必须大于等于0" };
                        }

                        if (position.Value > document.Content.End)
                        {
                            System.Diagnostics.Debug.WriteLine($"[DEBUG] 位置 {position.Value} 超出文档范围（文档长度: {document.Content.End}），将在文档末尾创建");
                            insertRange = document.Content;
                            insertRange.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                        }
                        else
                        {
                            insertRange = document.Range(position.Value, position.Value);
                            insertRange.Collapse(Word.WdCollapseDirection.wdCollapseStart);
                            System.Diagnostics.Debug.WriteLine($"[DEBUG] 在位置 {position.Value} 创建图表");
                        }
                    }
                    catch (Exception ex)
                    {
                        string errorMsg = $"无法在指定位置创建图表: {ex.Message}，将使用文档末尾";
                        System.Diagnostics.Debug.WriteLine($"[DEBUG] {errorMsg}");
                        insertRange = document.Content;
                        insertRange.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                    }
                }
                else
                {
                    // 默认在当前光标处插入图表
                    insertRange = document.Application.Selection.Range;
                    System.Diagnostics.Debug.WriteLine("[DEBUG] 在当前光标处创建图表");
                }

                // 创建图表（Word 会启动内嵌 Excel 作为数据源，无法完全禁止进程启动）
                Word.InlineShape chartShape = document.InlineShapes.AddChart();
                Word.Chart chart = chartShape.Chart;
                chartForUi = chart;
                ChartComUiHelper.TrySuppressEmbeddedExcel(chart, uiState);

                // 立即禁用标题（防止Word自动添加标题），如果后面有提供标题再启用
                try
                {
                    chart.HasTitle = false;
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"[DEBUG] 初始禁用标题失败: {ex.Message}");
                }

                // 设置图表类型（通过ChartType属性）
                try
                {
                    chart.ChartType = (Microsoft.Office.Core.XlChartType)ChartEngine.GetChartType(config.Chart.Type);
                }
                catch (Exception ex)
                {
                    string errorMsg = $"设置图表类型失败: {ex.Message}";
                    errors.Add(errorMsg);
                    System.Diagnostics.Debug.WriteLine($"[DEBUG] {errorMsg}");
                }

                // 设置图表大小（宽度不能超过400，如果超过则按比例缩放宽高）
                int configuredWidth = config.Chart.Width > 0 ? config.Chart.Width : 400;
                int configuredHeight = config.Chart.Height > 0 ? config.Chart.Height : 300;
                
                if (configuredWidth > 400)
                {
                    // 计算缩放比例
                    double scaleRatio = 400.0 / configuredWidth;
                    chartShape.Width = 400;
                    chartShape.Height = (int)Math.Round(configuredHeight * scaleRatio);
                    System.Diagnostics.Debug.WriteLine($"[DEBUG] 图表宽度{configuredWidth}超过400，按比例缩放: 宽度=400, 高度={chartShape.Height} (原高度={configuredHeight}, 缩放比例={scaleRatio:F2})");
                }
                else
                {
                    chartShape.Width = configuredWidth;
                    chartShape.Height = configuredHeight;
                }

                // 设置图表主题
                if (!string.IsNullOrEmpty(config.Chart.Theme))
                {
                    try
                    {
                        chart.ChartStyle = ChartEngine.GetChartStyle(config.Chart.Theme);
                    }
                    catch (Exception ex)
                    {
                        string errorMsg = $"设置主题失败: {ex.Message}";
                        errors.Add(errorMsg);
                        System.Diagnostics.Debug.WriteLine($"[DEBUG] {errorMsg}");
                    }
                }

                // 填充数据
                var fillDataResult = ChartEngine.FillChartData(chart, config.Chart);
                if (!fillDataResult.Success)
                {
                    errors.Add($"填充图表数据失败: {fillDataResult.Error}");
                }

                // 写数据可能再次弹出 Excel，立即隐藏
                ChartComUiHelper.TryHideEmbeddedExcel(chart);

                // 设置标题（可选）
                if (!string.IsNullOrEmpty(config.Chart.Title))
                {
                    try
                    {
                        chart.HasTitle = true;
                        chart.ChartTitle.Text = config.Chart.Title;
                        ChartEngine.ApplyChartTitleFormat(chart, config.Chart, errors);
                    }
                    catch (Exception ex)
                    {
                        string errorMsg = $"设置标题失败: {ex.Message}";
                        errors.Add(errorMsg);
                        System.Diagnostics.Debug.WriteLine($"[DEBUG] {errorMsg}");
                    }
                }
                else
                {
                    // 如果没有标题，明确隐藏标题
                    try
                    {
                        chart.HasTitle = false;
                        System.Diagnostics.Debug.WriteLine("[DEBUG] 未提供标题，已隐藏图表标题");
                    }
                    catch (Exception ex)
                    {
                        string errorMsg = $"隐藏标题失败: {ex.Message}";
                        errors.Add(errorMsg);
                        System.Diagnostics.Debug.WriteLine($"[DEBUG] {errorMsg}");
                    }
                }

                // 设置图例位置（默认放在顶部）
                try
                {
                    chart.HasLegend = true;
                    string legendPosition = !string.IsNullOrEmpty(config.Chart.LegendPosition) 
                        ? config.Chart.LegendPosition 
                        : "top"; // 默认图例位置为顶部
                    chart.Legend.Position = (Word.XlLegendPosition)ChartEngine.GetLegendPosition(legendPosition);
                    
                    // 设置图例字体大小（默认8号）
                    try
                    {
                        chart.Legend.Font.Size = 8;
                        System.Diagnostics.Debug.WriteLine("[DEBUG] 已设置图例字体大小为8号");
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"[DEBUG] 设置图例字体大小失败: {ex.Message}");
                    }
                }
                catch (Exception ex)
                {
                    string errorMsg = $"设置图例位置失败: {ex.Message}";
                    errors.Add(errorMsg);
                    System.Diagnostics.Debug.WriteLine($"[DEBUG] {errorMsg}");
                }

                // 设置网格线（默认显示虚线主要网格线）
                if (!string.IsNullOrEmpty(config.Chart.Gridlines))
                {
                    var gridlinesResult = ChartEngine.SetGridlines(chart, config.Chart.Gridlines);
                    if (!gridlinesResult.Success)
                    {
                        errors.Add($"设置网格线失败: {gridlinesResult.Error}");
                    }
                }
                else
                {
                    // 默认显示虚线主要网格线
                    try
                    {
                        Word.Axis valueAxis = chart.Axes(Word.XlAxisType.xlValue);
                        valueAxis.HasMajorGridlines = true;
                        valueAxis.HasMinorGridlines = false;
                        valueAxis.MajorGridlines.Format.Line.DashStyle = Microsoft.Office.Core.MsoLineDashStyle.msoLineDash;
                        System.Diagnostics.Debug.WriteLine("[DEBUG] 已设置默认虚线主要网格线");
                    }
                    catch (Exception ex)
                    {
                        string errorMsg = $"设置默认网格线失败: {ex.Message}";
                        errors.Add(errorMsg);
                        System.Diagnostics.Debug.WriteLine($"[DEBUG] {errorMsg}");
                    }
                }

                // 根据图表类型应用特定格式
                var formattingResult = ChartEngine.ApplyChartSpecificFormatting(chart, config.Chart);
                if (!formattingResult.Success)
                {
                    errors.Add($"应用图表格式失败: {formattingResult.Error}");
                }

                // 最后再次确保标题状态正确（防止格式操作重新启用标题）
                try
                {
                    if (string.IsNullOrEmpty(config.Chart.Title))
                    {
                        chart.HasTitle = false;
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"[DEBUG] 最终检查标题状态失败: {ex.Message}");
                }

                System.Diagnostics.Debug.WriteLine($"[DEBUG] 图表创建成功: {config.Chart.Type}");

                try
                {
                    PostModifyNavigateHelper.NavigateAfterInsertObject(
                        document.Application, chartShape.Range, "create_chart");
                }
                catch (Exception ex)
                {
                    string errorMsg = $"移动光标失败: {ex.Message}";
                    errors.Add(errorMsg);
                    System.Diagnostics.Debug.WriteLine($"[DEBUG] {errorMsg}");
                }

                if (errors.Count > 0)
                {
                    return new ChartCreateResult
                    {
                        Success = true,
                        Error = $"图表已创建，但存在以下警告: {string.Join("; ", errors)}",
                        ChartStart = chartShape.Range.Start
                    };
                }

                return new ChartCreateResult { Success = true, Error = null, ChartStart = chartShape.Range.Start };
            }
            catch (Exception ex)
            {
                string errorMsg = $"图表创建异常: {ex.Message}";
                if (ex.InnerException != null)
                {
                    errorMsg += $"\n详细信息: {ex.InnerException.Message}";
                }
                System.Diagnostics.Debug.WriteLine($"[DEBUG] {errorMsg}");
                return new ChartCreateResult { Success = false, Error = errorMsg };
            }
        }
    }
}
