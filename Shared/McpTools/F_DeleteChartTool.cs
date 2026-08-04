using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    /// <summary>
    /// 按 chart_id 删除 InlineShape 图表；删后光标落原图后沿，供 F_create_chart_from_xml 重建。
    /// </summary>
    public static class F_DeleteChartTool
    {
        public static void Register(
            Dictionary<string, Func<Dictionary<string, object>, Task<ToolResult>>> toolRegistry,
            object wordApplication)
        {
            toolRegistry["F_delete_chart"] = async (args) =>
            {
                try
                {
                    if (wordApplication == null)
                    {
                        return new ToolResult { Success = false, Error = "Word应用程序不可用" };
                    }

                    dynamic wordApp = wordApplication;
                    if (wordApp.ActiveDocument == null)
                    {
                        return new ToolResult { Success = false, Error = "没有活动的Word文档" };
                    }

                    DocumentState.BindAndActivate(wordApp.ActiveDocument);

                    string chartId = args.ContainsKey("chart_id") ? args["chart_id"]?.ToString() : "";
                    if (string.IsNullOrWhiteSpace(chartId))
                    {
                        return new ToolResult { Success = false, Error = "未提供图表编号" };
                    }

                    Word.Document document = wordApp.ActiveDocument;

                    try
                    {
                        WordReader.ReadWord(document);
                    }
                    catch (Exception ex)
                    {
                        return new ToolResult { Success = false, Error = $"生成图表映射表失败：{ex.Message}" };
                    }

                    int chartOrderIndex = DocumentState.GetChartIndex(chartId);
                    if (chartOrderIndex < 0)
                    {
                        return new ToolResult
                        {
                            Success = false,
                            Error = $"找不到图表编号 '{chartId}'，请确认图表编号是否正确"
                        };
                    }

                    Word.InlineShape inlineShape = ChartResolveHelper.ResolveInlineShapeByChartId(document, chartId);
                    if (inlineShape == null)
                    {
                        List<Word.InlineShape> allCharts = ChartResolveHelper.GetInlineChartsInOrder(document);
                        return new ToolResult
                        {
                            Success = false,
                            Error =
                                $"图表编号 '{chartId}' 对应的索引{chartOrderIndex}无效，文档中只有{allCharts.Count}个图表"
                        };
                    }

                    int endPos = inlineShape.Range.End;
                    string chartType = TryGetChartTypeString(inlineShape.Chart);
                    int seriesCount = TryGetSeriesCount(inlineShape.Chart);

                    System.Diagnostics.Debug.WriteLine(
                        $"[DeleteChart] 即将删除 chart_id={chartId}, endPos={endPos}, " +
                        $"type={chartType}, series_count={seriesCount}");

                    try
                    {
                        inlineShape.Delete();
                    }
                    catch (Exception ex)
                    {
                        return new ToolResult { Success = false, Error = $"删除图表失败：{ex.Message}" };
                    }

                    try
                    {
                        WordReader.ReadWord(document);
                    }
                    catch (Exception ex)
                    {
                        return new ToolResult { Success = false, Error = $"删图后刷新映射失败：{ex.Message}" };
                    }

                    Word.Application app = wordApp as Word.Application;
                    if (app != null)
                    {
                        int docEnd = document.Content.End;
                        int pos = endPos < 0 ? 0 : (endPos > docEnd ? docEnd : endPos);
                        Word.Range afterChartRange = document.Range(pos, pos);
                        PostModifyNavigateHelper.NavigateAfterEnd(
                            app, afterChartRange, $"delete_chart:{chartId}");
                    }

                    string message = BuildSuccessMessage(chartId, chartType, seriesCount);

                    System.Diagnostics.Debug.WriteLine(
                        $"[DeleteChart] 成功 chart_id={chartId}, 光标落点 endPos={endPos}");

                    await Task.CompletedTask;

                    return new ToolResult
                    {
                        Success = true,
                        Data = new
                        {
                            chart_deleted = true,
                            chart_id = chartId,
                            chart_type = chartType,
                            series_count = seriesCount,
                            message = message,
                        },
                    };
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"[DeleteChart] 执行失败: {ex.Message}");
                    return new ToolResult { Success = false, Error = $"执行失败: {ex.Message}" };
                }
            };
        }

        private static string TryGetChartTypeString(Word.Chart chart)
        {
            if (chart == null)
            {
                return "";
            }

            try
            {
                int chartType = (int)chart.ChartType;
                if (chartType == (int)Excel.XlChartType.xlColumnClustered)
                {
                    return "column_clustered";
                }

                if (chartType == (int)Excel.XlChartType.xlLine)
                {
                    return "line";
                }

                if (chartType == (int)Excel.XlChartType.xl3DPie)
                {
                    return "pie3d";
                }

                return chart.ChartType.ToString().ToLowerInvariant();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[DeleteChart] 读取图表类型失败: {ex.Message}");
                return "";
            }
        }

        private static int TryGetSeriesCount(Word.Chart chart)
        {
            if (chart == null)
            {
                return 0;
            }

            try
            {
                return chart.SeriesCollection().Count;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[DeleteChart] 读取系列数失败: {ex.Message}");
                return 0;
            }
        }

        private static string BuildSuccessMessage(string chartId, string chartType, int seriesCount)
        {
            var sb = new StringBuilder();
            sb.Append($"图表 {chartId} 已删除。");
            if (!string.IsNullOrEmpty(chartType))
            {
                sb.Append($"（原类型：{chartType}");
                if (seriesCount > 0)
                {
                    sb.Append($"，{seriesCount} 个系列");
                }

                sb.Append("）");
            }

            sb.Append("光标已停留在被删图表原后沿处。");
            sb.Append("请在此位置调用 F_create_chart_from_xml 重建新图（无需传入 position 参数）。");
            return sb.ToString();
        }
    }
}
