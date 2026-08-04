using System;
using System.Collections.Generic;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    public sealed class ChartExtractSelectionResult
    {
        public bool Success { get; set; }
        public string Error { get; set; }
        public string ChartId { get; set; }
        public int ChartIndex { get; set; }
        public string SelectionType { get; set; }
    }

    /// <summary>
    /// 图表 extract/apply/delete 选图（仅 by_id，对齐 TableExtractSelectionHelper）。
    /// </summary>
    public static class ChartSelectionHelper
    {
        public static ChartExtractSelectionResult Resolve(Word.Document document, Dictionary<string, object> chartSelection)
        {
            var result = new ChartExtractSelectionResult();
            if (document == null)
            {
                result.Error = "没有活动的Word文档";
                return result;
            }

            if (chartSelection == null)
            {
                result.Error = "chart_selection必须是对象格式";
                return result;
            }

            string selectionType = chartSelection.ContainsKey("type")
                ? chartSelection["type"]?.ToString()
                : "";

            if (selectionType == "by_id")
            {
                return ResolveById(document, chartSelection, selectionType);
            }

            if (selectionType == "by_context")
            {
                result.Error =
                    "选择类型 by_context 暂不支持，请使用 by_id 并提供 chart_id（可先 get_document_content 获取 C_ 编号）";
                return result;
            }

            result.Error = $"不支持的选择类型：{selectionType}，请使用 by_id";
            return result;
        }

        private static ChartExtractSelectionResult ResolveById(
            Word.Document document,
            Dictionary<string, object> chartSelection,
            string selectionType)
        {
            var result = new ChartExtractSelectionResult { SelectionType = selectionType };

            if (!chartSelection.ContainsKey("chart_id"))
            {
                result.Error = "选择类型为by_id时，必须提供chart_id参数";
                return result;
            }

            string chartId = chartSelection["chart_id"]?.ToString();
            if (string.IsNullOrEmpty(chartId))
            {
                result.Error = "chart_id不能为空";
                return result;
            }

            int chartOrderIndex = DocumentState.GetChartIndex(chartId);
            if (chartOrderIndex < 0)
            {
                result.Error = $"找不到图表编号 '{chartId}'，请确认图表编号是否正确";
                return result;
            }

            List<Word.InlineShape> allCharts = ChartResolveHelper.GetInlineChartsInOrder(document);
            if (chartOrderIndex >= allCharts.Count)
            {
                result.Error =
                    $"图表编号 '{chartId}' 对应的索引{chartOrderIndex}无效，文档中只有{allCharts.Count}个图表";
                return result;
            }

            result.ChartId = chartId;
            result.ChartIndex = chartOrderIndex;
            result.Success = true;
            return result;
        }
    }
}
