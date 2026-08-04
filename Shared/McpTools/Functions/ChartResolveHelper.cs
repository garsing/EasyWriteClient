using System;
using System.Collections.Generic;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    /// <summary>
    /// 图表编号与 InlineShape 定位（PR-1 起用，PR-2/3 扩展）。
    /// </summary>
    public static class ChartResolveHelper
    {
        public static List<Word.InlineShape> GetInlineChartsInOrder(Word.Document document)
        {
            var charts = new List<Word.InlineShape>();
            if (document == null)
            {
                return charts;
            }

            try
            {
                foreach (Word.InlineShape shape in document.InlineShapes)
                {
                    if (shape.Type == Word.WdInlineShapeType.wdInlineShapeChart)
                    {
                        charts.Add(shape);
                    }
                }

                charts.Sort((a, b) => a.Range.Start.CompareTo(b.Range.Start));
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[ChartResolve] 收集 InlineShape 图表失败: {ex.Message}");
            }

            return charts;
        }

        public static Word.InlineShape ResolveInlineShapeByChartId(Word.Document document, string chartId)
        {
            if (document == null || string.IsNullOrEmpty(chartId))
            {
                return null;
            }

            int chartOrderIndex = DocumentState.GetChartIndex(chartId);
            if (chartOrderIndex < 0)
            {
                return null;
            }

            List<Word.InlineShape> allCharts = GetInlineChartsInOrder(document);
            if (chartOrderIndex >= allCharts.Count)
            {
                return null;
            }

            return allCharts[chartOrderIndex];
        }
    }
}
