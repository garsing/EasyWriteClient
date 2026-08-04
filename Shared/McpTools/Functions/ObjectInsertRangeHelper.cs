using System;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    /// <summary>
    /// 表 / 图对象之后的段落插入点（从 F_ProcessDocumentActionsTool 抽出，供 process_actions 与数据溯源共用）。
    /// </summary>
    public static class ObjectInsertRangeHelper
    {
        /// <summary>表后：Collapse(End) → InsertParagraphAfter → Collapse(Start) 新段起点。</summary>
        public static Word.Range FindRangeAfterTable(Word.Document doc, string tableId)
        {
            if (doc == null || string.IsNullOrEmpty(tableId))
            {
                return null;
            }

            try
            {
                Word.Table table = SentenceCodeLocator.ResolveTableById(doc, tableId);
                if (table == null)
                {
                    System.Diagnostics.Debug.WriteLine($"[ObjectInsertRange] 找不到表格 '{tableId}'");
                    return null;
                }

                Word.Range insertRange = table.Range;
                insertRange.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                insertRange.InsertParagraphAfter();
                // InsertParagraphAfter 后 Range 为新段；须 Collapse Start。
                // Collapse End 会落到段标之后（下一段开头），写入时易粘到后续正文。
                insertRange.Collapse(Word.WdCollapseDirection.wdCollapseStart);

                System.Diagnostics.Debug.WriteLine(
                    $"[ObjectInsertRange] 表 '{tableId}' 后插入点 Start={insertRange.Start}");
                return insertRange;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[ObjectInsertRange] 表后定位失败: {ex.Message}");
                return null;
            }
        }

        /// <summary>图后：Collapse(End) → InsertParagraphAfter → Collapse(Start) 新段起点。</summary>
        public static Word.Range FindRangeAfterChart(Word.Document doc, string chartId)
        {
            if (doc == null || string.IsNullOrEmpty(chartId))
            {
                return null;
            }

            try
            {
                Word.InlineShape chart = ChartResolveHelper.ResolveInlineShapeByChartId(doc, chartId);
                if (chart == null)
                {
                    System.Diagnostics.Debug.WriteLine($"[ObjectInsertRange] 找不到图表 '{chartId}'");
                    return null;
                }

                Word.Range insertRange = chart.Range;
                insertRange.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                insertRange.InsertParagraphAfter();
                insertRange.Collapse(Word.WdCollapseDirection.wdCollapseStart);

                System.Diagnostics.Debug.WriteLine(
                    $"[ObjectInsertRange] 图 '{chartId}' 后插入点 Start={insertRange.Start}");
                return insertRange;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[ObjectInsertRange] 图后定位失败: {ex.Message}");
                return null;
            }
        }

        /// <summary>预检用：只解析表是否存在，不插入段落、不改文档。</summary>
        public static bool ObjectExistsAfterTable(Word.Document doc, string tableId)
        {
            if (doc == null || string.IsNullOrEmpty(tableId))
            {
                return false;
            }

            try
            {
                return SentenceCodeLocator.ResolveTableById(doc, tableId) != null;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>预检用：只解析图是否存在，不插入段落、不改文档。</summary>
        public static bool ObjectExistsAfterChart(Word.Document doc, string chartId)
        {
            if (doc == null || string.IsNullOrEmpty(chartId))
            {
                return false;
            }

            try
            {
                return ChartResolveHelper.ResolveInlineShapeByChartId(doc, chartId) != null;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>图片后：与图后同逻辑（InsertParagraphAfter → Collapse Start）。</summary>
        public static Word.Range FindRangeAfterImage(Word.Document doc, string imageId)
        {
            if (doc == null || string.IsNullOrEmpty(imageId))
            {
                return null;
            }

            try
            {
                Word.InlineShape image = ImageResolveHelper.ResolveInlineShapeByImageId(doc, imageId);
                if (image == null)
                {
                    System.Diagnostics.Debug.WriteLine($"[ObjectInsertRange] 找不到图片 '{imageId}'");
                    return null;
                }

                Word.Range insertRange = image.Range;
                insertRange.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                insertRange.InsertParagraphAfter();
                insertRange.Collapse(Word.WdCollapseDirection.wdCollapseStart);

                System.Diagnostics.Debug.WriteLine(
                    $"[ObjectInsertRange] 图片 '{imageId}' 后插入点 Start={insertRange.Start}");
                return insertRange;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[ObjectInsertRange] 图片后定位失败: {ex.Message}");
                return null;
            }
        }

        /// <summary>预检用：只解析图片是否存在，不插入段落、不改文档。</summary>
        public static bool ObjectExistsAfterImage(Word.Document doc, string imageId)
        {
            if (doc == null || string.IsNullOrEmpty(imageId))
            {
                return false;
            }

            try
            {
                return ImageResolveHelper.ResolveInlineShapeByImageId(doc, imageId) != null;
            }
            catch
            {
                return false;
            }
        }
    }
}
