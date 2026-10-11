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
                WordAddIn1.DocumentHost.HostProbe.Log(
                    "tbl-after-skip",
                    "doc=" + (doc == null ? "null" : "ok") + " tableId=" + (tableId ?? "(null)"));
                return null;
            }

            try
            {
                Word.Table table = SentenceCodeLocator.ResolveTableById(doc, tableId);
                if (table == null)
                {
                    WordAddIn1.DocumentHost.HostProbe.Log("tbl-after-no-table", "id=" + tableId);
                    return null;
                }

                Word.Range insertRange = table.Range;
                int rawStart = insertRange.Start;
                int rawEnd = insertRange.End;
                WordAddIn1.DocumentHost.HostProbe.Log(
                    "tbl-after-range",
                    "id=" + tableId + " start=" + rawStart + " end=" + rawEnd);
                insertRange.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                WordAddIn1.DocumentHost.HostProbe.Log(
                    "tbl-after-collapsed",
                    "id=" + tableId + " start=" + insertRange.Start + " end=" + insertRange.End);
                insertRange.InsertParagraphAfter();
                WordAddIn1.DocumentHost.HostProbe.Log(
                    "tbl-after-para",
                    "id=" + tableId + " start=" + insertRange.Start + " end=" + insertRange.End);
                insertRange.Collapse(Word.WdCollapseDirection.wdCollapseStart);
                WordAddIn1.DocumentHost.HostProbe.Log(
                    "tbl-after-ok",
                    "id=" + tableId + " start=" + insertRange.Start);
                return insertRange;
            }
            catch (Exception ex)
            {
                WordAddIn1.DocumentHost.HostProbe.Log(
                    "tbl-after-ex",
                    "id=" + tableId + " " + WordAddIn1.DocumentHost.HostProbe.Ex(ex));
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
