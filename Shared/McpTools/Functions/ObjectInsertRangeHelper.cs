using System;
using System.Text;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    /// <summary>
    /// 表 / 图对象之后的段落插入点（从 F_ProcessDocumentActionsTool 抽出，供 process_actions 与数据溯源共用）。
    /// </summary>
    public static class ObjectInsertRangeHelper
    {
        /// <summary>
        /// 表后新段起点。先 Collapse(End) → InsertParagraphAfter（Word 表后已有段）。
        /// WPS 文末表 Collapse 常停在表内并 E_FAIL，再试不夹紧的 Range(表.End)。
        /// </summary>
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

                Word.Range tableRange = table.Range;
                int objectEnd = tableRange.End;
                if (WordAddIn1.DocumentHost.HostProbe.Enabled)
                {
                    WordAddIn1.DocumentHost.HostProbe.Log(
                        "tbl-after-range",
                        "id=" + tableId + " start=" + tableRange.Start + " end=" + objectEnd
                        + " docEnd=" + SafeContentEnd(doc));
                }

                Word.Range wordPath = TryInsertParagraphAfterCollapsedEnd(tableRange, tableId);
                if (wordPath != null && !RangeIsInCell(wordPath))
                {
                    WordAddIn1.DocumentHost.HostProbe.Log(
                        "tbl-after-ok",
                        "id=" + tableId + " via=insert-para-after start=" + wordPath.Start);
                    return wordPath;
                }

                Word.Range rawPath = TryRawEndInsert(doc, objectEnd, tableId);
                if (rawPath == null)
                {
                    WordAddIn1.DocumentHost.HostProbe.Log("tbl-after-fallback-null", "id=" + tableId);
                }

                return rawPath;
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
                ChartResolveHelper.BodyChart chart = ChartResolveHelper.ResolveBodyChartByChartId(doc, chartId);
                if (chart == null)
                {
                    EasyWriteDiagnostics.Log(DebugCategory.Chart, "[ObjectInsertRange] 找不到图表 '" + chartId + "'");
                    return null;
                }

                Word.Range insertRange = chart.NavigateRange();
                if (insertRange == null)
                {
                    return null;
                }
                insertRange.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                insertRange.InsertParagraphAfter();
                insertRange.Collapse(Word.WdCollapseDirection.wdCollapseStart);

                EasyWriteDiagnostics.Log(
                    DebugCategory.Chart,
                    "[ObjectInsertRange] 图 '" + chartId + "' 后插入点 Start=" + insertRange.Start);
                return insertRange;
            }
            catch (Exception ex)
            {
                EasyWriteDiagnostics.Log(DebugCategory.Chart, "[ObjectInsertRange] 图后定位失败: " + ex.Message);
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
                return ChartResolveHelper.ResolveBodyChartByChartId(doc, chartId) != null;
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
                    EasyWriteDiagnostics.Log(DebugCategory.Chart, "[ObjectInsertRange] 找不到图片 '" + imageId + "'");
                    return null;
                }

                Word.Range insertRange = image.Range;
                insertRange.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                insertRange.InsertParagraphAfter();
                insertRange.Collapse(Word.WdCollapseDirection.wdCollapseStart);

                EasyWriteDiagnostics.Log(
                    DebugCategory.Chart,
                    "[ObjectInsertRange] 图片 '" + imageId + "' 后插入点 Start=" + insertRange.Start);
                return insertRange;
            }
            catch (Exception ex)
            {
                EasyWriteDiagnostics.Log(DebugCategory.Chart, "[ObjectInsertRange] 图片后定位失败: " + ex.Message);
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

        private static Word.Range TryInsertParagraphAfterCollapsedEnd(Word.Range objectRange, string tableId)
        {
            try
            {
                Word.Range insertRange = objectRange.Duplicate;
                insertRange.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                WordAddIn1.DocumentHost.HostProbe.Log(
                    "tbl-after-collapsed",
                    "id=" + tableId + " start=" + insertRange.Start + " end=" + insertRange.End);
                insertRange.InsertParagraphAfter();
                WordAddIn1.DocumentHost.HostProbe.Log(
                    "tbl-after-para",
                    "id=" + tableId + " start=" + insertRange.Start + " end=" + insertRange.End);
                insertRange.Collapse(Word.WdCollapseDirection.wdCollapseStart);
                return insertRange;
            }
            catch (Exception ex)
            {
                WordAddIn1.DocumentHost.HostProbe.Log(
                    "tbl-after-para-fail",
                    "id=" + tableId + " " + WordAddIn1.DocumentHost.HostProbe.Ex(ex));
                return null;
            }
        }

        /// <summary>
        /// 不夹到 docEnd-1，直接试 Range(表.End) / Range(表.End+1)。
        /// WPS 表吃到文末时 Collapse 会停在表内，表.End 才是表后空锚点。
        /// </summary>
        private static Word.Range TryRawEndInsert(Word.Document doc, int objectEnd, string tableId)
        {
            int[] positions = { objectEnd, objectEnd + 1 };
            for (int i = 0; i < positions.Length; i++)
            {
                int pos = positions[i];
                Word.Range raw;
                try
                {
                    raw = doc.Range(pos, pos);
                    if (WordAddIn1.DocumentHost.HostProbe.Enabled)
                    {
                        WordAddIn1.DocumentHost.HostProbe.Log(
                            "tbl-after-raw",
                            "id=" + tableId + " asked=" + pos
                            + " got=" + raw.Start + "-" + raw.End
                            + " text=" + QuoteSpecial(raw.Text)
                            + " " + DescribeRange(raw));
                    }
                }
                catch (Exception ex)
                {
                    WordAddIn1.DocumentHost.HostProbe.Log(
                        "tbl-after-raw-ex",
                        "id=" + tableId + " pos=" + pos + " "
                        + WordAddIn1.DocumentHost.HostProbe.Ex(ex));
                    continue;
                }

                Word.Range landed = TryWriteAfterRaw(raw, tableId, pos, "insert-para-after");
                if (landed != null)
                {
                    return landed;
                }

                try
                {
                    raw = doc.Range(pos, pos);
                }
                catch (Exception)
                {
                    continue;
                }

                landed = TryWriteAfterRawInsert(raw, tableId, pos);
                if (landed != null)
                {
                    return landed;
                }
            }

            return null;
        }

        private static Word.Range TryWriteAfterRaw(
            Word.Range raw,
            string tableId,
            int pos,
            string via)
        {
            try
            {
                raw.InsertParagraphAfter();
                if (WordAddIn1.DocumentHost.HostProbe.Enabled)
                {
                    WordAddIn1.DocumentHost.HostProbe.Log(
                        "tbl-after-raw-para",
                        "id=" + tableId + " pos=" + pos
                        + " start=" + raw.Start + " end=" + raw.End
                        + " " + DescribeRange(raw));
                }
                raw.Collapse(Word.WdCollapseDirection.wdCollapseStart);
                if (RangeIsInCell(raw))
                {
                    WordAddIn1.DocumentHost.HostProbe.Log(
                        "tbl-after-raw-in",
                        "id=" + tableId + " pos=" + pos + " start=" + raw.Start);
                    return null;
                }

                WordAddIn1.DocumentHost.HostProbe.Log(
                    "tbl-after-ok",
                    "id=" + tableId + " via=" + via + " pos=" + pos + " start=" + raw.Start);
                return raw;
            }
            catch (Exception ex)
            {
                WordAddIn1.DocumentHost.HostProbe.Log(
                    "tbl-after-raw-para-fail",
                    "id=" + tableId + " pos=" + pos + " "
                    + WordAddIn1.DocumentHost.HostProbe.Ex(ex));
                return null;
            }
        }

        private static Word.Range TryWriteAfterRawInsert(Word.Range raw, string tableId, int pos)
        {
            try
            {
                raw.InsertAfter("\r");
                if (WordAddIn1.DocumentHost.HostProbe.Enabled)
                {
                    WordAddIn1.DocumentHost.HostProbe.Log(
                        "tbl-after-raw-ins",
                        "id=" + tableId + " pos=" + pos
                        + " start=" + raw.Start + " end=" + raw.End
                        + " " + DescribeRange(raw));
                }
                raw.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                if (RangeIsInCell(raw))
                {
                    WordAddIn1.DocumentHost.HostProbe.Log(
                        "tbl-after-raw-ins-in",
                        "id=" + tableId + " pos=" + pos + " start=" + raw.Start);
                    return null;
                }

                WordAddIn1.DocumentHost.HostProbe.Log(
                    "tbl-after-ok",
                    "id=" + tableId + " via=insert-after-cr pos=" + pos + " start=" + raw.Start);
                return raw;
            }
            catch (Exception ex)
            {
                WordAddIn1.DocumentHost.HostProbe.Log(
                    "tbl-after-raw-ins-fail",
                    "id=" + tableId + " pos=" + pos + " "
                    + WordAddIn1.DocumentHost.HostProbe.Ex(ex));
                return null;
            }
        }

        private static string DescribeRange(Word.Range probe)
        {
            if (probe == null)
            {
                return "in=?";
            }

            bool? inTable = TryWithInTable(probe);
            int tables = -1;
            int cells = -1;
            try
            {
                tables = probe.Tables.Count;
            }
            catch (Exception)
            {
            }

            try
            {
                cells = probe.Cells.Count;
            }
            catch (Exception)
            {
                cells = 0;
            }

            return "inTable=" + (inTable.HasValue ? inTable.Value.ToString() : "?")
                + " cells=" + cells + " tables=" + tables;
        }

        private static bool RangeIsInCell(Word.Range probe)
        {
            if (probe == null)
            {
                return false;
            }

            try
            {
                return probe.Cells != null && probe.Cells.Count > 0;
            }
            catch (Exception)
            {
                return false;
            }
        }

        private static bool? TryWithInTable(Word.Range probe)
        {
            try
            {
                object info = probe.Information[Word.WdInformation.wdWithInTable];
                if (info == null || info is DBNull)
                {
                    return null;
                }

                if (info is bool flag)
                {
                    return flag;
                }

                return Convert.ToInt32(info) != 0;
            }
            catch (Exception)
            {
                return null;
            }
        }

        private static string QuoteSpecial(string text)
        {
            if (text == null)
            {
                return "(null)";
            }

            var sb = new StringBuilder("\"");
            foreach (char ch in text)
            {
                if (ch == '\r')
                {
                    sb.Append("\\r");
                }
                else if (ch == '\n')
                {
                    sb.Append("\\n");
                }
                else if (ch == '\a')
                {
                    sb.Append("\\a");
                }
                else if (ch == '\t')
                {
                    sb.Append("\\t");
                }
                else if (ch == '\f')
                {
                    sb.Append("\\f");
                }
                else if (ch < 32)
                {
                    sb.Append("\\x").Append(((int)ch).ToString("X2"));
                }
                else
                {
                    sb.Append(ch);
                }
            }

            return sb.Append('"').ToString();
        }

        private static int SafeContentEnd(Word.Document doc)
        {
            try
            {
                return doc.Content.End;
            }
            catch (Exception)
            {
                return 0;
            }
        }
    }
}
