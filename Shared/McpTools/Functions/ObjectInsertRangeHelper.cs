using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    /// <summary>
    /// 表 / 图对象之后的段落插入点（从 F_ProcessDocumentActionsTool 抽出，供 process_actions 与数据溯源共用）。
    /// </summary>
    public static class ObjectInsertRangeHelper
    {
        /// <summary>
        /// 表后新段起点。先走 Collapse(End) → InsertParagraphAfter（Word）；
        /// WPS 在表 Range 上该调用常 E_FAIL，改用文档坐标走到表外再开段。
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
                int objectStart = tableRange.Start;
                int objectEnd = tableRange.End;
                int lastCellEnd = SafeLastCellEnd(table, objectEnd);
                WordAddIn1.DocumentHost.HostProbe.Log(
                    "tbl-after-range",
                    "id=" + tableId + " start=" + objectStart + " end=" + objectEnd
                    + " lastCellEnd=" + lastCellEnd + " docEnd=" + SafeContentEnd(doc));

                Word.Range wordPath = TryInsertParagraphAfterCollapsedEnd(tableRange, tableId);
                if (wordPath != null && !RangeIsInCell(wordPath))
                {
                    WordAddIn1.DocumentHost.HostProbe.Log(
                        "tbl-after-ok",
                        "id=" + tableId + " via=insert-para-after start=" + wordPath.Start);
                    return wordPath;
                }

                Word.Range fallback = OpenParagraphAfterTable(doc, table, objectStart, lastCellEnd, tableId);
                if (fallback == null)
                {
                    WordAddIn1.DocumentHost.HostProbe.Log("tbl-after-fallback-null", "id=" + tableId);
                }

                return fallback;
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

        private static Word.Range OpenParagraphAfterTable(
            Word.Document doc,
            Word.Table table,
            int objectStart,
            int objectEnd,
            string tableId)
        {
            int after = FirstPositionAfterTable(doc, table, objectStart, objectEnd);
            WordAddIn1.DocumentHost.HostProbe.Log(
                "tbl-after-walk",
                "id=" + tableId + " objectEnd=" + objectEnd + " after=" + after
                + " docEnd=" + SafeContentEnd(doc)
                + " " + ProbeInTable(doc, after));

            if (NeedsGrowAfterTable(doc, table, objectStart, after))
            {
                if (!TryGrowDocumentAfter(doc, table, objectEnd, tableId))
                {
                    return null;
                }

                try
                {
                    Word.Table live = SentenceCodeLocator.ResolveTableById(doc, tableId)
                        ?? FirstTableNear(doc, objectStart);
                    if (live != null)
                    {
                        table = live;
                    }

                    objectStart = table.Range.Start;
                    objectEnd = SafeLastCellEnd(table, table.Range.End);
                }
                catch (Exception ex)
                {
                    WordAddIn1.DocumentHost.HostProbe.Log(
                        "tbl-after-reread-ex",
                        "id=" + tableId + " " + WordAddIn1.DocumentHost.HostProbe.Ex(ex));
                }

                after = FirstPositionAfterTable(doc, table, objectStart, objectEnd);
                WordAddIn1.DocumentHost.HostProbe.Log(
                    "tbl-after-grew",
                    "id=" + tableId + " after=" + after + " start=" + objectStart
                    + " end=" + objectEnd + " docEnd=" + SafeContentEnd(doc)
                    + " " + ProbeInTable(doc, after)
                    + " " + ProbeInTable(doc, SafeContentEnd(doc) > 0 ? SafeContentEnd(doc) - 1 : 0));
            }

            Word.Range at = RangeOutsideCell(doc, after);
            if (at == null)
            {
                at = FindParagraphAfterLastCell(doc, objectStart, objectEnd);
            }

            if (at == null)
            {
                WordAddIn1.DocumentHost.HostProbe.Log(
                    "tbl-after-still-in",
                    "id=" + tableId + " after=" + after);
                return null;
            }

            if (!ParagraphIsBlank(at) && !TryOpenBreakAt(at, tableId))
            {
                WordAddIn1.DocumentHost.HostProbe.Log(
                    "tbl-after-break-fail",
                    "id=" + tableId + " start=" + at.Start);
            }

            at.Collapse(Word.WdCollapseDirection.wdCollapseStart);
            if (RangeIsInCell(at))
            {
                WordAddIn1.DocumentHost.HostProbe.Log(
                    "tbl-after-land-in",
                    "id=" + tableId + " start=" + at.Start);
                return null;
            }

            WordAddIn1.DocumentHost.HostProbe.Log(
                "tbl-after-ok",
                "id=" + tableId + " via=coord start=" + at.Start);
            return at;
        }

        private static bool NeedsGrowAfterTable(Word.Document doc, Word.Table table, int objectStart, int after)
        {
            int docEnd = SafeContentEnd(doc);
            if (after >= docEnd)
            {
                return true;
            }

            Word.Range probe = SafeDocRange(doc, after, after);
            return probe == null || RangeIsInCell(probe);
        }

        private static int FirstPositionAfterTable(
            Word.Document doc,
            Word.Table table,
            int objectStart,
            int objectEnd)
        {
            int docEnd = SafeContentEnd(doc);
            int pos = objectEnd;
            if (pos < 0)
            {
                pos = 0;
            }

            if (pos > docEnd)
            {
                pos = docEnd;
            }

            int guard = 0;
            while (pos < docEnd && guard < 64)
            {
                Word.Range probe = SafeDocRange(doc, pos, pos);
                if (probe == null || !RangeIsInCell(probe))
                {
                    break;
                }

                pos++;
                guard++;
            }

            return pos;
        }

        private static Word.Range RangeOutsideCell(Word.Document doc, int after)
        {
            int docEnd = SafeContentEnd(doc);
            int[] candidates = after >= docEnd && docEnd > 0
                ? new[] { docEnd - 1, after }
                : new[] { after, docEnd > 0 ? docEnd - 1 : after };

            for (int i = 0; i < candidates.Length; i++)
            {
                Word.Range probe = SafeDocRange(doc, candidates[i], candidates[i]);
                if (probe == null || RangeIsInCell(probe))
                {
                    continue;
                }

                probe.Collapse(Word.WdCollapseDirection.wdCollapseStart);
                return probe;
            }

            return null;
        }

        private static Word.Range FindParagraphAfterLastCell(Word.Document doc, int tableStart, int lastCellEnd)
        {
            Word.Range best = null;
            int bestStart = -1;
            try
            {
                foreach (Word.Paragraph paragraph in doc.Paragraphs)
                {
                    Word.Range pr = paragraph.Range;
                    if (RangeIsInCell(pr))
                    {
                        continue;
                    }

                    if (pr.Start < lastCellEnd && pr.End <= lastCellEnd)
                    {
                        continue;
                    }

                    if (pr.Start >= bestStart)
                    {
                        best = pr;
                        bestStart = pr.Start;
                    }
                }
            }
            catch (Exception ex)
            {
                WordAddIn1.DocumentHost.HostProbe.Log(
                    "tbl-after-para-scan-ex",
                    WordAddIn1.DocumentHost.HostProbe.Ex(ex) + " tableStart=" + tableStart);
            }

            if (best == null)
            {
                return null;
            }

            best.Collapse(Word.WdCollapseDirection.wdCollapseStart);
            return best;
        }

        private static bool HasParagraphAfterLastCell(Word.Document doc, string tableId, int hintStart)
        {
            try
            {
                Word.Table live = SentenceCodeLocator.ResolveTableById(doc, tableId)
                    ?? FirstTableNear(doc, hintStart);
                if (live == null)
                {
                    return false;
                }

                int last = SafeLastCellEnd(live, live.Range.End);
                return FindParagraphAfterLastCell(doc, live.Range.Start, last) != null
                    || RangeOutsideCell(doc, last) != null;
            }
            catch (Exception)
            {
                return false;
            }
        }

        private static Word.Table FirstTableNear(Word.Document doc, int hintStart)
        {
            try
            {
                Word.Table best = null;
                int bestDelta = int.MaxValue;
                foreach (Word.Table candidate in doc.Tables)
                {
                    int delta = Math.Abs(candidate.Range.Start - hintStart);
                    if (delta < bestDelta)
                    {
                        best = candidate;
                        bestDelta = delta;
                    }
                }

                return best;
            }
            catch (Exception)
            {
                return null;
            }
        }

        private static int SafeLastCellEnd(Word.Table table, int fallback)
        {
            try
            {
                Word.Cell last = table.Cell(table.Rows.Count, table.Columns.Count);
                return last.Range.End;
            }
            catch (Exception)
            {
                return fallback;
            }
        }

        private static bool TryGrowDocumentAfter(Word.Document doc, Word.Table table, int hintEnd, string tableId)
        {
            int objectStart = table.Range.Start;
            if (TryInjectParagraphAfterTableXml(doc, table, tableId)
                && HasPositionOutsideCell(doc))
            {
                return true;
            }

            try
            {
                Word.Paragraph added = doc.Paragraphs.Add();
                if (added?.Range != null
                    && !RangeIsInCell(added.Range)
                    && HasPositionOutsideCell(doc))
                {
                    WordAddIn1.DocumentHost.HostProbe.Log(
                        "tbl-after-grow",
                        "id=" + tableId + " via=paragraphs-add start=" + added.Range.Start);
                    return true;
                }

                WordAddIn1.DocumentHost.HostProbe.Log(
                    "tbl-after-grow-in",
                    "id=" + tableId + " via=paragraphs-add");
            }
            catch (Exception ex)
            {
                WordAddIn1.DocumentHost.HostProbe.Log(
                    "tbl-after-grow-ex",
                    "id=" + tableId + " paragraphs-add " + WordAddIn1.DocumentHost.HostProbe.Ex(ex));
            }

            WordAddIn1.DocumentHost.HostProbe.Log(
                "tbl-after-grow-fail",
                "id=" + tableId + " hintEnd=" + hintEnd);
            return false;
        }

        private static bool TryInjectParagraphAfterTableXml(Word.Document doc, Word.Table table, string tableId)
        {
            try
            {
                string stamp = "";
                try
                {
                    stamp = table.Cell(table.Rows.Count, table.Columns.Count).Range.Text ?? "";
                }
                catch (Exception)
                {
                }

                string raw = doc.Content.WordOpenXML;
                if (string.IsNullOrEmpty(raw))
                {
                    WordAddIn1.DocumentHost.HostProbe.Log("tbl-after-xml-empty", "id=" + tableId);
                    return false;
                }

                const string marker = "</w:tbl>";
                int last = raw.LastIndexOf(marker, StringComparison.Ordinal);
                if (last < 0)
                {
                    WordAddIn1.DocumentHost.HostProbe.Log("tbl-after-xml-no-tbl", "id=" + tableId);
                    return false;
                }

                int insertAt = last + marker.Length;
                string rest = raw.Substring(insertAt, Math.Min(64, raw.Length - insertAt)).TrimStart();
                WordAddIn1.DocumentHost.HostProbe.Log(
                    "tbl-after-body",
                    "id=" + tableId + " afterTbl=" + rest.Substring(0, Math.Min(40, rest.Length)).Replace('\r', ' ')
                    + " stampLen=" + stamp.Length);
                if (rest.StartsWith("<w:p", StringComparison.Ordinal))
                {
                    WordAddIn1.DocumentHost.HostProbe.Log("tbl-after-xml-has-p", "id=" + tableId);
                    return false;
                }

                string injected = raw.Insert(
                    insertAt,
                    "<w:p xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"/>");
                doc.Content.InsertXML(injected);

                string stampAfter = "";
                try
                {
                    Word.Table live = SentenceCodeLocator.ResolveTableById(doc, tableId)
                        ?? FirstTableNear(doc, table.Range.Start);
                    if (live != null)
                    {
                        stampAfter = live.Cell(live.Rows.Count, live.Columns.Count).Range.Text ?? "";
                    }
                }
                catch (Exception)
                {
                }

                WordAddIn1.DocumentHost.HostProbe.Log(
                    "tbl-after-grow",
                    "id=" + tableId + " via=ooxml-raw stampAfterLen=" + stampAfter.Length
                    + " stampKept=" + (stampAfter.IndexOf(stamp.TrimEnd('\r', '\a', '\n'), StringComparison.Ordinal) >= 0));
                return true;
            }
            catch (Exception ex)
            {
                WordAddIn1.DocumentHost.HostProbe.Log(
                    "tbl-after-ooxml-ex",
                    "id=" + tableId + " " + WordAddIn1.DocumentHost.HostProbe.Ex(ex));
                return false;
            }
        }

        private static XElement FindBodyTableElement(OpenXmlPackage pkg, Word.Document doc, Word.Table table)
        {
            List<XElement> xmlTables = pkg.Body.Elements(OpenXmlPackage.W + "tbl").ToList();
            if (xmlTables.Count == 0)
            {
                return null;
            }

            int wanted = table.Range.Start;
            int index = 0;
            foreach (Word.Table candidate in doc.Tables)
            {
                if (IsNestedTable(candidate))
                {
                    continue;
                }

                if (candidate.Range.Start == wanted)
                {
                    return index < xmlTables.Count ? xmlTables[index] : xmlTables[xmlTables.Count - 1];
                }

                index++;
            }

            return xmlTables[xmlTables.Count - 1];
        }

        private static bool IsNestedTable(Word.Table table)
        {
            try
            {
                int start = table.Range.Start;
                foreach (Word.Table other in table.Range.Document.Tables)
                {
                    if (object.ReferenceEquals(other, table))
                    {
                        continue;
                    }

                    if (other.Range.Start < start && other.Range.End > start)
                    {
                        return true;
                    }
                }
            }
            catch (Exception)
            {
            }

            return false;
        }

        private static bool HasPositionOutsideCell(Word.Document doc)
        {
            try
            {
                foreach (Word.Paragraph paragraph in doc.Paragraphs)
                {
                    if (!RangeIsInCell(paragraph.Range))
                    {
                        return true;
                    }
                }
            }
            catch (Exception)
            {
            }

            int docEnd = SafeContentEnd(doc);
            int begin = docEnd > 8 ? docEnd - 8 : 0;
            for (int pos = begin; pos < docEnd; pos++)
            {
                Word.Range probe = SafeDocRange(doc, pos, pos);
                if (probe != null && !RangeIsInCell(probe))
                {
                    return true;
                }
            }

            return false;
        }

        private static bool TryOpenBreakAt(Word.Range at, string tableId)
        {
            try
            {
                at.InsertParagraphBefore();
                WordAddIn1.DocumentHost.HostProbe.Log(
                    "tbl-after-break",
                    "id=" + tableId + " via=insert-para-before");
                return true;
            }
            catch (Exception)
            {
            }

            try
            {
                at.InsertAfter("\r");
                WordAddIn1.DocumentHost.HostProbe.Log(
                    "tbl-after-break",
                    "id=" + tableId + " via=insert-after-cr");
                return true;
            }
            catch (Exception)
            {
            }

            try
            {
                at.InsertParagraph();
                WordAddIn1.DocumentHost.HostProbe.Log(
                    "tbl-after-break",
                    "id=" + tableId + " via=insert-para");
                return true;
            }
            catch (Exception ex)
            {
                WordAddIn1.DocumentHost.HostProbe.Log(
                    "tbl-after-break-ex",
                    "id=" + tableId + " " + WordAddIn1.DocumentHost.HostProbe.Ex(ex));
                return false;
            }
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

        private static string ProbeInTable(Word.Document doc, int pos)
        {
            Word.Range probe = SafeDocRange(doc, pos, pos);
            if (probe == null)
            {
                return "pos=" + pos + " in=?";
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

            return "pos=" + pos + " inTable=" + (inTable.HasValue ? inTable.Value.ToString() : "?")
                + " cells=" + cells + " tables=" + tables;
        }

        private static bool ParagraphIsBlank(Word.Range at)
        {
            try
            {
                string text = at.Paragraphs[1].Range.Text ?? "";
                for (int i = 0; i < text.Length; i++)
                {
                    char c = text[i];
                    if (c != '\r' && c != '\a' && c != '\n' && c != '\f' && c != ' ' && c != '\t')
                    {
                        return false;
                    }
                }

                return true;
            }
            catch (Exception)
            {
                return true;
            }
        }

        private static Word.Range SafeDocRange(Word.Document doc, int start, int end)
        {
            try
            {
                int docEnd = SafeContentEnd(doc);
                if (start < 0)
                {
                    start = 0;
                }

                if (end < start)
                {
                    end = start;
                }

                if (start > docEnd)
                {
                    start = docEnd;
                }

                if (end > docEnd)
                {
                    end = docEnd;
                }

                if (start == docEnd && docEnd > 0)
                {
                    start = docEnd - 1;
                    end = start;
                }

                return doc.Range(start, end);
            }
            catch (Exception ex)
            {
                WordAddIn1.DocumentHost.HostProbe.Log(
                    "tbl-after-safe-range-ex",
                    WordAddIn1.DocumentHost.HostProbe.Ex(ex) + " start=" + start + " end=" + end);
                return null;
            }
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
