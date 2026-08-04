using System;
using System.Collections.Generic;
using System.Linq;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    public static class DataProvenanceWriteHelper
    {
        public static void WriteObjectLines(Word.Document doc, RenderedProvenanceItem item)
        {
            if (doc == null || item == null || item.RenderedObjectLines == null || item.RenderedObjectLines.Count == 0)
            {
                return;
            }

            string code = item.Anchor?.Code;
            Word.Range insert;
            if (item.Kind == ProvenanceAnchorKind.Table)
            {
                insert = ObjectInsertRangeHelper.FindRangeAfterTable(doc, code);
            }
            else if (item.Kind == ProvenanceAnchorKind.Chart)
            {
                insert = ObjectInsertRangeHelper.FindRangeAfterChart(doc, code);
            }
            else if (item.Kind == ProvenanceAnchorKind.Image)
            {
                insert = ObjectInsertRangeHelper.FindRangeAfterImage(doc, code);
            }
            else
            {
                throw new InvalidOperationException($"WriteObjectLines 不支持锚点类型 {item.Kind}");
            }

            if (insert == null)
            {
                throw new InvalidOperationException($"无法定位对象 {code}");
            }

            Word.Range firstLineRange = null;
            Word.Range lastLineRange = null;

            for (int k = 0; k < item.RenderedObjectLines.Count; k++)
            {
                if (k > 0)
                {
                    insert.InsertParagraphAfter();
                    insert.Collapse(Word.WdCollapseDirection.wdCollapseStart);
                }

                insert.Text = $"* {item.RenderedObjectLines[k]}";
                ProvenanceFormatting.ApplyStandaloneReferenceStyle(insert);
                // 若插入点误落在后续正文段内，强制在溯源后断段，避免粘行
                EnsureParagraphBreakAfter(insert);

                string lineBookmark;
                if (item.Kind == ProvenanceAnchorKind.Table)
                {
                    lineBookmark = ProvenanceBookmarkNames.TableLine(code, k + 1);
                }
                else if (item.Kind == ProvenanceAnchorKind.Chart)
                {
                    lineBookmark = ProvenanceBookmarkNames.ChartLine(code, k + 1);
                }
                else
                {
                    lineBookmark = ProvenanceBookmarkNames.ImageLine(code, k + 1);
                }

                doc.Bookmarks.Add(lineBookmark, insert.Duplicate);

                if (k == 0)
                {
                    firstLineRange = insert.Duplicate;
                }

                lastLineRange = insert.Duplicate;
            }

            if (firstLineRange != null && lastLineRange != null)
            {
                Word.Range blockRange = doc.Range(firstLineRange.Start, lastLineRange.End);
                string blockName;
                if (item.Kind == ProvenanceAnchorKind.Table)
                {
                    blockName = ProvenanceBookmarkNames.TableBlock(code);
                }
                else if (item.Kind == ProvenanceAnchorKind.Chart)
                {
                    blockName = ProvenanceBookmarkNames.ChartBlock(code);
                }
                else
                {
                    blockName = ProvenanceBookmarkNames.ImageBlock(code);
                }

                doc.Bookmarks.Add(blockName, blockRange);
            }
        }

        /// <summary>
        /// 若溯源文本仍与后续正文同段，在其后强制插入段标，避免粘行。
        /// </summary>
        private static void EnsureParagraphBreakAfter(Word.Range lineRange)
        {
            if (lineRange == null)
            {
                return;
            }

            try
            {
                Word.Paragraph para = lineRange.Paragraphs[1];
                if (para == null)
                {
                    return;
                }

                // 段末 ¶ 前仍有正文 → 溯源未独占本段
                if (lineRange.End < para.Range.End - 1)
                {
                    Word.Range splitAt = lineRange.Duplicate;
                    splitAt.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                    splitAt.InsertParagraphAfter();
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine(
                    $"[ProvenanceWrite] EnsureParagraphBreakAfter 失败: {ex.Message}");
            }
        }

        public static void WriteSentenceMarker(Word.Document doc, RenderedProvenanceItem item, int sentenceNumber)
        {
            if (doc == null || item == null)
            {
                return;
            }

            string code = item.Anchor?.Code;

            Word.Range sentence = DataProvenanceSentenceLocator.LocateLastSentenceRange(
                doc,
                item.Anchor,
                debugTag: $"provenance_write:{code}");

            sentence.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
            sentence.Text = $"[{sentenceNumber}]";
            ProvenanceFormatting.ApplyMarkerStyle(sentence);
            doc.Bookmarks.Add(ProvenanceBookmarkNames.Marker(sentenceNumber), sentence.Duplicate);
        }

        public static void WriteReferenceSection(Word.Document doc, IList<RenderedProvenanceItem> sentenceItems)
        {
            if (doc == null || sentenceItems == null || sentenceItems.Count == 0)
            {
                return;
            }

            Word.Range endRange = doc.Content;
            endRange.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
            endRange.InsertParagraphBefore();
            endRange.Collapse(Word.WdCollapseDirection.wdCollapseStart);

            endRange.Text = "参考数据来源\r";
            doc.Bookmarks.Add(ProvenanceBookmarkNames.RefTitle(), endRange.Duplicate);

            Word.Range titleRange = endRange.Duplicate;
            Word.Range lastRefRange = titleRange;

            var ordered = sentenceItems
                .Where(i => i.SentenceNumber.HasValue)
                .OrderBy(i => i.SentenceNumber.Value)
                .ToList();

            foreach (RenderedProvenanceItem item in ordered)
            {
                int n = item.SentenceNumber.Value;
                Word.Range lineRange = doc.Range(lastRefRange.End, lastRefRange.End);
                lineRange.InsertParagraphBefore();
                lineRange.Collapse(Word.WdCollapseDirection.wdCollapseStart);

                lineRange.Text = $"[{n}] {item.RenderedLineForReference}\r";
                ProvenanceFormatting.ApplyStandaloneReferenceStyle(lineRange);
                doc.Bookmarks.Add(ProvenanceBookmarkNames.RefLine(n), lineRange.Duplicate);
                lastRefRange = lineRange.Duplicate;
            }

            Word.Range blockRange = doc.Range(titleRange.Start, lastRefRange.End);
            doc.Bookmarks.Add(ProvenanceBookmarkNames.RefBlock(), blockRange);
        }

        public static List<RenderedProvenanceItem> SortSentenceItemsForMarkerWrite(
            Word.Document doc,
            IList<RenderedProvenanceItem> sentenceItems)
        {
            var withStart = new List<(RenderedProvenanceItem Item, int Start)>();
            foreach (RenderedProvenanceItem item in sentenceItems)
            {
                string code = item.Anchor?.Code;
                Word.Range range = DataProvenanceSentenceLocator.TryLocateLastSentenceRange(
                    doc,
                    item.Anchor,
                    debugTag: $"provenance_sort:{code}",
                    out Word.Range located,
                    out _)
                    ? located
                    : null;
                int start = range?.Start ?? 0;
                withStart.Add((item, start));
            }

            return withStart
                .OrderByDescending(x => x.Start)
                .Select(x => x.Item)
                .ToList();
        }
    }
}
