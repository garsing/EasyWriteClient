using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
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
            var spec = new InsertTargetLocatorHelper.LocatorSpec();
            if (item.Kind == ProvenanceAnchorKind.Table)
            {
                spec.Kind = InsertTargetLocatorHelper.LocatorKind.Table;
                spec.TableId = code;
            }
            else if (item.Kind == ProvenanceAnchorKind.Chart)
            {
                spec.Kind = InsertTargetLocatorHelper.LocatorKind.Chart;
                spec.ChartId = code;
            }
            else if (item.Kind == ProvenanceAnchorKind.Image)
            {
                spec.Kind = InsertTargetLocatorHelper.LocatorKind.Image;
                spec.ImageId = code;
            }
            else
            {
                throw new InvalidOperationException($"WriteObjectLines 不支持锚点类型 {item.Kind}");
            }

            var locate = InsertTargetLocatorHelper.TryResolveInsertRange(
                doc,
                spec,
                InsertTargetLocatorHelper.LocateRangeKind.InsertAfter);
            if (!locate.Success || locate.Range == null)
            {
                throw new InvalidOperationException(locate.Error ?? $"无法定位对象 {code}");
            }

            Word.Range insert = locate.Range;

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

            CollapseEndBeforeTrailingBreaks(sentence);
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

            Word.Range endRange = doc.Paragraphs[doc.Paragraphs.Count].Range;
            endRange.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
            endRange.Text = "\r参考数据来源\r";
            Word.Range titleRange = LastParagraphOf(endRange);
            doc.Bookmarks.Add(ProvenanceBookmarkNames.RefTitle(), titleRange.Duplicate);

            Word.Range lastRefRange = titleRange.Duplicate;

            var ordered = sentenceItems
                .Where(i => i.SentenceNumber.HasValue)
                .OrderBy(i => i.SentenceNumber.Value)
                .ToList();

            foreach (RenderedProvenanceItem item in ordered)
            {
                int n = item.SentenceNumber.Value;
                Word.Range lineRange = lastRefRange.Duplicate;
                lineRange.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                lineRange.Text = $"\r[{n}] {item.RenderedLineForReference}\r";
                lineRange = LastParagraphOf(lineRange);
                ProvenanceFormatting.ApplyStandaloneReferenceStyle(lineRange);
                doc.Bookmarks.Add(ProvenanceBookmarkNames.RefLine(n), lineRange.Duplicate);
                lastRefRange = lineRange.Duplicate;
            }

            Word.Range blockRange = doc.Range(titleRange.Start, lastRefRange.End);
            doc.Bookmarks.Add(ProvenanceBookmarkNames.RefBlock(), blockRange);
        }

        private static Word.Range LastParagraphOf(Word.Range range)
        {
            if (range == null || range.Paragraphs.Count == 0)
            {
                return range;
            }

            return range.Paragraphs[range.Paragraphs.Count].Range;
        }

        /// <summary>
        /// 探针：写完句末角标、写参考区之前，按 Characters 打出每个字符和 Start/End。
        /// </summary>
        public static void DumpCharactersAfterMarkers(Word.Document doc)
        {
            if (doc == null)
            {
                return;
            }

            var sb = new StringBuilder();
            int contentStart = doc.Content.Start;
            int contentEnd = doc.Content.End;
            sb.AppendLine("=== after [n] / before 参考区 ===");
            sb.AppendLine("Content.Start=" + contentStart + " Content.End=" + contentEnd);
            sb.AppendLine("Paragraphs.Count=" + doc.Paragraphs.Count);
            Word.Paragraph last = doc.Paragraphs[doc.Paragraphs.Count];
            sb.AppendLine(
                "LastPara Start=" + last.Range.Start
                + " End=" + last.Range.End
                + " Text=" + EscapeStoryChar(last.Range.Text));
            int planned = contentEnd - 1;
            sb.AppendLine("planned insert pos=Content.End-1=" + planned);
            if (planned >= contentStart && planned < contentEnd)
            {
                try
                {
                    string at = doc.Range(planned, planned + 1).Text;
                    sb.AppendLine("char@End-1=" + EscapeStoryChar(at));
                }
                catch (Exception ex)
                {
                    sb.AppendLine("char@End-1 THROW " + ex.Message);
                }
            }

            Word.Characters chars = doc.Characters;
            sb.AppendLine("Characters.Count=" + chars.Count);
            sb.AppendLine("--- Characters ---");
            for (int i = 1; i <= chars.Count; i++)
            {
                Word.Range c = chars[i];
                sb.AppendLine(
                    "i=" + i
                    + " Start=" + c.Start
                    + " End=" + c.End
                    + " " + EscapeStoryChar(c.Text));
            }

            string path = Path.Combine(
                @"D:\project\EasyWrite",
                "char-stream-after-marker.txt");
            File.AppendAllText(path, sb.ToString() + Environment.NewLine, Encoding.UTF8);
            Console.WriteLine(sb.ToString());
        }

        private static string EscapeStoryChar(string text)
        {
            if (text == null)
            {
                return "text=<null>";
            }

            var sb = new StringBuilder("text=");
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
                else if (ch == '\f')
                {
                    sb.Append("\\f");
                }
                else if (ch == '\x01')
                {
                    sb.Append("\\x01");
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

            sb.Append(" codes=");
            foreach (char ch in text)
            {
                sb.Append(((int)ch).ToString()).Append(' ');
            }

            return sb.ToString().TrimEnd();
        }

        private static void CollapseEndBeforeTrailingBreaks(Word.Range sentence)
        {
            if (sentence == null)
            {
                return;
            }

            Word.Document doc = sentence.Document;
            int start = sentence.Start;
            int end = sentence.End;
            while (end > start)
            {
                string ch;
                try
                {
                    ch = doc.Range(end - 1, end).Text ?? "";
                }
                catch
                {
                    break;
                }

                if (ch != "\r" && ch != "\n" && ch != "\a" && ch != "\f")
                {
                    break;
                }

                end--;
            }

            sentence.SetRange(start, end);
            sentence.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
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
