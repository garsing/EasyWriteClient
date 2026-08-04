using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    public static class DataProvenanceClearHelper
    {
        private static readonly Regex MarkerFormatRegex = new Regex(@"^\[\d+\]$", RegexOptions.Compiled);
        private const int MaxFormatScanSteps = 200000;

        public static List<string> ClearIfNeeded(Word.Document doc)
        {
            var warnings = new List<string>();
            if (doc == null)
            {
                return warnings;
            }

            var provenanceBookmarks = CollectProvenanceBookmarks(doc);
            if (provenanceBookmarks.Count == 0)
            {
                return warnings;
            }

            bool formatHit = ScanAndRemoveSuperscriptMarkers(doc);
            if (formatHit)
            {
                warnings.Add("marker_removed_by_format_scan");
            }

            var blockBookmarks = provenanceBookmarks
                .Where(b => IsBlockBookmark(b.Name))
                .OrderByDescending(b => b.Range.Start)
                .ToList();

            var deletedSpans = new List<(int Start, int End)>();
            foreach (Word.Bookmark block in blockBookmarks)
            {
                try
                {
                    int start = block.Range.Start;
                    int end = block.Range.End;
                    block.Range.Delete();
                    deletedSpans.Add((start, end));
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"[DataProvenanceClear] 删除块书签失败: {ex.Message}");
                }
            }

            var remaining = provenanceBookmarks
                .Where(b => !IsBlockBookmark(b.Name))
                .OrderByDescending(b => b.Range.Start)
                .ToList();

            foreach (Word.Bookmark bookmark in remaining)
            {
                try
                {
                    int start = bookmark.Range.Start;
                    if (deletedSpans.Any(span => start >= span.Start && start <= span.End))
                    {
                        continue;
                    }

                    bookmark.Range.Delete();
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"[DataProvenanceClear] 删除书签失败: {ex.Message}");
                }
            }

            return warnings;
        }

        private static List<Word.Bookmark> CollectProvenanceBookmarks(Word.Document doc)
        {
            var list = new List<Word.Bookmark>();
            foreach (Word.Bookmark bookmark in doc.Bookmarks)
            {
                if (ProvenanceBookmarkNames.IsProvenanceBookmark(bookmark.Name))
                {
                    list.Add(bookmark);
                }
            }

            return list;
        }

        private static bool IsBlockBookmark(string name)
        {
            if (string.IsNullOrEmpty(name))
            {
                return false;
            }

            if (name == ProvenanceBookmarkNames.RefBlock())
            {
                return true;
            }

            return name.StartsWith("yw_prov_tbl_block_", StringComparison.Ordinal)
                || name.StartsWith("yw_prov_cht_block_", StringComparison.Ordinal);
        }

        private static bool ScanAndRemoveSuperscriptMarkers(Word.Document doc)
        {
            bool hit = false;
            try
            {
                int pos = doc.Content.Start;
                int end = doc.Content.End;
                int steps = 0;

                while (pos < end && steps < MaxFormatScanSteps)
                {
                    steps++;
                    AgentRunCancellation.ThrowIfCancelled();

                    Word.Range wordRange = doc.Range(pos, pos);
                    wordRange.MoveEnd(Word.WdUnits.wdWord, 1);
                    if (wordRange.End <= pos)
                    {
                        pos++;
                        continue;
                    }

                    string text = (wordRange.Text ?? "").Trim();
                    if (wordRange.Font.Superscript != 0 && MarkerFormatRegex.IsMatch(text))
                    {
                        wordRange.Delete();
                        hit = true;
                        end = doc.Content.End;
                        continue;
                    }

                    pos = wordRange.End;
                }
            }
            catch (OperationCanceledException)
            {
                throw;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[DataProvenanceClear] 格式扫描失败: {ex.Message}");
            }

            return hit;
        }
    }
}
