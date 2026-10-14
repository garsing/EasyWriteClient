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

        public const string AlreadyHasProvenanceError =
            "文档已有引用，须先 F_clear_data_provenance，再按清除后的编码建新引用。";

        public static bool HasProvenance(Word.Document doc)
        {
            if (doc == null)
            {
                return false;
            }

            if (CollectProvenanceBookmarkInfos(doc).Count > 0)
            {
                return true;
            }

            return ScanSuperscriptMarkers(doc, delete: false);
        }

        public static void ClearAll(Word.Document doc)
        {
            if (!HasProvenance(doc))
            {
                return;
            }

            var infos = CollectProvenanceBookmarkInfos(doc);

            var deletedSpans = new List<(int Start, int End)>();
            foreach (BookmarkInfo info in infos.Where(i => i.IsBlock).OrderByDescending(i => i.Start))
            {
                TryDeleteBookmarkByName(doc, info.Name, deletedSpans);
            }

            foreach (BookmarkInfo info in infos.Where(i => !i.IsBlock).OrderByDescending(i => i.Start))
            {
                if (deletedSpans.Any(span => info.Start >= span.Start && info.Start <= span.End))
                {
                    continue;
                }

                TryDeleteBookmarkByName(doc, info.Name, deletedSpans);
            }

            ScanSuperscriptMarkers(doc, delete: true);
            TrimTrailingEmptyParagraphs(doc);
        }

        private struct BookmarkInfo
        {
            public string Name;
            public int Start;
            public bool IsBlock;
        }

        private static List<BookmarkInfo> CollectProvenanceBookmarkInfos(Word.Document doc)
        {
            var list = new List<BookmarkInfo>();
            foreach (Word.Bookmark bookmark in doc.Bookmarks)
            {
                string name;
                try
                {
                    name = bookmark.Name;
                }
                catch
                {
                    continue;
                }

                if (!ProvenanceBookmarkNames.IsProvenanceBookmark(name))
                {
                    continue;
                }

                int start = int.MaxValue;
                try
                {
                    start = bookmark.Range.Start;
                }
                catch
                {
                }

                list.Add(new BookmarkInfo
                {
                    Name = name,
                    Start = start,
                    IsBlock = IsBlockBookmark(name)
                });
            }

            return list;
        }

        private static void TryDeleteBookmarkByName(
            Word.Document doc,
            string name,
            List<(int Start, int End)> deletedSpans)
        {
            try
            {
                if (string.IsNullOrEmpty(name) || !doc.Bookmarks.Exists(name))
                {
                    return;
                }

                Word.Bookmark bookmark = doc.Bookmarks[name];
                int start = bookmark.Range.Start;
                int end = bookmark.Range.End;
                bookmark.Range.Delete();
                deletedSpans.Add((start, end));
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[DataProvenanceClear] 删除书签失败: {ex.Message}");
            }
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
                || name.StartsWith("yw_prov_cht_block_", StringComparison.Ordinal)
                || name.StartsWith("yw_prov_img_block_", StringComparison.Ordinal);
        }

        private static void TrimTrailingEmptyParagraphs(Word.Document doc)
        {
            try
            {
                while (doc.Paragraphs.Count > 1)
                {
                    Word.Paragraph last = doc.Paragraphs[doc.Paragraphs.Count];
                    string text = last.Range.Text ?? "";
                    string trimmed = text.Trim('\r', '\n', '\a', '\t', ' ', '\u00A0');
                    if (trimmed.Length > 0)
                    {
                        break;
                    }

                    last.Range.Delete();
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[DataProvenanceClear] 收文末空段失败: {ex.Message}");
            }
        }

        private static bool ScanSuperscriptMarkers(Word.Document doc, bool delete)
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

                    Word.Range probe = doc.Range(pos, pos + 1);
                    string ch = probe.Text ?? "";
                    if (ch == "[")
                    {
                        Word.Range marker = doc.Range(pos, pos);
                        marker.MoveEnd(Word.WdUnits.wdCharacter, 16);
                        string raw = marker.Text ?? "";
                        int close = raw.IndexOf(']');
                        if (close > 0)
                        {
                            Word.Range candidate = doc.Range(pos, pos + close + 1);
                            string text = (candidate.Text ?? "").Trim();
                            if (MarkerFormatRegex.IsMatch(text) && candidate.Font.Superscript != 0)
                            {
                                hit = true;
                                if (!delete)
                                {
                                    return true;
                                }

                                candidate.Delete();
                                end = doc.Content.End;
                                continue;
                            }
                        }
                    }

                    pos++;
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
