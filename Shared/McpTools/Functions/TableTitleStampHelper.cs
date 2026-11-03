using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text.RegularExpressions;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    /// <summary>
    /// 表格 Title 稳定编号：ew:T_ + 8 位。只负责发号/写戳，不按 Title 找表。
    /// </summary>
    public static class TableTitleStampHelper
    {
        public const int TableTitleMaxChars = 255;
        public const string StampHead = "ew:";

        private const string Alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789";
        private static readonly Regex StampRegex = new Regex(
            @"^ew:T_([A-Za-z0-9]{8})",
            RegexOptions.Compiled);
        private static readonly Regex BookmarkRegex = new Regex(
            @"^ewT_([A-Za-z0-9]{8,32})$",
            RegexOptions.Compiled);
        private static readonly Regex AlnumTableIdRegex = new Regex(
            @"^T_[A-Za-z0-9]+$",
            RegexOptions.Compiled);

        private struct StampBookmark
        {
            public string Name;
            public string TableId;
            public int Start;
        }

        public static bool TryParse(string title, out string tableId, out string userSuffix)
        {
            tableId = null;
            userSuffix = "";
            if (string.IsNullOrEmpty(title))
            {
                return false;
            }

            Match m = StampRegex.Match(title);
            if (!m.Success)
            {
                return false;
            }

            tableId = "T_" + m.Groups[1].Value;
            string rest = title.Substring(m.Length);
            if (rest.StartsWith(" ", StringComparison.Ordinal))
            {
                rest = rest.Substring(1);
            }

            userSuffix = rest;
            return true;
        }

        public static string Format(string tableId, string userSuffix)
        {
            string stamp = StampHead + tableId;
            if (string.IsNullOrEmpty(userSuffix))
            {
                return stamp;
            }

            return stamp + " " + userSuffix;
        }

        public static bool TryFitToMaxLength(string formatted, out string fitted)
        {
            fitted = formatted ?? "";
            if (fitted.Length <= TableTitleMaxChars)
            {
                return true;
            }

            Match m = StampRegex.Match(fitted);
            if (!m.Success)
            {
                return false;
            }

            string stamp = fitted.Substring(0, m.Length);
            bool hasSuffix = fitted.Length > m.Length && fitted[m.Length] == ' ';
            int minLen = hasSuffix ? stamp.Length + 1 : stamp.Length;
            if (minLen > TableTitleMaxChars)
            {
                return false;
            }

            fitted = fitted.Substring(0, TableTitleMaxChars);
            return true;
        }

        public static string NewRandomTableId(ISet<string> occupied)
        {
            if (occupied == null)
            {
                occupied = new HashSet<string>(StringComparer.Ordinal);
            }

            byte[] buf = new byte[8];
            for (int attempt = 0; attempt < 32; attempt++)
            {
                using (var rng = RandomNumberGenerator.Create())
                {
                    rng.GetBytes(buf);
                }
                char[] chars = new char[8];
                for (int i = 0; i < 8; i++)
                {
                    chars[i] = Alphabet[buf[i] % Alphabet.Length];
                }

                string id = "T_" + new string(chars);
                if (occupied.Add(id))
                {
                    return id;
                }
            }

            return null;
        }

        public static string DecideTableId(string caption, ISet<string> occupied)
        {
            if (occupied == null)
            {
                throw new ArgumentNullException(nameof(occupied));
            }

            if (TryParse(caption, out string existing, out _) && occupied.Add(existing))
            {
                return existing;
            }

            return NewRandomTableId(occupied);
        }

        public static bool IsAlnumTableId(string tableId)
        {
            return !string.IsNullOrEmpty(tableId) && AlnumTableIdRegex.IsMatch(tableId);
        }

        /// <summary>内容哈希表号。按哈希字节映射到与戳相同的字母数字，不用 Base64。</summary>
        public static string MintAlnumTableId(byte[] hash, int charCount)
        {
            if (hash == null || hash.Length == 0 || charCount < 1)
            {
                return null;
            }

            char[] chars = new char[charCount];
            for (int i = 0; i < charCount; i++)
            {
                chars[i] = Alphabet[hash[i % hash.Length] % Alphabet.Length];
            }

            return "T_" + new string(chars);
        }

        public static bool TryParseBookmark(string name, out string tableId)
        {
            tableId = null;
            if (string.IsNullOrEmpty(name))
            {
                return false;
            }

            Match m = BookmarkRegex.Match(name);
            if (!m.Success)
            {
                return false;
            }

            tableId = "T_" + m.Groups[1].Value;
            return true;
        }

        public static string BookmarkNameFor(string tableId)
        {
            if (!IsAlnumTableId(tableId))
            {
                return null;
            }

            string body = tableId.Substring(2);
            if (body.Length < 8 || body.Length > 32)
            {
                return null;
            }

            return "ewT_" + body;
        }

        public const string ChannelTitle = "title";
        public const string ChannelBookmark = "bookmark";

        public sealed class StampPassResult
        {
            public string Channel;
            public string NumberedMarkup;
            public List<string> Ids = new List<string>();
        }

        /// <summary>
        /// 骨架对齐 COM 后读戳或盖戳。一份文档只走 Title 或只走书签。已有号不动。
        /// </summary>
        public static StampPassResult StampAlignedTables(
            Word.Document document,
            string skeletonMarkup)
        {
            var result = new StampPassResult
            {
                Channel = "",
                NumberedMarkup = skeletonMarkup ?? "",
                Ids = new List<string>()
            };
            if (document == null)
            {
                return result;
            }

            List<Word.Table> tables = TableResolveHelper.GetAllTablesInOrder(document, skeletonMarkup);
            string channel = DecideChannel(document, tables);
            result.Channel = channel;
            var occupied = new HashSet<string>(StringComparer.Ordinal);
            var ids = new List<string>();
            for (int i = 0; i < tables.Count; i++)
            {
                Word.Table table = tables[i];
                string id = ReadExistingId(document, table, channel, occupied);
                string how = string.IsNullOrEmpty(id) ? "新发" : (channel == ChannelTitle ? "沿用Title" : "沿用书签");
                if (string.IsNullOrEmpty(id))
                {
                    id = NewRandomTableId(occupied);
                    if (!string.IsNullOrEmpty(id) && WriteStamp(document, table, id, channel))
                    {
                        how = channel == ChannelTitle ? "写入 Title" : "写入书签";
                    }
                }

                ids.Add(id);
                LogStamp($"[TableStamp] Word真源 i={i} {how} {id ?? "(空)"}");
            }

            result.Ids = ids;
            result.NumberedMarkup = ReplaceSkeletonIds(skeletonMarkup, ids);
            return result;
        }

        public static List<string> CollectIdsFromWord(Word.Document document, ISet<string> occupied)
        {
            return CollectIdsFromWord(document, occupied, display: null);
        }

        public static List<string> CollectIdsFromWord(
            Word.Document document,
            ISet<string> occupied,
            string display)
        {
            StampPassResult pass = StampAlignedTables(document, display);
            if (occupied != null)
            {
                occupied.Clear();
                for (int i = 0; i < pass.Ids.Count; i++)
                {
                    if (!string.IsNullOrEmpty(pass.Ids[i]))
                    {
                        occupied.Add(pass.Ids[i]);
                    }
                }
            }

            DocumentState.SetTableStampChannel(pass.Channel);
            DocumentState.SetNumberedTableMarkup(pass.NumberedMarkup);
            return pass.Ids;
        }

        private static string DecideChannel(Word.Document document, IList<Word.Table> tables)
        {
            bool anyTitle = false;
            bool anyBookmark = false;
            for (int i = 0; i < tables.Count; i++)
            {
                if (TryReadTitle(tables[i], out string title)
                    && TryParse(title, out _, out _))
                {
                    anyTitle = true;
                    break;
                }

                if (!string.IsNullOrEmpty(BookmarkNameOnTable(document, tables[i])))
                {
                    anyBookmark = true;
                    break;
                }
            }

            if (anyTitle)
            {
                return ChannelTitle;
            }

            if (anyBookmark)
            {
                return ChannelBookmark;
            }

            if (tables.Count > 0 && TryWriteTitleProbe(tables[0]))
            {
                return ChannelTitle;
            }

            return ChannelBookmark;
        }

        private static bool TryWriteTitleProbe(Word.Table table)
        {
            if (table == null)
            {
                return false;
            }

            string previous = "";
            bool couldRead = TryReadTitle(table, out previous);
            if (!couldRead)
            {
                previous = "";
            }

            if (!TryWriteTitle(table, previous))
            {
                return false;
            }

            return true;
        }

        private static string ReadExistingId(
            Word.Document document,
            Word.Table table,
            string channel,
            ISet<string> occupied)
        {
            if (channel == ChannelTitle)
            {
                if (TryReadTitle(table, out string title)
                    && TryParse(title, out string existing, out _)
                    && occupied.Add(existing))
                {
                    return existing;
                }

                return null;
            }

            string nameOwned = BookmarkNameOnTable(document, table);
            if (!string.IsNullOrEmpty(nameOwned)
                && TryParseBookmark(nameOwned, out string fromBookmark)
                && occupied.Add(fromBookmark))
            {
                return fromBookmark;
            }

            return null;
        }

        private static bool WriteStamp(
            Word.Document document,
            Word.Table table,
            string tableId,
            string channel)
        {
            if (channel == ChannelTitle)
            {
                bool ok = TryWriteStamp(table, tableId);
                if (ok)
                {
                    LogStamp("[TableStamp] 写入 Title " + tableId);
                }

                return ok;
            }

            bool wrote = TryWriteBookmarkOnTable(document, table, tableId);
            if (wrote)
            {
                LogStamp("[TableStamp] 写入书签 " + tableId);
            }

            return wrote;
        }

        public static bool TryWriteBookmarkOnTable(Word.Document document, Word.Table table, string tableId)
        {
            if (document == null || table == null)
            {
                return false;
            }

            string name = BookmarkNameFor(tableId);
            if (name == null)
            {
                return false;
            }

            try
            {
                if (document.Bookmarks.Exists(name))
                {
                    Word.Bookmark existing = document.Bookmarks[name];
                    if (BookmarkCoversTable(existing, table))
                    {
                        return true;
                    }

                    LogStamp("[TableStamp] 书签名已被其他表占用: " + name);
                    return false;
                }

                Word.Range cover = table.Range.Duplicate;
                document.Bookmarks.Add(name, cover);
                return document.Bookmarks.Exists(name);
            }
            catch (Exception ex)
            {
                LogStamp("[TableStamp] 写书签失败: " + Describe(ex));
                return false;
            }
        }

        public static Word.Range TryGetBookmarkRange(Word.Document document, string tableId)
        {
            if (document == null || string.IsNullOrEmpty(tableId))
            {
                return null;
            }

            string name = BookmarkNameFor(tableId);
            if (name == null)
            {
                return null;
            }

            try
            {
                if (!document.Bookmarks.Exists(name))
                {
                    return null;
                }

                return document.Bookmarks[name].Range;
            }
            catch (Exception ex)
            {
                LogStamp("[TableStamp] 取书签失败: " + Describe(ex));
                return null;
            }
        }

        private static string BookmarkNameOnTable(Word.Document document, Word.Table table)
        {
            if (document == null || table == null)
            {
                return null;
            }

            try
            {
                int start;
                int end;
                try
                {
                    start = table.Range.Start;
                    end = table.Range.End;
                }
                catch (Exception)
                {
                    return null;
                }

                int count = document.Bookmarks.Count;
                for (int i = 1; i <= count; i++)
                {
                    Word.Bookmark bookmark = document.Bookmarks[i];
                    if (!TryParseBookmark(bookmark.Name, out _))
                    {
                        continue;
                    }

                    int markStart = bookmark.Range.Start;
                    int markEnd = bookmark.Range.End;
                    if (markStart == start && markEnd == end)
                    {
                        return bookmark.Name;
                    }
                }
            }
            catch (Exception ex)
            {
                LogStamp("[TableStamp] 读书签失败: " + Describe(ex));
            }

            return null;
        }

        private static bool BookmarkCoversTable(Word.Bookmark bookmark, Word.Table table)
        {
            if (bookmark == null || table == null)
            {
                return false;
            }

            try
            {
                return bookmark.Range.Start == table.Range.Start
                    && bookmark.Range.End == table.Range.End;
            }
            catch (Exception)
            {
                return false;
            }
        }

        public static string ReplaceSkeletonIds(string skeleton, IList<string> ids)
        {
            if (string.IsNullOrEmpty(skeleton) || ids == null || ids.Count == 0)
            {
                return skeleton ?? "";
            }

            var regex = new Regex(@"<table\s+id=""T_sk\d+""");
            int seen = -1;
            return regex.Replace(skeleton, match =>
            {
                seen++;
                if (seen >= 0 && seen < ids.Count && !string.IsNullOrEmpty(ids[seen]))
                {
                    return "<table id=\"" + ids[seen] + "\"";
                }

                return match.Value;
            });
        }

        public static bool TryReadTitle(Word.Table table, out string title)
        {
            title = "";
            if (table == null)
            {
                return false;
            }

            try
            {
                title = table.Title ?? "";
                return true;
            }
            catch (Exception ex)
            {
                LogStamp("[TableStamp] 读 Title 失败: " + Describe(ex));
                return false;
            }
        }

        public static bool TryWriteTitle(Word.Table table, string title)
        {
            if (table == null)
            {
                return false;
            }

            try
            {
                table.Title = title ?? "";
                return true;
            }
            catch (Exception ex)
            {
                LogStamp("[TableStamp] 写 Title 失败: " + Describe(ex));
                return false;
            }
        }

        /// <summary>
        /// Title 写不上时，把同一个 T_ 写进该书签。书签名 ewT_ + 表号正文，锚在该表 Range 起点。
        /// </summary>
        public static bool TryWriteBookmark(
            Word.Document document,
            IList<Word.Table> tables,
            int index,
            string tableId)
        {
            if (document == null || tables == null || index < 0 || index >= tables.Count)
            {
                return false;
            }

            string name = BookmarkNameFor(tableId);
            Word.Table table = tables[index];
            if (name == null || table == null)
            {
                return false;
            }

            try
            {
                if (document.Bookmarks.Exists(name))
                {
                    Word.Bookmark existing = document.Bookmarks[name];
                    if (OwnerIndex(tables, existing.Range.Start) == index)
                    {
                        return true;
                    }

                    LogStamp("[TableStamp] 书签名已被其他表占用: " + name);
                    return false;
                }

                var owned = new List<string>();
                int count = document.Bookmarks.Count;
                for (int i = 1; i <= count; i++)
                {
                    Word.Bookmark bookmark = document.Bookmarks[i];
                    if (!TryParseBookmark(bookmark.Name, out _))
                    {
                        continue;
                    }

                    if (OwnerIndex(tables, bookmark.Range.Start) == index)
                    {
                        owned.Add(bookmark.Name);
                    }
                }

                for (int i = 0; i < owned.Count; i++)
                {
                    DeleteBookmark(document, owned[i]);
                }

                Word.Range anchor = table.Range.Duplicate;
                anchor.Collapse(Word.WdCollapseDirection.wdCollapseStart);
                document.Bookmarks.Add(name, anchor);
                return document.Bookmarks.Exists(name);
            }
            catch (Exception ex)
            {
                LogStamp("[TableStamp] 写书签失败: " + Describe(ex));
                return false;
            }
        }

        /// <summary>给单张表写入指定 T_（空则整段戳，已有非戳则前缀）。</summary>
        public static bool TryWriteStamp(Word.Table table, string tableId)
        {
            if (table == null || string.IsNullOrEmpty(tableId) || !tableId.StartsWith("T_", StringComparison.Ordinal))
            {
                return false;
            }

            string suffix = "";
            if (TryReadTitle(table, out string current) && !string.IsNullOrEmpty(current))
            {
                if (TryParse(current, out _, out string parsedSuffix))
                {
                    suffix = parsedSuffix;
                }
                else
                {
                    suffix = current;
                }
            }

            string formatted = Format(tableId, suffix);
            if (!TryFitToMaxLength(formatted, out formatted))
            {
                return false;
            }

            return TryWriteTitle(table, formatted);
        }

        /// <summary>
        /// 按序把 assignedIds 写入各表 Title。Title 写失败则写书签，书签也失败才改为内容哈希号（就地改 assignedIds）。
        /// tableContents 可空；空则用 table.Range.Text 做哈希兜底。
        /// </summary>
        public static void ApplyAssignedIds(
            Word.Document document,
            IList<string> assignedIds,
            IList<string> tableContents)
        {
            if (document == null || assignedIds == null || assignedIds.Count == 0)
            {
                return;
            }

            List<Word.Table> tables = TableResolveHelper.GetAllTablesInOrder(document);
            int n = Math.Min(tables.Count, assignedIds.Count);
            if (tables.Count != assignedIds.Count)
            {
                System.Diagnostics.Debug.WriteLine(
                    $"[TableStamp] 表数 {tables.Count} 与编号数 {assignedIds.Count} 不一致，按 min={n} 写入");
            }

            for (int i = 0; i < n; i++)
            {
                string wantId = assignedIds[i];
                Word.Table table = tables[i];
                if (string.IsNullOrEmpty(wantId))
                {
                    continue;
                }

                bool titleOk = TryReadTitle(table, out string current)
                    && TryParse(current, out string existingId, out _)
                    && existingId == wantId;
                if (titleOk)
                {
                    RemoveStampBookmarks(document, tables, i);
                    continue;
                }

                if (TryWriteStamp(table, wantId))
                {
                    RemoveStampBookmarks(document, tables, i);
                    LogStamp($"[TableStamp] i={i} 写入 Title {wantId}");
                    continue;
                }

                if (TryWriteBookmark(document, tables, i, wantId))
                {
                    LogStamp($"[TableStamp] i={i} 写入书签 {wantId}");
                    continue;
                }

                string hashContent = null;
                if (tableContents != null && i < tableContents.Count)
                {
                    hashContent = tableContents[i];
                }
                else
                {
                    try
                    {
                        hashContent = table.Range != null ? (table.Range.Text ?? "") : "";
                    }
                    catch
                    {
                        hashContent = "";
                    }
                }

                string hashId = DocumentState.FindOrCreateTableId(hashContent);
                assignedIds[i] = hashId;
                LogStamp($"[TableStamp] i={i} 写 Title 与书签都失败，哈希 {hashId}");
            }
        }

        private static void RemoveStampBookmarks(Word.Document document, IList<Word.Table> tables, int index)
        {
            if (document == null || tables == null)
            {
                return;
            }

            try
            {
                var owned = new List<string>();
                int count = document.Bookmarks.Count;
                for (int i = 1; i <= count; i++)
                {
                    Word.Bookmark bookmark = document.Bookmarks[i];
                    if (!TryParseBookmark(bookmark.Name, out _))
                    {
                        continue;
                    }

                    if (OwnerIndex(tables, bookmark.Range.Start) == index)
                    {
                        owned.Add(bookmark.Name);
                    }
                }

                for (int i = 0; i < owned.Count; i++)
                {
                    DeleteBookmark(document, owned[i]);
                }
            }
            catch (Exception ex)
            {
                LogStamp("[TableStamp] 清理书签失败: " + Describe(ex));
            }
        }

        private static List<StampBookmark> ReadStampBookmarks(Word.Document document)
        {
            var marks = new List<StampBookmark>();
            if (document == null)
            {
                return marks;
            }

            try
            {
                int count = document.Bookmarks.Count;
                for (int i = 1; i <= count; i++)
                {
                    Word.Bookmark bookmark = document.Bookmarks[i];
                    if (!TryParseBookmark(bookmark.Name, out string tableId))
                    {
                        continue;
                    }

                    marks.Add(new StampBookmark
                    {
                        Name = bookmark.Name,
                        TableId = tableId,
                        Start = bookmark.Range.Start
                    });
                }
            }
            catch (Exception ex)
            {
                LogStamp("[TableStamp] 读书签失败: " + Describe(ex));
            }

            return marks;
        }

        private static string BookmarkIdOwnedBy(IList<Word.Table> tables, int index, List<StampBookmark> marks)
        {
            if (marks == null)
            {
                return null;
            }

            for (int i = 0; i < marks.Count; i++)
            {
                if (OwnerIndex(tables, marks[i].Start) == index)
                {
                    return marks[i].TableId;
                }
            }

            return null;
        }

        private static int OwnerIndex(IList<Word.Table> tables, int position)
        {
            int best = -1;
            int bestSpan = int.MaxValue;
            if (tables == null)
            {
                return best;
            }

            for (int i = 0; i < tables.Count; i++)
            {
                Word.Table table = tables[i];
                if (table == null)
                {
                    continue;
                }

                int start;
                int end;
                try
                {
                    start = table.Range.Start;
                    end = table.Range.End;
                }
                catch (Exception)
                {
                    continue;
                }

                if (position < start || position >= end)
                {
                    continue;
                }

                int span = end - start;
                if (span < bestSpan)
                {
                    best = i;
                    bestSpan = span;
                }
            }

            return best;
        }

        private static void DeleteBookmark(Word.Document document, string name)
        {
            try
            {
                if (document.Bookmarks.Exists(name))
                {
                    document.Bookmarks[name].Delete();
                }
            }
            catch (Exception ex)
            {
                LogStamp("[TableStamp] 删书签失败 " + name + ": " + Describe(ex));
            }
        }

        private static string Describe(Exception ex)
        {
            if (ex == null)
            {
                return "";
            }

            COMException com = ex as COMException;
            if (com == null)
            {
                return ex.Message;
            }

            return ex.Message + " hresult=0x" + com.ErrorCode.ToString("X8");
        }

        private static void LogStamp(string message)
        {
            System.Diagnostics.Debug.WriteLine(message);
            try
            {
                EasyWriteLog.WriteLine(message);
            }
            catch (Exception)
            {
            }
        }

        public static string ReplaceTableIdAtOccurrence(string text, int occurrenceIndex, string newId)
        {
            if (string.IsNullOrEmpty(text) || occurrenceIndex < 0 || string.IsNullOrEmpty(newId))
            {
                return text;
            }

            var regex = new Regex(@"<table\s+id=""([^""]+)""");
            int seen = -1;
            return regex.Replace(text, match =>
            {
                seen++;
                if (seen == occurrenceIndex)
                {
                    return "<table id=\"" + newId + "\"";
                }

                return match.Value;
            });
        }
    }
}
