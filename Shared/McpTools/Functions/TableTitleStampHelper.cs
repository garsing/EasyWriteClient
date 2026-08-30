using System;
using System.Collections.Generic;
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

        /// <summary>
        /// 以 Word Title 为真源：按 GetAllTablesInOrder 下标认号。
        /// 合法且本遍未占用则沿用；空或撞号则新发（写回仍走 Apply）。
        /// </summary>
        public static List<string> CollectIdsFromWord(Word.Document document, ISet<string> occupied)
        {
            var ids = new List<string>();
            if (document == null)
            {
                return ids;
            }

            if (occupied == null)
            {
                throw new ArgumentNullException(nameof(occupied));
            }

            List<Word.Table> tables = TableResolveHelper.GetAllTablesInOrder(document);
            for (int i = 0; i < tables.Count; i++)
            {
                string id = null;
                bool reused = false;
                if (TryReadTitle(tables[i], out string title)
                    && TryParse(title, out string existing, out _)
                    && occupied.Add(existing))
                {
                    id = existing;
                    reused = true;
                }
                else
                {
                    id = NewRandomTableId(occupied);
                }

                ids.Add(id);
                System.Diagnostics.Debug.WriteLine(
                    $"[TableStamp] Word真源 i={i} {(reused ? "沿用" : "新发")} {id ?? "(空)"}");
            }

            return ids;
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
                System.Diagnostics.Debug.WriteLine("[TableStamp] 读 Title 失败: " + ex.Message);
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
                System.Diagnostics.Debug.WriteLine("[TableStamp] 写 Title 失败: " + ex.Message);
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
        /// 按序把 assignedIds 写入各表 Title。写失败则该位改为内容哈希号（就地改 assignedIds）。
        /// tableContents 可空；空则用 table.Range.Text 做 O7。
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

                bool alreadyOk = TryReadTitle(table, out string current)
                    && TryParse(current, out string existingId, out _)
                    && existingId == wantId;
                if (alreadyOk)
                {
                    continue;
                }

                if (TryWriteStamp(table, wantId))
                {
                    System.Diagnostics.Debug.WriteLine($"[TableStamp] i={i} 写入 {wantId}");
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
                System.Diagnostics.Debug.WriteLine($"[TableStamp] i={i} 写失败，O7 哈希 {hashId}");
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
