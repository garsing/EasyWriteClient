using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace WordAddIn1
{
    /// <summary>单次 invoke 内从 ReadText/DisplayContent 构建的表内句子索引（不写入 DocumentState 长期存储）。</summary>
    public sealed class TableScopeIndex
    {
        private readonly Dictionary<string, IReadOnlyList<string>> _tableToSentenceCodes;

        public TableScopeIndex(Dictionary<string, List<string>> tableToSentenceCodes)
        {
            if (tableToSentenceCodes == null)
            {
                _tableToSentenceCodes = new Dictionary<string, IReadOnlyList<string>>(StringComparer.Ordinal);
                return;
            }

            _tableToSentenceCodes = tableToSentenceCodes.ToDictionary(
                kv => kv.Key,
                kv => (IReadOnlyList<string>)kv.Value,
                StringComparer.Ordinal);
        }

        public IReadOnlyDictionary<string, IReadOnlyList<string>> TableToSentenceCodes => _tableToSentenceCodes;

        public bool TryGetTableSentenceCodes(string tableId, out IReadOnlyList<string> codes)
        {
            return _tableToSentenceCodes.TryGetValue(tableId, out codes);
        }

        public bool ContainsTableId(string tableId)
        {
            return !string.IsNullOrEmpty(tableId) && _tableToSentenceCodes.ContainsKey(tableId);
        }

        public List<int> GetTableLocalPositionsForCode(string tableId, string code)
        {
            var positions = new List<int>();
            if (!_tableToSentenceCodes.TryGetValue(tableId, out IReadOnlyList<string> codes))
            {
                return positions;
            }

            for (int i = 0; i < codes.Count; i++)
            {
                if (string.Equals(codes[i], code, StringComparison.Ordinal))
                {
                    positions.Add(i);
                }
            }

            return positions;
        }

        public bool TryResolveTableOccurrence(
            string tableId,
            string code,
            int tableOccurrenceIndex,
            out int tableLocalPosition)
        {
            tableLocalPosition = -1;
            List<int> positions = GetTableLocalPositionsForCode(tableId, code);
            if (tableOccurrenceIndex < 0 || tableOccurrenceIndex >= positions.Count)
            {
                return false;
            }

            tableLocalPosition = positions[tableOccurrenceIndex];
            return true;
        }
    }

    public static class TableSentenceScopeHelper
    {
        private static readonly Regex TableStartRegex = new Regex(
            @"<table\s+id=""([^""]+)""[^>]*>",
            RegexOptions.Compiled | RegexOptions.IgnoreCase);

        private static readonly Regex TableOpenTagRegex = new Regex(
            @"<table\b",
            RegexOptions.Compiled | RegexOptions.IgnoreCase);

        private static readonly Regex SentenceStartRegex = new Regex(
            @"\[((S_[0-9a-zA-Z]{5}))-start\]",
            RegexOptions.Compiled);

        private const string CloseTableTag = "</table>";

        /// <summary>
        /// 从 ProcessDocument 缓存构建表内索引。优先 DisplayContent（含 [S_*-start]），ReadText 仅有表界无句标。
        /// </summary>
        public static TableScopeIndex BuildIndexFromProcessCache(ProcessDocumentCacheEntry cache)
        {
            if (cache == null)
            {
                return BuildIndex("");
            }

            string display = cache.ChunkInfo?.FirstOrDefault()?.DisplayContent;
            if (!string.IsNullOrEmpty(display))
            {
                TableScopeIndex fromDisplay = BuildIndex(display);
                if (HasAnyTableSentenceCodes(fromDisplay))
                {
                    System.Diagnostics.Debug.WriteLine("[TableScope] BuildIndex 使用 DisplayContent");
                    return fromDisplay;
                }
            }

            System.Diagnostics.Debug.WriteLine("[TableScope] BuildIndex 回退 ReadText（DisplayContent 无表内句标）");
            return BuildIndex(cache.ReadText ?? "");
        }

        public static TableScopeIndex BuildIndex(string readText)
        {
            var tableToCodes = new Dictionary<string, List<string>>(StringComparer.Ordinal);
            if (string.IsNullOrEmpty(readText))
            {
                System.Diagnostics.Debug.WriteLine("[TableScope] BuildIndex: 输入文本为空，返回空索引");
                return new TableScopeIndex(tableToCodes);
            }

            foreach (Match startMatch in TableStartRegex.Matches(readText))
            {
                string tableId = startMatch.Groups[1].Value;
                if (string.IsNullOrEmpty(tableId))
                {
                    continue;
                }

                int blockStart = startMatch.Index;
                int blockEnd = FindMatchingCloseTag(readText, blockStart);
                if (blockEnd < 0)
                {
                    System.Diagnostics.Debug.WriteLine($"[TableScope] WARN 表格 {tableId} 未找到闭合标签，pos={blockStart}");
                    continue;
                }

                string block = readText.Substring(blockStart, blockEnd - blockStart);
                List<string> codes = ExtractSentenceCodesFromOuterTable(block);
                tableToCodes[tableId] = codes;
                System.Diagnostics.Debug.WriteLine(
                    $"[TableScope] 解析表格 {tableId}：{codes.Count} 个 S_ 标记");
            }

            System.Diagnostics.Debug.WriteLine($"[TableScope] BuildIndex 完成：{tableToCodes.Count} 张表");
            return new TableScopeIndex(tableToCodes);
        }

        /// <summary>
        /// 解析 step3 实际使用的 table_id：显式指定优先；否则当编码在快照中的出现次数与唯一表内次数一致时自动限定表域。
        /// </summary>
        public static string ResolveEffectiveTableId(
            string code,
            string explicitTableId,
            TableScopeIndex tableScope)
        {
            if (!string.IsNullOrEmpty(explicitTableId))
            {
                return explicitTableId;
            }

            if (tableScope == null || string.IsNullOrEmpty(code))
            {
                return null;
            }

            string soleTableId = null;
            int soleTableOccurrences = 0;

            foreach (KeyValuePair<string, IReadOnlyList<string>> kv in tableScope.TableToSentenceCodes)
            {
                int count = tableScope.GetTableLocalPositionsForCode(kv.Key, code).Count;
                if (count == 0)
                {
                    continue;
                }

                if (soleTableId != null)
                {
                    return null;
                }

                soleTableId = kv.Key;
                soleTableOccurrences = count;
            }

            if (soleTableId == null)
            {
                return null;
            }

            int snapshotCount = DocumentState.GetSnapshotPositionsForCode(code, DocumentState.Snapshot).Count;
            if (snapshotCount == soleTableOccurrences)
            {
                System.Diagnostics.Debug.WriteLine(
                    $"[TableScope] 自动限定表格域 code={code} table={soleTableId}（快照 {snapshotCount} 次均在该表）");
                return soleTableId;
            }

            return null;
        }

        private static int FindMatchingCloseTag(string text, int openTagStart)
        {
            int depth = 0;
            int i = openTagStart;

            while (i < text.Length)
            {
                Match open = TableOpenTagRegex.Match(text, i);
                if (open.Success && open.Index == i)
                {
                    depth++;
                    i = open.Index + open.Length;
                    continue;
                }

                if (i + CloseTableTag.Length <= text.Length
                    && string.Compare(text, i, CloseTableTag, 0, CloseTableTag.Length, StringComparison.OrdinalIgnoreCase) == 0)
                {
                    depth--;
                    i += CloseTableTag.Length;
                    if (depth == 0)
                    {
                        return i;
                    }

                    continue;
                }

                i++;
            }

            return -1;
        }

        private static List<string> ExtractSentenceCodesFromOuterTable(string block)
        {
            var codes = new List<string>();
            if (string.IsNullOrEmpty(block))
            {
                return codes;
            }

            int contentStart = block.IndexOf('>');
            if (contentStart < 0)
            {
                return codes;
            }

            contentStart++;
            int nestedDepth = 0;
            int i = contentStart;

            while (i < block.Length)
            {
                if (TryMatchTag(block, i, "<table", out int openLen))
                {
                    nestedDepth++;
                    i += openLen;
                    continue;
                }

                if (TryMatchTag(block, i, CloseTableTag, out int closeLen))
                {
                    if (nestedDepth == 0)
                    {
                        break;
                    }

                    nestedDepth--;
                    i += closeLen;
                    continue;
                }

                if (nestedDepth == 0)
                {
                    Match sentenceMatch = SentenceStartRegex.Match(block, i);
                    if (sentenceMatch.Success && sentenceMatch.Index == i)
                    {
                        codes.Add(sentenceMatch.Groups[1].Value);
                        i = sentenceMatch.Index + sentenceMatch.Length;
                        continue;
                    }
                }

                i++;
            }

            return codes;
        }

        private static bool HasAnyTableSentenceCodes(TableScopeIndex index)
        {
            if (index?.TableToSentenceCodes == null)
            {
                return false;
            }

            foreach (IReadOnlyList<string> codes in index.TableToSentenceCodes.Values)
            {
                if (codes != null && codes.Count > 0)
                {
                    return true;
                }
            }

            return false;
        }

        private static bool TryMatchTag(string text, int index, string tag, out int length)
        {
            length = 0;
            if (index + tag.Length > text.Length)
            {
                return false;
            }

            if (string.Compare(text, index, tag, 0, tag.Length, StringComparison.OrdinalIgnoreCase) != 0)
            {
                return false;
            }

            length = tag.Length;
            return true;
        }
    }
}
