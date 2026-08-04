using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace WordAddIn1.DocumentMapping.CodeResolve
{
    /// <summary>
    /// 从 paragraph display（[P_*-start]）构建表内段落索引，镜像 TableSentenceScopeHelper。
    /// </summary>
    public static class TableParagraphScopeHelper
    {
        private static readonly Regex TableStartRegex = new Regex(
            @"<table\s+id=""([^""]+)""[^>]*>",
            RegexOptions.Compiled | RegexOptions.IgnoreCase);

        private static readonly Regex TableOpenTagRegex = new Regex(
            @"<table\b",
            RegexOptions.Compiled | RegexOptions.IgnoreCase);

        private static readonly Regex ParagraphStartRegex = new Regex(
            @"\[((P_[0-9a-zA-Z]{5}))-start\]",
            RegexOptions.Compiled);

        private const string CloseTableTag = "</table>";

        /// <summary>从 ProcessDocument 缓存构建表内 P_ 索引；优先 ParagraphDisplayContent。</summary>
        public static TableScopeIndex BuildIndexFromProcessCache(ProcessDocumentCacheEntry cache)
        {
            if (cache == null)
            {
                return BuildIndexFromDocumentState();
            }

            if (cache.ChunkInfo != null)
            {
                var parts = new List<string>();
                foreach (WordDocumentExtractor.ChunkInfo chunk in cache.ChunkInfo)
                {
                    if (!string.IsNullOrEmpty(chunk?.ParagraphDisplayContent))
                    {
                        parts.Add(chunk.ParagraphDisplayContent);
                    }
                }

                if (parts.Count > 0)
                {
                    TableScopeIndex fromChunk = BuildIndex(string.Concat(parts));
                    if (HasAnyTableParagraphCodes(fromChunk))
                    {
                        System.Diagnostics.Debug.WriteLine("[TableParaScope] BuildIndex 使用 ChunkInfo.ParagraphDisplayContent");
                        return fromChunk;
                    }
                }
            }

            TableScopeIndex fromState = BuildIndexFromDocumentState();
            if (HasAnyTableParagraphCodes(fromState))
            {
                System.Diagnostics.Debug.WriteLine("[TableParaScope] BuildIndex 使用 DocumentState.ChunkParagraphDisplayContents");
                return fromState;
            }

            System.Diagnostics.Debug.WriteLine("[TableParaScope] BuildIndex 回退 ReadText（无表内 P_ 标记）");
            return BuildIndex(cache.ReadText ?? "");
        }

        public static TableScopeIndex BuildIndexFromDocumentState()
        {
            if (DocumentState.ChunkParagraphDisplayContents == null
                || DocumentState.ChunkParagraphDisplayContents.Count == 0)
            {
                return BuildIndex("");
            }

            return BuildIndex(string.Concat(DocumentState.ChunkParagraphDisplayContents));
        }

        public static TableScopeIndex BuildIndex(string paragraphDisplayText)
        {
            var tableToCodes = new Dictionary<string, List<string>>(StringComparer.Ordinal);
            if (string.IsNullOrEmpty(paragraphDisplayText))
            {
                return new TableScopeIndex(tableToCodes);
            }

            foreach (Match startMatch in TableStartRegex.Matches(paragraphDisplayText))
            {
                string tableId = startMatch.Groups[1].Value;
                if (string.IsNullOrEmpty(tableId))
                {
                    continue;
                }

                int blockStart = startMatch.Index;
                int blockEnd = FindMatchingCloseTag(paragraphDisplayText, blockStart);
                if (blockEnd < 0)
                {
                    continue;
                }

                string block = paragraphDisplayText.Substring(blockStart, blockEnd - blockStart);
                List<string> codes = ExtractParagraphCodesFromOuterTable(block);
                tableToCodes[tableId] = codes;
            }

            return new TableScopeIndex(tableToCodes);
        }

        /// <summary>
        /// P_ 自动表域：当编码在 paragraph display 中的总次数与唯一表内次数一致时限定 table_id。
        /// </summary>
        public static string ResolveEffectiveTableId(
            string paragraphCode,
            string explicitTableId,
            TableScopeIndex tableScope)
        {
            if (!string.IsNullOrEmpty(explicitTableId))
            {
                return explicitTableId;
            }

            if (tableScope == null || string.IsNullOrEmpty(paragraphCode))
            {
                return null;
            }

            string soleTableId = null;
            int soleTableOccurrences = 0;

            foreach (KeyValuePair<string, IReadOnlyList<string>> kv in tableScope.TableToSentenceCodes)
            {
                int count = tableScope.GetTableLocalPositionsForCode(kv.Key, paragraphCode).Count;
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

            int displayCount = CountCodeInParagraphDisplay(paragraphCode);
            if (displayCount == soleTableOccurrences && displayCount > 0)
            {
                System.Diagnostics.Debug.WriteLine(
                    $"[TableParaScope] 自动限定表格域 P_={paragraphCode} table={soleTableId}（display {displayCount} 次均在该表）");
                return soleTableId;
            }

            return null;
        }

        private static int CountCodeInParagraphDisplay(string paragraphCode)
        {
            string marker = "[" + paragraphCode + "-start]";
            int total = 0;
            foreach (string display in DocumentState.ChunkParagraphDisplayContents)
            {
                if (string.IsNullOrEmpty(display))
                {
                    continue;
                }

                int idx = 0;
                while ((idx = display.IndexOf(marker, idx, StringComparison.Ordinal)) >= 0)
                {
                    total++;
                    idx += marker.Length;
                }
            }

            return total;
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

        private static List<string> ExtractParagraphCodesFromOuterTable(string block)
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
                    Match paraMatch = ParagraphStartRegex.Match(block, i);
                    if (paraMatch.Success && paraMatch.Index == i)
                    {
                        codes.Add(paraMatch.Groups[1].Value);
                        i = paraMatch.Index + paraMatch.Length;
                        continue;
                    }
                }

                i++;
            }

            return codes;
        }

        private static bool HasAnyTableParagraphCodes(TableScopeIndex index)
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
