using System;
using System.Collections.Generic;
using System.Text;
using WordAddIn1;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1.DocumentMapping.CodeResolve
{
    public static class DisplayPositionRangeResolver
    {
        /// <summary>
        /// 多码跨度定位：将编码序列对应 storedText 拼接后整体 Find（短文本枚举 nth / 长文本包围搜索），
        /// 不再分别定位首句与末句再拼 Range（短/高频末句如「）」会导致 occurrence 错配）。
        /// </summary>
        public static Word.Range ResolveSpan(
            Word.Document doc,
            IReadOnlyList<string> codeList,
            int displayStartPosition,
            string tableId,
            TableScopeIndex tableScope,
            IReadOnlyList<string> domainOrderedCodes = null,
            string debugTag = null)
        {
            if (doc == null || codeList == null || codeList.Count == 0 || displayStartPosition < 0)
            {
                return null;
            }

            IReadOnlyList<string> domain = domainOrderedCodes ?? BuildDomain(tableId, tableScope);
            if (domain == null || displayStartPosition + codeList.Count > domain.Count)
            {
                return null;
            }

            for (int i = 0; i < codeList.Count; i++)
            {
                if (!string.Equals(domain[displayStartPosition + i], codeList[i], StringComparison.Ordinal))
                {
                    return null;
                }
            }

            string combinedText = BuildCombinedStoredText(codeList);
            if (string.IsNullOrEmpty(combinedText))
            {
                System.Diagnostics.Debug.WriteLine(
                    $"[DisplayPositionRangeResolver] ResolveSpan 拼接文本为空 codes=[{string.Join(",", codeList)}]");
                return null;
            }

            int sequenceOccurrence = ComputeSequenceOccurrenceIndexBefore(
                domain,
                codeList,
                displayStartPosition);

            string tag = string.IsNullOrEmpty(debugTag)
                ? $"span[{string.Join(",", codeList)}]"
                : debugTag;

            System.Diagnostics.Debug.WriteLine(
                $"[DisplayPositionRangeResolver] ResolveSpan combinedFind " +
                $"codes={codeList.Count} textLen={combinedText.Length} nth={sequenceOccurrence} tag={tag}");

            if (EasyWriteDiagnostics.IsEnabled(DebugCategory.DocumentMapping))
            {
                LogCombinedSpanDetails(codeList, combinedText, sequenceOccurrence, tag);
            }

            return StoredTextRangeLocator.Locate(
                doc,
                combinedText,
                sequenceOccurrence,
                tableId,
                tag);
        }

        public static Word.Range ResolveSingle(
            Word.Document doc,
            string code,
            int displayStartPosition,
            string tableId,
            TableScopeIndex tableScope,
            IReadOnlyList<string> domainOrderedCodes = null,
            string debugTag = null)
        {
            if (doc == null || string.IsNullOrEmpty(code) || displayStartPosition < 0)
            {
                return null;
            }

            IReadOnlyList<string> domain = domainOrderedCodes ?? BuildDomain(tableId, tableScope);
            if (domain == null || displayStartPosition >= domain.Count)
            {
                return null;
            }

            if (!string.Equals(domain[displayStartPosition], code, StringComparison.Ordinal))
            {
                return null;
            }

            int occurrenceIndex = ComputeOccurrenceIndexBefore(domain, code, displayStartPosition);
            return LocateSingleCode(doc, code, occurrenceIndex, tableId, tableScope, debugTag);
        }

        public static Word.Range ResolveParagraphSingle(
            Word.Document doc,
            string paragraphCode,
            int displayStartPosition,
            string tableId,
            TableScopeIndex tableScope,
            IReadOnlyList<string> domainOrderedCodes = null,
            string debugTag = null)
        {
            if (doc == null || string.IsNullOrEmpty(paragraphCode) || displayStartPosition < 0)
            {
                return null;
            }

            IReadOnlyList<string> domain = domainOrderedCodes ?? BuildParagraphDomain(tableId, tableScope);
            if (domain == null || displayStartPosition >= domain.Count)
            {
                return null;
            }

            if (!string.Equals(domain[displayStartPosition], paragraphCode, StringComparison.Ordinal))
            {
                return null;
            }

            int occurrenceIndex = ComputeOccurrenceIndexBefore(domain, paragraphCode, displayStartPosition);
            return ParagraphCodeLocator.LocateRange(
                doc,
                paragraphCode,
                occurrenceIndex,
                tableId,
                tableScope,
                allowAutoTableScope: false,
                debugTag: string.IsNullOrEmpty(debugTag) ? paragraphCode : debugTag);
        }

        public static List<int> ComputeOccurrenceIndexesFromDisplayStart(
            IReadOnlyList<string> codeList,
            int displayStartPosition,
            IReadOnlyList<string> domain)
        {
            var indexes = new List<int>();
            if (codeList == null || domain == null || displayStartPosition < 0)
            {
                return indexes;
            }

            for (int i = 0; i < codeList.Count; i++)
            {
                int pos = displayStartPosition + i;
                indexes.Add(pos < domain.Count
                    ? ComputeOccurrenceIndexBefore(domain, codeList[i], pos)
                    : 0);
            }

            return indexes;
        }

        public static int ComputeOccurrenceIndexBefore(
            IReadOnlyList<string> domain,
            string code,
            int displayPosition)
        {
            if (domain == null || string.IsNullOrEmpty(code) || displayPosition < 0)
            {
                return 0;
            }

            int count = 0;
            int limit = Math.Min(displayPosition, domain.Count);
            for (int i = 0; i < limit; i++)
            {
                if (string.Equals(domain[i], code, StringComparison.Ordinal))
                {
                    count++;
                }
            }

            return count;
        }

        /// <summary>
        /// 统计 displayStartPosition 之前，与 codeList 完全一致的连续序列出现次数（0-based nth）。
        /// </summary>
        public static int ComputeSequenceOccurrenceIndexBefore(
            IReadOnlyList<string> domain,
            IReadOnlyList<string> codeList,
            int displayStartPosition)
        {
            if (domain == null || codeList == null || codeList.Count == 0 || displayStartPosition <= 0)
            {
                return 0;
            }

            int count = 0;
            int maxStart = Math.Min(displayStartPosition, domain.Count - codeList.Count + 1);
            for (int i = 0; i < maxStart; i++)
            {
                if (MatchesSequenceAt(domain, codeList, i))
                {
                    count++;
                }
            }

            return count;
        }

        private static bool MatchesSequenceAt(
            IReadOnlyList<string> domain,
            IReadOnlyList<string> codeList,
            int start)
        {
            if (start < 0 || start + codeList.Count > domain.Count)
            {
                return false;
            }

            for (int j = 0; j < codeList.Count; j++)
            {
                if (!string.Equals(domain[start + j], codeList[j], StringComparison.Ordinal))
                {
                    return false;
                }
            }

            return true;
        }

        private static string BuildCombinedStoredText(IReadOnlyList<string> codeList)
        {
            var sb = new StringBuilder();
            foreach (string code in codeList)
            {
                string content = DocumentState.GetSentenceContent(code);
                if (string.IsNullOrEmpty(content))
                {
                    System.Diagnostics.Debug.WriteLine(
                        $"[DisplayPositionRangeResolver] 编码无 storedText: {code}");
                    return null;
                }

                sb.Append(content);
            }

            return sb.ToString();
        }

        private static void LogCombinedSpanDetails(
            IReadOnlyList<string> codeList,
            string combinedText,
            int sequenceOccurrence,
            string tag)
        {
            bool hasCr = combinedText.IndexOf('\r') >= 0;
            bool hasLf = combinedText.IndexOf('\n') >= 0;
            System.Diagnostics.Debug.WriteLine(
                $"[DisplayPositionRangeResolver] ResolveSpan detail tag={tag} nth={sequenceOccurrence} " +
                $"textLen={combinedText.Length} hasCR={hasCr} hasLF={hasLf}");

            for (int i = 0; i < codeList.Count; i++)
            {
                string code = codeList[i];
                string content = DocumentState.GetSentenceContent(code) ?? "";
                System.Diagnostics.Debug.WriteLine(
                    $"[DisplayPositionRangeResolver]   [{i}] {code} len={content.Length} " +
                    $"text=\"{TextUtils.EscapeNewlinesForDb(content)}\"");
            }

            System.Diagnostics.Debug.WriteLine(
                $"[DisplayPositionRangeResolver] ResolveSpan combined=\"{TextUtils.EscapeNewlinesForDb(combinedText)}\"");
        }

        private static IReadOnlyList<string> BuildDomain(string tableId, TableScopeIndex tableScope)
        {
            if (!string.IsNullOrEmpty(tableId) && tableScope != null
                && tableScope.TryGetTableSentenceCodes(tableId, out IReadOnlyList<string> tableCodes))
            {
                return tableCodes;
            }

            return DocumentState.GetOrderedSnapshotCodes();
        }

        private static IReadOnlyList<string> BuildParagraphDomain(string tableId, TableScopeIndex tableScope)
        {
            if (!string.IsNullOrEmpty(tableId) && tableScope != null
                && tableScope.TryGetTableSentenceCodes(tableId, out IReadOnlyList<string> tableCodes))
            {
                return tableCodes;
            }

            return ParagraphCodeAmbiguityHelper.GetOrderedParagraphCodesFromDisplay();
        }

        private static Word.Range LocateSingleCode(
            Word.Document doc,
            string code,
            int occurrenceIndex,
            string tableId,
            TableScopeIndex tableScope,
            string debugTag)
        {
            return SentenceCodeLocator.LocateRange(
                doc,
                code,
                occurrenceIndex,
                tableId,
                tableScope,
                allowAutoTableScope: false,
                debugTag: debugTag);
        }
    }
}
