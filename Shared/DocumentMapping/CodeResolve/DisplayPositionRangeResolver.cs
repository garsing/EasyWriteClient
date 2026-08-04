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
        /// 多码跨度定位：将编码序列对应 storedText 拼接后整体 Find（短文本单次 / 长文本包围搜索），
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

            string tag = string.IsNullOrEmpty(debugTag)
                ? $"span[{string.Join(",", codeList)}]"
                : debugTag;

            System.Diagnostics.Debug.WriteLine(
                $"[DisplayPositionRangeResolver] ResolveSpan combinedFind " +
                $"codes={codeList.Count} textLen={combinedText.Length} tag={tag}");

            if (EasyWriteDiagnostics.IsEnabled(DebugCategory.DocumentMapping))
            {
                LogCombinedSpanDetails(codeList, combinedText, tag);
            }

            return StoredTextRangeLocator.LocateFirst(
                doc,
                combinedText,
                tableId,
                tag);
        }

        /// <summary>
        /// P_ 多码跨度定位：拼接段落 storedText 后整体 Find（镜像 ResolveSpan）。
        /// </summary>
        public static Word.Range ResolveParagraphSpan(
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

            IReadOnlyList<string> domain = domainOrderedCodes ?? BuildParagraphDomain(tableId, tableScope);
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

            string combinedText = BuildCombinedParagraphStoredText(codeList);
            if (string.IsNullOrEmpty(combinedText))
            {
                System.Diagnostics.Debug.WriteLine(
                    $"[DisplayPositionRangeResolver] ResolveParagraphSpan 拼接文本为空 codes=[{string.Join(",", codeList)}]");
                return null;
            }

            string tag = string.IsNullOrEmpty(debugTag)
                ? $"para_span[{string.Join(",", codeList)}]"
                : debugTag;

            System.Diagnostics.Debug.WriteLine(
                $"[DisplayPositionRangeResolver] ResolveParagraphSpan combinedFind " +
                $"codes={codeList.Count} textLen={combinedText.Length} tag={tag}");

            return StoredTextRangeLocator.LocateFirst(
                doc,
                combinedText,
                tableId,
                tag);
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

        private static string BuildCombinedParagraphStoredText(IReadOnlyList<string> codeList)
        {
            var sb = new StringBuilder();
            foreach (string code in codeList)
            {
                string content = DocumentState.GetParagraphContent(code);
                if (string.IsNullOrEmpty(content))
                {
                    System.Diagnostics.Debug.WriteLine(
                        $"[DisplayPositionRangeResolver] 段落编码无 storedText: {code}");
                    return null;
                }

                sb.Append(content);
            }

            return sb.ToString();
        }

        private static void LogCombinedSpanDetails(
            IReadOnlyList<string> codeList,
            string combinedText,
            string tag)
        {
            bool hasCr = combinedText.IndexOf('\r') >= 0;
            bool hasLf = combinedText.IndexOf('\n') >= 0;
            System.Diagnostics.Debug.WriteLine(
                $"[DisplayPositionRangeResolver] ResolveSpan detail tag={tag} " +
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
    }
}
