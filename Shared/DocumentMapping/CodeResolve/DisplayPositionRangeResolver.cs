using System;
using System.Collections.Generic;
using System.Text;
using WordAddIn1;
using WordAddIn1.DocumentHost;
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
                HostProbe.Log(
                    "span-null",
                    "why=args doc=" + (doc == null ? "null" : "ok")
                    + " codes=" + (codeList == null ? "null" : codeList.Count.ToString())
                    + " start=" + displayStartPosition);
                return null;
            }

            IReadOnlyList<string> domain = domainOrderedCodes ?? BuildDomain(tableId, tableScope);
            if (domain == null || displayStartPosition + codeList.Count > domain.Count)
            {
                HostProbe.Log(
                    "span-null",
                    "why=domain domainN=" + (domain == null ? "null" : domain.Count.ToString())
                    + " start=" + displayStartPosition
                    + " need=" + codeList.Count
                    + " passedDomain=" + (domainOrderedCodes == null ? "no" : "yes"));
                return null;
            }

            for (int i = 0; i < codeList.Count; i++)
            {
                if (!string.Equals(domain[displayStartPosition + i], codeList[i], StringComparison.Ordinal))
                {
                    HostProbe.Log(
                        "span-null",
                        "why=mismatch i=" + i
                        + " domain=" + domain[displayStartPosition + i]
                        + " want=" + codeList[i]);
                    return null;
                }
            }

            string combinedText = BuildCombinedStoredText(codeList);
            if (string.IsNullOrEmpty(combinedText))
            {
                HostProbe.Log(
                    "span-null",
                    "why=empty-stored codes=" + string.Join(",", codeList));
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

            Word.Range hit = StoredTextRangeLocator.LocateFirst(
                doc,
                combinedText,
                tableId,
                tag);
            HostProbe.Log(
                "span-find",
                "tag=" + tag
                + " textLen=" + combinedText.Length
                + " text=" + TextUtils.EscapeNewlinesForDb(combinedText)
                + " table=" + (tableId ?? "")
                + " hit=" + (hit == null ? "null" : (hit.Start + "-" + hit.End)));
            return hit;
        }

        /// <summary>
        /// P_ 多码跨度：display 连续后按段号合并 Range，不 Find。
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

            string tag = string.IsNullOrEmpty(debugTag)
                ? $"para_span[{string.Join(",", codeList)}]"
                : debugTag;

            System.Diagnostics.Debug.WriteLine(
                $"[DisplayPositionRangeResolver] ResolveParagraphSpan 段号合并 " +
                $"codes={codeList.Count} tag={tag}");

            return ParagraphRangeLocator.TryResolve(doc, codeList, domain, tag);
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
