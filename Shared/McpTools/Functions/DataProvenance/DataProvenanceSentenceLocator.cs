using System;
using System.Collections.Generic;
using System.Linq;
using WordAddIn1.DocumentMapping.CodeResolve;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    /// <summary>
    /// 句子溯源锚点定位：序列消歧 → ResolveSpan → Span 内末句（无 occurrenceIndex）。
    /// </summary>
    public static class DataProvenanceSentenceLocator
    {
        public static Word.Range LocateLastSentenceRange(
            Word.Document doc,
            ProvenanceAnchorInput anchor,
            string debugTag)
        {
            if (!TryLocateLastSentenceRange(doc, anchor, debugTag, out Word.Range range, out string error))
            {
                throw new InvalidOperationException(error ?? "无法定位句子");
            }

            return range;
        }

        public static bool TryLocateLastSentenceRange(
            Word.Document doc,
            ProvenanceAnchorInput anchor,
            string debugTag,
            out Word.Range range,
            out string error)
        {
            range = null;
            error = null;

            if (doc == null || anchor == null)
            {
                error = "缺少 anchor";
                return false;
            }

            string codesRaw = anchor.Code;
            List<string> codeList = ParseCodeList(codesRaw)
                .Where(c => c.StartsWith("S_", StringComparison.Ordinal))
                .ToList();

            if (codeList.Count == 0)
            {
                error = $"无法定位句子 {codesRaw}";
                return false;
            }

            string tableId = anchor.TableId;
            TableScopeIndex tableScope = null;
            if (!string.IsNullOrEmpty(tableId))
            {
                var scopeArgs = new ApplyFormatScopeArgs { TableId = tableId };
                tableScope = ApplyFormatAmbiguityHelper.BuildTableScopeOrNull(scopeArgs, out string cacheError);
                if (cacheError != null)
                {
                    error = cacheError;
                    return false;
                }

                if (tableScope == null || !tableScope.ContainsTableId(tableId))
                {
                    error = $"表域 '{tableId}' 不可用，请先 F_get_document_content";
                    return false;
                }
            }

            var locateOptions = new DisplaySequenceLocateOptions
            {
                TableId = tableId,
                TableScope = tableScope,
                StreamKind = CodeStreamKind.SentenceReadText,
                Snapshot = DocumentState.Snapshot,
                SegSnapshot = DocumentState.SegSnapshot,
            };

            DisplaySequenceLocateResult locate = DisplaySequenceLocator.TryResolve(codeList, locateOptions);
            if (locate == null || locate.MatchCount == 0)
            {
                error = $"无法定位句子 {codesRaw}";
                return false;
            }

            if (locate.MatchCount != 1 || !locate.DisplayStartPosition.HasValue)
            {
                error = $"句子编码歧义（{DisplaySequenceLocator.ErrorAmbiguousSentenceCode}），请扩充 anchor.code 或加 in_table";
                return false;
            }

            Word.Range span = DisplayPositionRangeResolver.ResolveSpan(
                doc,
                codeList,
                locate.DisplayStartPosition.Value,
                tableId,
                tableScope,
                locate.OrderedDomainCodes,
                debugTag: debugTag ?? "provenance");

            if (span == null)
            {
                error = $"无法定位句子 {codesRaw}";
                return false;
            }

            range = SpanSentenceLocator.LocateLastSentence(doc, span, codeList, debugTag);
            if (range == null)
            {
                error = $"无法定位句子 {codesRaw}";
                return false;
            }

            return true;
        }

        public static string GetLastSentenceCode(string codesRaw)
        {
            List<string> codeList = ParseCodeList(codesRaw);
            return codeList.Count > 0 ? codeList[codeList.Count - 1] : codesRaw;
        }

        private static List<string> ParseCodeList(string codes)
        {
            if (string.IsNullOrWhiteSpace(codes))
            {
                return new List<string>();
            }

            return codes.Split(',')
                .Select(s => s.Trim())
                .Where(s => !string.IsNullOrEmpty(s))
                .ToList();
        }
    }
}
