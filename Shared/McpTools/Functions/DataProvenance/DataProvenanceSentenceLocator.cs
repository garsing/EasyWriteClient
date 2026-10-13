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
            var locate = InsertTargetLocatorHelper.TryResolveInsertRange(
                doc,
                new InsertTargetLocatorHelper.LocatorSpec
                {
                    Kind = InsertTargetLocatorHelper.LocatorKind.Sentence,
                    Codes = string.Join(",", codeList),
                    InTableId = tableId
                },
                InsertTargetLocatorHelper.LocateRangeKind.ContentSpan);
            if (!locate.Success || locate.Range == null)
            {
                error = locate.Error ?? $"无法定位句子 {codesRaw}";
                return false;
            }

            Word.Range span = locate.Range;

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
