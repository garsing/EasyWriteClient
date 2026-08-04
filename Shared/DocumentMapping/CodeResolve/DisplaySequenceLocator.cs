using System;
using System.Collections.Generic;
using System.Linq;
using WordAddIn1;

namespace WordAddIn1.DocumentMapping.CodeResolve
{
    public static class DisplaySequenceLocator
    {
        public const string ErrorAmbiguousSentenceCode = "ambiguous_sentence_code";
        public const string ErrorInvalidAnchorSequence = "invalid_anchor_sequence";
        public const string ErrorUnknownTableId = "unknown_table_id";
        public const string ErrorSentenceNotInTable = "sentence_not_in_table";

        public static DisplaySequenceLocateResult TryResolve(
            IReadOnlyList<string> codeList,
            DisplaySequenceLocateOptions options)
        {
            options = options ?? new DisplaySequenceLocateOptions();
            var result = new DisplaySequenceLocateResult();

            if (codeList == null || codeList.Count == 0)
            {
                result.MatchCount = 0;
                result.ErrorCode = ErrorInvalidAnchorSequence;
                return result;
            }

            List<string> normalized = codeList
                .Where(c => !string.IsNullOrWhiteSpace(c))
                .Select(c => c.Trim())
                .ToList();

            if (normalized.Count == 0)
            {
                result.MatchCount = 0;
                result.ErrorCode = ErrorInvalidAnchorSequence;
                return result;
            }

            if (normalized.Any(c => !StartsWithExpectedCodePrefix(c, options.StreamKind)))
            {
                result.MatchCount = 0;
                result.ErrorCode = ErrorInvalidAnchorSequence;
                return result;
            }

            if (!TryBuildDomain(options, out IReadOnlyList<string> domain, out int[] segIndices, out string domainError))
            {
                result.MatchCount = 0;
                result.ErrorCode = domainError ?? ErrorInvalidAnchorSequence;
                return result;
            }

            result.OrderedDomainCodes = domain;
            result.SegIndexByPosition = segIndices;

            if (domain.Count == 0)
            {
                result.MatchCount = 0;
                result.ErrorCode = ErrorInvalidAnchorSequence;
                return result;
            }

            if (normalized.Count > 1 && !HasSameSegPlacement(normalized, domain, segIndices))
            {
                result.MatchCount = 0;
                result.ErrorCode = ErrorInvalidAnchorSequence;
                return result;
            }

            List<int> matchStarts = FindMatchStarts(normalized, domain, segIndices);
            result.MatchStartPositions = matchStarts;
            result.MatchCount = matchStarts.Count;

            if (matchStarts.Count == 0)
            {
                result.ErrorCode = ErrorInvalidAnchorSequence;
                return result;
            }

            if (matchStarts.Count == 1)
            {
                result.DisplayStartPosition = matchStarts[0];
                return result;
            }

            result.ErrorCode = ErrorAmbiguousSentenceCode;
            return result;
        }

        public static int CountSubstringMatches(
            IReadOnlyList<string> codeList,
            IReadOnlyList<string> domain,
            int[] segIndices = null)
        {
            if (codeList == null || domain == null || codeList.Count == 0 || domain.Count < codeList.Count)
            {
                return 0;
            }

            return FindMatchStarts(codeList.ToList(), domain, segIndices).Count;
        }

        public static string TryExpandSuggestedLocatorCodes(
            IReadOnlyList<string> anchorCodeList,
            int matchStartPosition,
            IReadOnlyList<string> domain,
            int[] segIndices)
        {
            if (anchorCodeList == null || anchorCodeList.Count == 0 || domain == null)
            {
                return null;
            }

            if (matchStartPosition < 0 || matchStartPosition >= domain.Count)
            {
                return null;
            }

            var codes = new List<string>(anchorCodeList);
            int seg = GetSegIndex(segIndices, matchStartPosition);

            while (true)
            {
                int nextPos = matchStartPosition + codes.Count;
                if (nextPos >= domain.Count || GetSegIndex(segIndices, nextPos) != seg)
                {
                    break;
                }

                codes.Add(domain[nextPos]);
                if (CountSubstringMatches(codes, domain, segIndices) == 1)
                {
                    return string.Join(",", codes);
                }
            }

            return null;
        }

        private static bool TryBuildDomain(
            DisplaySequenceLocateOptions options,
            out IReadOnlyList<string> domain,
            out int[] segIndices,
            out string errorCode)
        {
            domain = Array.Empty<string>();
            segIndices = Array.Empty<int>();
            errorCode = null;

            if (options?.StreamKind == CodeStreamKind.ParagraphDisplay)
            {
                return TryBuildParagraphDomain(options, out domain, out segIndices, out errorCode);
            }

            string snapshot = options.Snapshot ?? DocumentState.Snapshot ?? "";
            string segSnapshot = options.SegSnapshot ?? DocumentState.SegSnapshot ?? "";

            if (!string.IsNullOrEmpty(options.TableId))
            {
                TableScopeIndex tableScope = options.TableScope;
                if (tableScope == null)
                {
                    errorCode = ErrorUnknownTableId;
                    return false;
                }

                if (!tableScope.ContainsTableId(options.TableId)
                    && DocumentState.GetTableIndex(options.TableId) < 0)
                {
                    errorCode = ErrorUnknownTableId;
                    return false;
                }

                if (!tableScope.TryGetTableSentenceCodes(options.TableId, out IReadOnlyList<string> tableCodes)
                    || tableCodes == null
                    || tableCodes.Count == 0)
                {
                    errorCode = ErrorSentenceNotInTable;
                    return false;
                }

                domain = tableCodes.ToList();
                segIndices = BuildSegIndicesForTableDomain(domain, snapshot, segSnapshot);
                return true;
            }

            domain = DocumentState.GetOrderedSentenceCodesFromSnapshot(snapshot);
            segIndices = DocumentState.BuildSegIndexBySnapshotPosition(snapshot, segSnapshot);
            return true;
        }

        private static bool TryBuildParagraphDomain(
            DisplaySequenceLocateOptions options,
            out IReadOnlyList<string> domain,
            out int[] segIndices,
            out string errorCode)
        {
            domain = Array.Empty<string>();
            segIndices = Array.Empty<int>();
            errorCode = null;

            if (!string.IsNullOrEmpty(options.TableId))
            {
                TableScopeIndex tableScope = options.TableScope;
                if (tableScope == null)
                {
                    errorCode = ErrorUnknownTableId;
                    return false;
                }

                if (!tableScope.ContainsTableId(options.TableId)
                    && DocumentState.GetTableIndex(options.TableId) < 0)
                {
                    errorCode = ErrorUnknownTableId;
                    return false;
                }

                if (!tableScope.TryGetTableSentenceCodes(options.TableId, out IReadOnlyList<string> tableCodes)
                    || tableCodes == null
                    || tableCodes.Count == 0)
                {
                    errorCode = ErrorSentenceNotInTable;
                    return false;
                }

                domain = tableCodes.ToList();
                segIndices = new int[domain.Count];
                return true;
            }

            List<string> ordered = ParagraphCodeAmbiguityHelper.GetOrderedParagraphCodesFromDisplay();
            domain = ordered;
            segIndices = ordered.Count > 0 ? new int[ordered.Count] : Array.Empty<int>();
            return true;
        }

        private static bool StartsWithExpectedCodePrefix(string code, CodeStreamKind streamKind)
        {
            if (string.IsNullOrEmpty(code))
            {
                return false;
            }

            return streamKind == CodeStreamKind.ParagraphDisplay
                ? code.StartsWith("P_", StringComparison.Ordinal)
                : code.StartsWith("S_", StringComparison.Ordinal);
        }

        internal static int[] BuildSegIndicesForTableDomain(
            IReadOnlyList<string> tableDomain,
            string snapshot,
            string segSnapshot)
        {
            if (tableDomain == null || tableDomain.Count == 0)
            {
                return Array.Empty<int>();
            }

            IReadOnlyList<string> globalCodes = DocumentState.GetOrderedSentenceCodesFromSnapshot(snapshot);
            int[] globalSeg = DocumentState.BuildSegIndexBySnapshotPosition(snapshot, segSnapshot);
            List<int> globalPositions = MapSubsequenceToGlobalPositions(tableDomain, globalCodes);
            if (globalPositions == null || globalPositions.Count != tableDomain.Count)
            {
                return new int[tableDomain.Count];
            }

            var segIndices = new int[tableDomain.Count];
            for (int i = 0; i < tableDomain.Count; i++)
            {
                int globalPos = globalPositions[i];
                segIndices[i] = globalPos >= 0 && globalPos < globalSeg.Length ? globalSeg[globalPos] : 0;
            }

            return segIndices;
        }

        internal static List<int> MapSubsequenceToGlobalPositions(
            IReadOnlyList<string> subsequence,
            IReadOnlyList<string> globalCodes)
        {
            var positions = new List<int>();
            if (subsequence == null || globalCodes == null || subsequence.Count == 0)
            {
                return positions;
            }

            int subIndex = 0;
            for (int globalIndex = 0; globalIndex < globalCodes.Count && subIndex < subsequence.Count; globalIndex++)
            {
                if (string.Equals(globalCodes[globalIndex], subsequence[subIndex], StringComparison.Ordinal))
                {
                    positions.Add(globalIndex);
                    subIndex++;
                }
            }

            return subIndex == subsequence.Count ? positions : null;
        }

        private static bool HasSameSegPlacement(
            IReadOnlyList<string> codeList,
            IReadOnlyList<string> domain,
            int[] segIndices)
        {
            for (int start = 0; start <= domain.Count - codeList.Count; start++)
            {
                if (!MatchesAt(codeList, domain, start))
                {
                    continue;
                }

                if (IsSameSeg(start, codeList.Count, segIndices))
                {
                    return true;
                }
            }

            return false;
        }

        private static List<int> FindMatchStarts(
            IReadOnlyList<string> codeList,
            IReadOnlyList<string> domain,
            int[] segIndices)
        {
            var starts = new List<int>();
            if (codeList.Count == 0 || domain.Count < codeList.Count)
            {
                return starts;
            }

            for (int start = 0; start <= domain.Count - codeList.Count; start++)
            {
                if (!MatchesAt(codeList, domain, start))
                {
                    continue;
                }

                if (!IsSameSeg(start, codeList.Count, segIndices))
                {
                    continue;
                }

                starts.Add(start);
            }

            return starts;
        }

        private static bool MatchesAt(IReadOnlyList<string> codeList, IReadOnlyList<string> domain, int start)
        {
            for (int i = 0; i < codeList.Count; i++)
            {
                if (!string.Equals(domain[start + i], codeList[i], StringComparison.Ordinal))
                {
                    return false;
                }
            }

            return true;
        }

        private static bool IsSameSeg(int start, int length, int[] segIndices)
        {
            if (length <= 1)
            {
                return true;
            }

            if (segIndices == null || segIndices.Length == 0)
            {
                return true;
            }

            int seg = GetSegIndex(segIndices, start);
            for (int i = start + 1; i < start + length; i++)
            {
                if (GetSegIndex(segIndices, i) != seg)
                {
                    return false;
                }
            }

            return true;
        }

        private static int GetSegIndex(int[] segIndices, int position)
        {
            if (segIndices == null || position < 0 || position >= segIndices.Length)
            {
                return 0;
            }

            return segIndices[position];
        }
    }
}
