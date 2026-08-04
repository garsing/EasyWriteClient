using System;
using System.Collections.Generic;
using System.Linq;

namespace WordAddIn1
{
    public sealed class SentenceCodeText
    {
        public string Code { get; set; }
        public string Text { get; set; }
    }

    public sealed class SnapshotDiffResult
    {
        public List<SentenceCodeText> Sentences { get; set; } = new List<SentenceCodeText>();
        public string Source { get; set; }
        public string Warning { get; set; }
        public int? PreIndex { get; set; }
        public int? PostIndex { get; set; }

        public bool IsVerified => string.Equals(Source, "snapshot_verified", StringComparison.Ordinal);
    }

    /// <summary>
    /// step1/step4 快照邻居书签法定码（02）。
    /// </summary>
    public static class SnapshotDiffHelper
    {
        public const string SourceVerified = "snapshot_verified";
        public const string SourceUnavailable = "unavailable";

        public static List<string> ParseFlatSnapshot(string snapshot)
        {
            return DocumentState.GetOrderedSentenceCodesFromSnapshot(snapshot);
        }

        public static SnapshotDiffResult BuildForReplace(
            SnapshotRecord pre,
            SnapshotRecord post,
            List<string> anchorCodes)
        {
            if (pre == null || post == null)
            {
                return Unavailable(null, null, "snapshot_history_missing");
            }

            if (!TryDeriveReplaceBookmarks(pre, anchorCodes, out string left, out string right, out string warning))
            {
                return Unavailable(pre.Index, post.Index, warning ?? "anchor_not_found_in_pre");
            }

            return BuildBetweenBookmarks(pre, post, left, right);
        }

        public static SnapshotDiffResult BuildForInsert(
            SnapshotRecord pre,
            SnapshotRecord post,
            List<string> preAnchorCodes,
            List<string> postAnchorCodes)
        {
            if (pre == null || post == null)
            {
                return Unavailable(null, null, "snapshot_history_missing");
            }

            if (!TryDeriveInsertBookmarks(pre, preAnchorCodes, postAnchorCodes, out string left, out string right, out string warning))
            {
                return Unavailable(pre.Index, post.Index, warning ?? "insert_no_anchor");
            }

            return BuildBetweenBookmarks(pre, post, left, right);
        }

        public static SnapshotDiffResult BuildBulkEnvelope(
            SnapshotRecord pre,
            SnapshotRecord post,
            IEnumerable<BookmarkPair> pairs)
        {
            if (pre == null || post == null)
            {
                return Unavailable(null, null, "snapshot_history_missing");
            }

            var preCodes = ParseFlatSnapshot(pre.Snapshot);
            if (preCodes.Count == 0)
            {
                return Unavailable(pre.Index, post.Index, "anchor_not_found_in_pre");
            }

            bool anyMissingLeft = false;
            bool anyMissingRight = false;
            int? minLeftIdx = null;
            int? maxRightIdx = null;

            var pairList = (pairs ?? Enumerable.Empty<BookmarkPair>()).Where(p => p != null).ToList();
            if (pairList.Count == 0)
            {
                return Unavailable(pre.Index, post.Index, "ambiguous_interval");
            }

            foreach (BookmarkPair pair in pairList)
            {

                if (string.IsNullOrEmpty(pair.LeftBookmark))
                {
                    anyMissingLeft = true;
                }
                else
                {
                    int leftIdx = preCodes.IndexOf(pair.LeftBookmark);
                    if (leftIdx < 0)
                    {
                        return Unavailable(pre.Index, post.Index, "neighbor_not_found_in_post");
                    }

                    if (!minLeftIdx.HasValue || leftIdx < minLeftIdx.Value)
                    {
                        minLeftIdx = leftIdx;
                    }
                }

                if (string.IsNullOrEmpty(pair.RightBookmark))
                {
                    anyMissingRight = true;
                }
                else
                {
                    int rightIdx = preCodes.IndexOf(pair.RightBookmark);
                    if (rightIdx < 0)
                    {
                        return Unavailable(pre.Index, post.Index, "neighbor_not_found_in_post");
                    }

                    if (!maxRightIdx.HasValue || rightIdx > maxRightIdx.Value)
                    {
                        maxRightIdx = rightIdx;
                    }
                }
            }

            string leftCode = anyMissingLeft ? null : (minLeftIdx.HasValue ? preCodes[minLeftIdx.Value] : null);
            string rightCode = anyMissingRight ? null : (maxRightIdx.HasValue ? preCodes[maxRightIdx.Value] : null);
            return BuildBetweenBookmarks(pre, post, leftCode, rightCode);
        }

        public static bool TryDeriveReplaceBookmarks(
            SnapshotRecord pre,
            List<string> anchorCodes,
            out string leftBookmark,
            out string rightBookmark,
            out string warning)
        {
            leftBookmark = null;
            rightBookmark = null;
            warning = null;

            var preCodes = ParseFlatSnapshot(pre?.Snapshot);
            if (anchorCodes == null || anchorCodes.Count == 0 || !TryFindContinuousBlock(preCodes, anchorCodes, out int start, out int end))
            {
                warning = "anchor_not_found_in_pre";
                return false;
            }

            if (start > 0)
            {
                leftBookmark = preCodes[start - 1];
            }

            if (end < preCodes.Count - 1)
            {
                rightBookmark = preCodes[end + 1];
            }

            return true;
        }

        public static bool TryDeriveInsertBookmarks(
            SnapshotRecord pre,
            List<string> preAnchorCodes,
            List<string> postAnchorCodes,
            out string leftBookmark,
            out string rightBookmark,
            out string warning)
        {
            leftBookmark = null;
            rightBookmark = null;
            warning = null;

            var preSCodes = FilterSentenceCodes(preAnchorCodes);
            var postSCodes = FilterSentenceCodes(postAnchorCodes);

            if (HasTableOnlyAnchors(preAnchorCodes, postAnchorCodes, preSCodes, postSCodes))
            {
                warning = "insert_table_anchor";
                return false;
            }

            if (preSCodes.Count == 0 && postSCodes.Count == 0)
            {
                warning = "insert_no_anchor";
                return false;
            }

            var preCodes = ParseFlatSnapshot(pre?.Snapshot);

            if (preSCodes.Count > 0 && postSCodes.Count > 0)
            {
                leftBookmark = preSCodes[preSCodes.Count - 1];
                rightBookmark = postSCodes[0];
                return ValidateBookmarksInPre(preCodes, leftBookmark, rightBookmark, out warning);
            }

            if (preSCodes.Count > 0)
            {
                string anchor = preSCodes[preSCodes.Count - 1];
                int i = preCodes.IndexOf(anchor);
                if (i < 0)
                {
                    warning = "anchor_not_found_in_pre";
                    return false;
                }

                leftBookmark = preCodes[i];
                rightBookmark = i + 1 < preCodes.Count ? preCodes[i + 1] : null;
                return true;
            }

            string postAnchor = postSCodes[0];
            int j = preCodes.IndexOf(postAnchor);
            if (j < 0)
            {
                warning = "anchor_not_found_in_pre";
                return false;
            }

            leftBookmark = j > 0 ? preCodes[j - 1] : null;
            rightBookmark = preCodes[j];
            return true;
        }

        public static SnapshotDiffResult BuildBetweenBookmarks(
            SnapshotRecord pre,
            SnapshotRecord post,
            string leftBookmark,
            string rightBookmark)
        {
            var postCodes = ParseFlatSnapshot(post.Snapshot);
            if (postCodes.Count == 0 && (string.IsNullOrEmpty(leftBookmark) && string.IsNullOrEmpty(rightBookmark)))
            {
                return Unavailable(pre?.Index, post?.Index, "empty_interval");
            }

            int postStart;
            if (!string.IsNullOrEmpty(leftBookmark))
            {
                int leftIdx = postCodes.IndexOf(leftBookmark);
                if (leftIdx < 0)
                {
                    return Unavailable(pre?.Index, post?.Index, "neighbor_not_found_in_post");
                }

                postStart = leftIdx + 1;
            }
            else
            {
                postStart = 0;
            }

            int postEnd;
            if (!string.IsNullOrEmpty(rightBookmark))
            {
                int rightIdx = postCodes.IndexOf(rightBookmark);
                if (rightIdx < 0)
                {
                    return Unavailable(pre?.Index, post?.Index, "neighbor_not_found_in_post");
                }

                postEnd = rightIdx - 1;
            }
            else
            {
                postEnd = postCodes.Count - 1;
            }

            if (postStart > postEnd)
            {
                return Unavailable(pre?.Index, post?.Index, "empty_interval");
            }

            var sentences = new List<SentenceCodeText>();
            for (int i = postStart; i <= postEnd; i++)
            {
                string code = postCodes[i];
                sentences.Add(new SentenceCodeText
                {
                    Code = code,
                    Text = DocumentState.GetSentenceContent(code) ?? string.Empty
                });
            }

            var result = new SnapshotDiffResult
            {
                Sentences = sentences,
                Source = SourceVerified,
                PreIndex = pre?.Index,
                PostIndex = post?.Index
            };

            LogDiffResult(leftBookmark, rightBookmark, result);
            return result;
        }

        private static bool TryFindContinuousBlock(List<string> preCodes, List<string> anchorCodes, out int start, out int end)
        {
            start = -1;
            end = -1;
            if (preCodes == null || preCodes.Count == 0 || anchorCodes == null || anchorCodes.Count == 0)
            {
                return false;
            }

            int firstIdx = preCodes.IndexOf(anchorCodes[0]);
            if (firstIdx < 0)
            {
                return false;
            }

            for (int i = 0; i < anchorCodes.Count; i++)
            {
                int idx = firstIdx + i;
                if (idx >= preCodes.Count || !string.Equals(preCodes[idx], anchorCodes[i], StringComparison.Ordinal))
                {
                    return false;
                }
            }

            start = firstIdx;
            end = firstIdx + anchorCodes.Count - 1;
            return true;
        }

        private static bool ValidateBookmarksInPre(List<string> preCodes, string left, string right, out string warning)
        {
            warning = null;
            if (preCodes.IndexOf(left) < 0 || preCodes.IndexOf(right) < 0)
            {
                warning = "anchor_not_found_in_pre";
                return false;
            }

            return true;
        }

        private static List<string> FilterSentenceCodes(List<string> codes)
        {
            if (codes == null)
            {
                return new List<string>();
            }

            return codes
                .Where(c => !string.IsNullOrEmpty(c) && c.StartsWith("S_", StringComparison.Ordinal))
                .ToList();
        }

        private static bool HasTableOnlyAnchors(
            List<string> preAnchorCodes,
            List<string> postAnchorCodes,
            List<string> preSCodes,
            List<string> postSCodes)
        {
            bool hasTable = (preAnchorCodes ?? new List<string>()).Any(c => c.StartsWith("T_", StringComparison.Ordinal))
                || (postAnchorCodes ?? new List<string>()).Any(c => c.StartsWith("T_", StringComparison.Ordinal));
            return hasTable && preSCodes.Count == 0 && postSCodes.Count == 0;
        }

        private static SnapshotDiffResult Unavailable(int? preIndex, int? postIndex, string warning)
        {
            var result = new SnapshotDiffResult
            {
                Sentences = new List<SentenceCodeText>(),
                Source = SourceUnavailable,
                Warning = warning,
                PreIndex = preIndex,
                PostIndex = postIndex
            };
            System.Diagnostics.Debug.WriteLine(
                $"[new_sentences] unavailable reason={warning} pre_idx={preIndex} post_idx={postIndex}");
            return result;
        }

        private static void LogDiffResult(string leftBookmark, string rightBookmark, SnapshotDiffResult result)
        {
            string codes = string.Join(",", result.Sentences.Select(s => s.Code));
            System.Diagnostics.Debug.WriteLine(
                $"[new_sentences] source={result.Source} pre_idx={result.PreIndex} post_idx={result.PostIndex} " +
                $"left={leftBookmark ?? "(none)"} right={rightBookmark ?? "(none)"} → {codes}");
        }

        public sealed class BookmarkPair
        {
            public string LeftBookmark { get; set; }
            public string RightBookmark { get; set; }
        }
    }
}
