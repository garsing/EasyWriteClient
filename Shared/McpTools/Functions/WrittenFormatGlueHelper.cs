using System;
using System.Collections.Generic;

namespace WordAddIn1
{
    /// <summary>
    /// 写入套格式：按 `\r` 切前缀/中间/后缀；用 SegSnapshot + 存文往前推定粘前。无 COM。
    /// </summary>
    public static class WrittenFormatGlueHelper
    {
        public struct BreakSlices
        {
            public int FirstCr;
            public int LastCr;
            public string Prefix;
            public string Middle;
            public string Suffix;
        }

        public static BreakSlices SplitByOuterBreaks(string text)
        {
            var slices = new BreakSlices
            {
                FirstCr = -1,
                LastCr = -1,
                Prefix = null,
                Middle = null,
                Suffix = null
            };

            text = text ?? "";
            int first = text.IndexOf('\r');
            if (first < 0)
            {
                slices.Suffix = text;
                return slices;
            }

            int last = text.LastIndexOf('\r');
            slices.FirstCr = first;
            slices.LastCr = last;
            slices.Prefix = text.Substring(0, first + 1);
            if (last > first)
            {
                slices.Middle = text.Substring(first + 1, last - first);
            }

            if (last < text.Length - 1)
            {
                slices.Suffix = text.Substring(last + 1);
            }

            return slices;
        }

        public static bool PreviousRunEndsWithBreak(string previousText)
        {
            return !string.IsNullOrEmpty(previousText)
                && previousText[previousText.Length - 1] == '\r';
        }

        public static bool CanMapPayloadToComText(string payload, string comText)
        {
            payload = payload ?? "";
            comText = comText ?? "";
            if (string.Equals(comText, payload, StringComparison.Ordinal))
            {
                return true;
            }

            if (comText.Length != payload.Length)
            {
                return false;
            }

            int first = payload.IndexOf('\r');
            int last = payload.LastIndexOf('\r');
            if (first >= 0 && comText[first] != '\r')
            {
                return false;
            }

            if (last >= 0 && comText[last] != '\r')
            {
                return false;
            }

            return first >= 0 || last >= 0;
        }

        public static bool IsPrefixSafe(string anchorCodes, bool insertAfter)
        {
            return IsPrefixSafe(
                anchorCodes,
                insertAfter,
                DocumentState.Snapshot,
                DocumentState.SegSnapshot,
                DocumentState.GetSentenceContent);
        }

        public static bool IsPrefixSafe(
            string anchorCodes,
            bool insertAfter,
            string snapshot,
            string segSnapshot,
            Func<string, string> getStored)
        {
            if (getStored == null)
            {
                return true;
            }

            List<string> names = ParseSentenceCodes(anchorCodes);
            if (names.Count == 0)
            {
                return true;
            }

            if (EmptyCellCode.IsEmptyCellCode(names[0]))
            {
                return true;
            }

            List<string> ordered = DocumentState.GetOrderedSentenceCodesFromSnapshot(snapshot);
            if (ordered == null || ordered.Count == 0)
            {
                return true;
            }

            int[] segs = DocumentState.BuildSegIndexBySnapshotPosition(snapshot, segSnapshot);
            int start = FindSequenceStart(ordered, names);
            if (start < 0)
            {
                return true;
            }

            int walkFrom;
            if (insertAfter)
            {
                walkFrom = start + names.Count - 1;
            }
            else
            {
                if (start == 0)
                {
                    return true;
                }

                if (segs != null && start < segs.Length
                    && segs[start] != segs[start - 1])
                {
                    return true;
                }

                walkFrom = start - 1;
            }

            return WalkBackSameSeg(ordered, segs, walkFrom, getStored);
        }

        private static bool WalkBackSameSeg(
            List<string> ordered,
            int[] segs,
            int fromIndex,
            Func<string, string> getStored)
        {
            if (fromIndex < 0 || fromIndex >= ordered.Count)
            {
                return true;
            }

            int seg = segs != null && fromIndex < segs.Length ? segs[fromIndex] : 0;
            for (int i = fromIndex; i >= 0; i--)
            {
                int thisSeg = segs != null && i < segs.Length ? segs[i] : 0;
                if (thisSeg != seg)
                {
                    return true;
                }

                string stored = getStored(ordered[i]) ?? "";
                int kind = ClassifyTail(stored);
                if (kind == 1)
                {
                    return true;
                }

                if (kind == 2)
                {
                    return false;
                }
            }

            return true;
        }

        /// <returns>0 空继续；1 安全（`\r`）；2 粘前</returns>
        public static int ClassifyTail(string stored)
        {
            if (string.IsNullOrEmpty(stored))
            {
                return 0;
            }

            int i = stored.Length - 1;
            if (stored[i] == '>' && IsImageTagEnd(stored, i))
            {
                return 2;
            }

            if (stored[i] == '\r')
            {
                return 1;
            }

            return 2;
        }

        private static bool IsImageTagEnd(string text, int closeIndex)
        {
            int open = text.LastIndexOf('<', closeIndex);
            if (open < 0)
            {
                return false;
            }

            int len = closeIndex - open + 1;
            if (len < 6)
            {
                return false;
            }

            return string.Compare(text, open, "<image", 0, 6, StringComparison.OrdinalIgnoreCase) == 0;
        }

        private static int FindSequenceStart(List<string> ordered, List<string> needle)
        {
            if (needle.Count == 0 || ordered.Count < needle.Count)
            {
                return -1;
            }

            for (int i = 0; i <= ordered.Count - needle.Count; i++)
            {
                bool match = true;
                for (int j = 0; j < needle.Count; j++)
                {
                    if (!string.Equals(ordered[i + j], needle[j], StringComparison.Ordinal))
                    {
                        match = false;
                        break;
                    }
                }

                if (match)
                {
                    return i;
                }
            }

            return ordered.IndexOf(needle[0]);
        }

        private static List<string> ParseSentenceCodes(string codes)
        {
            var names = new List<string>();
            if (string.IsNullOrWhiteSpace(codes))
            {
                return names;
            }

            foreach (string part in codes.Split(','))
            {
                string t = part.Trim();
                if (t.Length == 0)
                {
                    continue;
                }

                names.Add(t);
            }

            return names;
        }
    }
}
