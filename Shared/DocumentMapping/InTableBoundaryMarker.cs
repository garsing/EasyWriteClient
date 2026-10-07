using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace WordAddIn1.DocumentMapping
{
    /// <summary>
    /// 每张表 display 顺序的最后一个 S_/P_，在 [*-start] 后标 [in_table=T_xxx]。
    /// 不改编码括号；定位仍抄 S_/P_。
    /// </summary>
    public static class InTableBoundaryMarker
    {
        public const string MarkerPrefix = "[in_table=";

        private static readonly Regex TableStartRegex = new Regex(
            @"<table\s+id=""([^""]+)""[^>]*>",
            RegexOptions.Compiled | RegexOptions.IgnoreCase);

        private static readonly Regex TableOpenTagRegex = new Regex(
            @"<table\b",
            RegexOptions.Compiled | RegexOptions.IgnoreCase);

        private static readonly Regex SentenceStartRegex = new Regex(
            @"\[(S_[0-9a-zA-Z]{5})-start\]",
            RegexOptions.Compiled);

        private static readonly Regex ParagraphStartRegex = new Regex(
            @"\[(P_[0-9a-zA-Z]{5}(?:_[1-9][0-9]*)?)(?:-start| start)\]",
            RegexOptions.Compiled);

        private static readonly Regex InTableMarkerRegex = new Regex(
            @"\[in_table=[^\]]+\]",
            RegexOptions.Compiled);

        private const string CloseTableTag = "</table>";

        public static string AnnotateSentences(string display)
        {
            return Annotate(display, SentenceStartRegex);
        }

        public static string AnnotateParagraphs(string display)
        {
            return Annotate(display, ParagraphStartRegex);
        }

        public static string Strip(string text)
        {
            if (string.IsNullOrEmpty(text))
            {
                return text;
            }

            return InTableMarkerRegex.Replace(text, "");
        }

        private static string Annotate(string display, Regex codeStartRegex)
        {
            if (string.IsNullOrEmpty(display))
            {
                return display;
            }

            var inserts = new List<Tuple<int, string>>();
            foreach (Match startMatch in TableStartRegex.Matches(display))
            {
                string tableId = startMatch.Groups[1].Value;
                if (string.IsNullOrEmpty(tableId))
                {
                    continue;
                }

                int blockStart = startMatch.Index;
                int blockEnd = FindMatchingCloseTag(display, blockStart);
                if (blockEnd < 0)
                {
                    continue;
                }

                int lastStartEnd = FindLastOuterCodeStartEnd(
                    display,
                    startMatch.Index + startMatch.Length,
                    blockEnd,
                    codeStartRegex);
                if (lastStartEnd < 0)
                {
                    continue;
                }

                if (HasInTableMarkerAt(display, lastStartEnd))
                {
                    continue;
                }

                inserts.Add(Tuple.Create(lastStartEnd, MarkerPrefix + tableId + "]"));
            }

            if (inserts.Count == 0)
            {
                return display;
            }

            inserts.Sort((a, b) => b.Item1.CompareTo(a.Item1));
            string result = display;
            for (int i = 0; i < inserts.Count; i++)
            {
                Tuple<int, string> item = inserts[i];
                result = result.Insert(item.Item1, item.Item2);
            }

            return result;
        }

        private static bool HasInTableMarkerAt(string display, int index)
        {
            if (index < 0 || index >= display.Length)
            {
                return false;
            }

            return string.Compare(
                display,
                index,
                MarkerPrefix,
                0,
                MarkerPrefix.Length,
                StringComparison.Ordinal) == 0;
        }

        private static int FindLastOuterCodeStartEnd(
            string text,
            int contentStart,
            int blockEnd,
            Regex codeStartRegex)
        {
            int lastEnd = -1;
            int nestedDepth = 0;
            int i = contentStart;

            while (i < blockEnd)
            {
                if (TryMatchTag(text, i, "<table", out int openLen))
                {
                    nestedDepth++;
                    i += openLen;
                    continue;
                }

                if (TryMatchTag(text, i, CloseTableTag, out int closeLen))
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
                    Match codeMatch = codeStartRegex.Match(text, i);
                    if (codeMatch.Success && codeMatch.Index == i)
                    {
                        lastEnd = codeMatch.Index + codeMatch.Length;
                        i = lastEnd;
                        continue;
                    }
                }

                i++;
            }

            return lastEnd;
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
                    && string.Compare(
                        text,
                        i,
                        CloseTableTag,
                        0,
                        CloseTableTag.Length,
                        StringComparison.OrdinalIgnoreCase) == 0)
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
