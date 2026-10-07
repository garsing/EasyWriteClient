using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using WordAddIn1;
using WordAddIn1.DocumentMapping.CodeResolve;

namespace WordAddIn1.DocumentMapping
{
    /// <summary>
    /// 先把 display 转成对齐字流，再按字符序把每个 P_ 落到 Paragraphs 段号。
    /// 当场算，不写进长期映射。
    /// </summary>
    public static class DisplayParagraphIndexMap
    {
        private static readonly Regex PBlock = new Regex(
            @"\[(P_[0-9a-zA-Z]{5}(?:_[1-9][0-9]*)?)(?:-start| start)\](.*?)\[\1(?:-end| end)\]",
            RegexOptions.Compiled | RegexOptions.Singleline);

        public sealed class Hit
        {
            public string Code { get; set; }
            public int ParaNo { get; set; }
            public string Inner { get; set; }
        }

        public static string CombinedParagraphDisplay()
        {
            IReadOnlyList<string> chunks = DocumentState.ChunkParagraphDisplayContents;
            if (chunks == null || chunks.Count == 0)
            {
                return "";
            }

            return string.Concat(chunks);
        }

        public static List<Hit> Build(string display)
        {
            var hits = new List<Hit>();
            if (string.IsNullOrEmpty(display))
            {
                return hits;
            }

            string marks = DisplayStoryConvert.ToMarks(display);
            List<string> pieces = DisplayStoryConvert.SplitByParaEnds(marks);
            int[] starts = new int[pieces.Count + 1];
            int pos = 0;
            for (int i = 0; i < pieces.Count; i++)
            {
                starts[i] = pos;
                pos += pieces[i].Length;
            }

            starts[pieces.Count] = pos;

            int cursor = 0;
            foreach (Match m in PBlock.Matches(display))
            {
                string code = m.Groups[1].Value;
                string inner = DisplayStoryConvert.StripNeedle(m.Groups[2].Value ?? "");
                var hit = new Hit { Code = code, Inner = inner, ParaNo = 0 };
                hits.Add(hit);

                int idx;
                if (inner.Length == 0)
                {
                    idx = cursor;
                }
                else
                {
                    idx = marks.IndexOf(inner, cursor, StringComparison.Ordinal);
                    if (idx < 0)
                    {
                        continue;
                    }
                }

                int piece = PieceAt(starts, idx, marks.Length);
                if (piece < 0)
                {
                    continue;
                }

                hit.ParaNo = piece + 1;
                cursor = starts[piece + 1];
            }

            return hits;
        }

        public static int ResolveParaIndex(string paragraphCode)
        {
            return ResolveParaIndex(CombinedParagraphDisplay(), paragraphCode);
        }

        public static int ResolveParaIndex(string display, string paragraphCode)
        {
            if (!ParagraphPublicCode.TryParse(paragraphCode, out ParagraphPublicCode.Parsed parsed))
            {
                return 0;
            }

            List<Hit> hits = Build(display);
            int paraNo = 0;
            int n = 0;
            for (int i = 0; i < hits.Count; i++)
            {
                if (!ParagraphPublicCode.TryParse(hits[i].Code, out ParagraphPublicCode.Parsed hitParsed)
                    || !string.Equals(hitParsed.BaseCode, parsed.BaseCode, StringComparison.Ordinal))
                {
                    continue;
                }

                n++;
                if (!parsed.RequiresUnique && n == parsed.Occurrence)
                {
                    return hits[i].ParaNo;
                }

                paraNo = hits[i].ParaNo;
            }

            if (parsed.RequiresUnique && n == 1)
            {
                return paraNo;
            }

            return 0;
        }

        private static int PieceAt(int[] starts, int idx, int markLen)
        {
            if (starts.Length <= 1)
            {
                return -1;
            }

            if (idx >= markLen)
            {
                return -1;
            }

            for (int i = 0; i < starts.Length - 1; i++)
            {
                int a = starts[i];
                int b = starts[i + 1];
                if (idx >= a && idx < b)
                {
                    return i;
                }

                if (idx == a && a == b)
                {
                    return i;
                }
            }

            if (idx == markLen && starts.Length >= 2)
            {
                return starts.Length - 2;
            }

            return -1;
        }
    }
}
