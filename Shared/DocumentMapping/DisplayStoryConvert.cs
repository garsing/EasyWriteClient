using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;

namespace WordAddIn1.DocumentMapping
{
    /// <summary>
    /// paragraph display → 与 Content.Text 对齐的字流。
    /// 表在这一步化成 \r\a；映射时只按字符序走。
    /// </summary>
    public static class DisplayStoryConvert
    {
        private static readonly Regex CodeStart = new Regex(
            @"\[(?:P|S)_[^\]]+?(?:-start| start)[^\]]*\]", RegexOptions.Compiled);
        private static readonly Regex CodeEnd = new Regex(
            @"\[(?:P|S)_[^\]]+?(?:-end| end)\]", RegexOptions.Compiled);
        private static readonly Regex InTable = new Regex(
            @"\[in_table=[^\]]+\]", RegexOptions.Compiled);
        private static readonly Regex Image = new Regex(
            @"<image\b[^>]*/?>", RegexOptions.Compiled | RegexOptions.IgnoreCase);
        private static readonly Regex Chart = new Regex(
            @"<Chart\b[^>]*>[\s\S]*?</Chart\s*>|<Chart\b[^>]*/>",
            RegexOptions.Compiled | RegexOptions.IgnoreCase);
        private static readonly Regex AttrInt = new Regex(
            @"(colspan|rowspan)\s*=\s*""?(\d+)""?",
            RegexOptions.Compiled | RegexOptions.IgnoreCase);

        public static string ToMarks(string display)
        {
            if (string.IsNullOrEmpty(display))
            {
                return "";
            }

            return ConvertTables(StripWrappers(display)).Replace("\n", "\x0B");
        }

        public static string ToStory(string display)
        {
            return ToMarks(display).Replace("^b", "\f");
        }

        public static string StripWrappers(string display)
        {
            if (string.IsNullOrEmpty(display))
            {
                return "";
            }

            string s = CodeStart.Replace(display, "");
            s = CodeEnd.Replace(s, "");
            s = InTable.Replace(s, "");
            s = Chart.Replace(s, "");
            s = Image.Replace(s, "");
            return s;
        }

        public static string StripNeedle(string inner)
        {
            if (string.IsNullOrEmpty(inner))
            {
                return "";
            }

            string s = InTable.Replace(inner, "");
            s = Chart.Replace(s, "");
            s = Image.Replace(s, "");
            return s.Replace("\n", "\x0B");
        }

        public static List<string> SplitByParaEnds(string marks)
        {
            var pieces = new List<string>();
            if (string.IsNullOrEmpty(marks))
            {
                return pieces;
            }

            var sb = new StringBuilder();
            for (int i = 0; i < marks.Length; i++)
            {
                if (marks[i] == '\r')
                {
                    sb.Append('\r');
                    if (i + 1 < marks.Length && marks[i + 1] == '\a')
                    {
                        sb.Append('\a');
                        i++;
                    }

                    pieces.Add(sb.ToString());
                    sb.Length = 0;
                    continue;
                }

                if (marks[i] == '^'
                    && i + 1 < marks.Length
                    && marks[i + 1] == 'b')
                {
                    sb.Append('^').Append('b');
                    pieces.Add(sb.ToString());
                    sb.Length = 0;
                    i++;
                    continue;
                }

                sb.Append(marks[i]);
            }

            if (sb.Length > 0)
            {
                pieces.Add(sb.ToString());
            }

            return pieces;
        }

        private static string ConvertTables(string text)
        {
            var sb = new StringBuilder();
            int i = 0;
            while (i < text.Length)
            {
                int tableAt = IndexOfIgnoreCase(text, "<table", i);
                if (tableAt < 0)
                {
                    sb.Append(text.Substring(i));
                    break;
                }

                sb.Append(text.Substring(i, tableAt - i));
                int openEnd = text.IndexOf('>', tableAt);
                if (openEnd < 0)
                {
                    sb.Append(text.Substring(tableAt));
                    break;
                }

                int close = FindMatchingClose(text, openEnd + 1, "<table", "</table>");
                if (close < 0)
                {
                    sb.Append(text.Substring(tableAt));
                    break;
                }

                string inner = text.Substring(openEnd + 1, close - (openEnd + 1));
                sb.Append(ConvertOneTable(inner));
                int after = text.IndexOf('>', close);
                i = after >= 0 ? after + 1 : close + 8;
            }

            return sb.ToString();
        }

        private static string ConvertOneTable(string inner)
        {
            var rows = ParseRows(inner);
            if (rows.Count == 0)
            {
                return "";
            }

            int extra = 0;
            foreach (List<DispCell> row in rows)
            {
                foreach (DispCell cell in row)
                {
                    extra = Math.Max(extra, cell.RowSpan - 1);
                }
            }

            int nrows = rows.Count + extra;
            var carry = new int[nrows];
            for (int r = 0; r < rows.Count; r++)
            {
                int used = 0;
                foreach (DispCell cell in rows[r])
                {
                    used += cell.ColSpan;
                    for (int dr = 1; dr < cell.RowSpan && r + dr < nrows; dr++)
                    {
                        carry[r + dr] += cell.ColSpan;
                    }
                }

                carry[r] += used;
            }

            int ncols = 0;
            for (int r = 0; r < nrows; r++)
            {
                if (carry[r] > ncols)
                {
                    ncols = carry[r];
                }
            }

            var origin = new int[nrows, ncols];
            for (int r = 0; r < nrows; r++)
            {
                for (int c = 0; c < ncols; c++)
                {
                    origin[r, c] = -1;
                }
            }

            for (int r = 0; r < rows.Count; r++)
            {
                int c = 0;
                foreach (DispCell cell in rows[r])
                {
                    while (c < ncols && origin[r, c] != -1)
                    {
                        c++;
                    }

                    if (c < ncols)
                    {
                        origin[r, c] = 0;
                    }

                    for (int dc = 1; dc < cell.ColSpan && c + dc < ncols; dc++)
                    {
                        origin[r, c + dc] = -2;
                    }

                    for (int dr = 1; dr < cell.RowSpan && r + dr < nrows; dr++)
                    {
                        if (c < ncols)
                        {
                            origin[r + dr, c] = cell.ColSpan;
                        }

                        for (int dc = 1; dc < cell.ColSpan && c + dc < ncols; dc++)
                        {
                            origin[r + dr, c + dc] = -2;
                        }
                    }

                    c += cell.ColSpan;
                }
            }

            var sb = new StringBuilder();
            for (int r = 0; r < rows.Count; r++)
            {
                int c = 0;
                int cellIdx = 0;
                while (c < ncols)
                {
                    if (origin[r, c] > 0)
                    {
                        sb.Append('\r').Append('\a');
                        c += origin[r, c];
                        continue;
                    }

                    if (origin[r, c] == -2)
                    {
                        c++;
                        continue;
                    }

                    if (cellIdx < rows[r].Count)
                    {
                        DispCell cell = rows[r][cellIdx++];
                        sb.Append(ConvertTables(cell.Inner));
                        sb.Append('\r').Append('\a');
                        c += cell.ColSpan;
                    }
                    else
                    {
                        c++;
                    }
                }

                sb.Append('\r').Append('\a');
            }

            return sb.ToString();
        }

        private sealed class DispCell
        {
            public int ColSpan = 1;
            public int RowSpan = 1;
            public string Inner = "";
        }

        private static List<List<DispCell>> ParseRows(string inner)
        {
            var rows = new List<List<DispCell>>();
            int i = 0;
            while (i < inner.Length)
            {
                int tr = IndexOfIgnoreCase(inner, "<tr", i);
                if (tr < 0)
                {
                    break;
                }

                int trEnd = inner.IndexOf('>', tr);
                if (trEnd < 0)
                {
                    break;
                }

                int trClose = FindMatchingClose(inner, trEnd + 1, "<tr", "</tr>");
                if (trClose < 0)
                {
                    break;
                }

                string rowHtml = inner.Substring(trEnd + 1, trClose - (trEnd + 1));
                rows.Add(ParseCells(rowHtml));
                int after = inner.IndexOf('>', trClose);
                i = after >= 0 ? after + 1 : trClose + 5;
            }

            return rows;
        }

        private static List<DispCell> ParseCells(string rowHtml)
        {
            var cells = new List<DispCell>();
            int i = 0;
            while (i < rowHtml.Length)
            {
                int td = IndexOfIgnoreCase(rowHtml, "<td", i);
                int th = IndexOfIgnoreCase(rowHtml, "<th", i);
                int open = MinPos(td, th);
                if (open < 0)
                {
                    break;
                }

                bool isTh = th >= 0 && (td < 0 || th < td);
                int openEnd = rowHtml.IndexOf('>', open);
                if (openEnd < 0)
                {
                    break;
                }

                string tag = rowHtml.Substring(open, openEnd - open);
                int close = FindMatchingClose(
                    rowHtml,
                    openEnd + 1,
                    isTh ? "<th" : "<td",
                    isTh ? "</th>" : "</td>");
                if (close < 0)
                {
                    break;
                }

                cells.Add(new DispCell
                {
                    Inner = rowHtml.Substring(openEnd + 1, close - (openEnd + 1)),
                    ColSpan = ReadSpan(tag, "colspan"),
                    RowSpan = ReadSpan(tag, "rowspan")
                });
                int after = rowHtml.IndexOf('>', close);
                i = after >= 0 ? after + 1 : close + 5;
            }

            return cells;
        }

        private static int ReadSpan(string tag, string name)
        {
            Match m = AttrInt.Match(tag);
            while (m.Success)
            {
                if (string.Equals(m.Groups[1].Value, name, StringComparison.OrdinalIgnoreCase))
                {
                    int n;
                    if (int.TryParse(m.Groups[2].Value, out n) && n > 0)
                    {
                        return n;
                    }
                }

                m = m.NextMatch();
            }

            return 1;
        }

        private static int FindMatchingClose(string text, int from, string open, string close)
        {
            int depth = 1;
            int i = from;
            while (i < text.Length && depth > 0)
            {
                int nextOpen = IndexOfIgnoreCase(text, open, i);
                int nextClose = IndexOfIgnoreCase(text, close, i);
                if (nextClose < 0)
                {
                    return -1;
                }

                if (nextOpen >= 0 && nextOpen < nextClose)
                {
                    depth++;
                    i = nextOpen + open.Length;
                }
                else
                {
                    depth--;
                    if (depth == 0)
                    {
                        return nextClose;
                    }

                    i = nextClose + close.Length;
                }
            }

            return -1;
        }

        private static int IndexOfIgnoreCase(string text, string value, int start)
        {
            return text.IndexOf(value, start, StringComparison.OrdinalIgnoreCase);
        }

        private static int MinPos(int a, int b)
        {
            if (a < 0)
            {
                return b;
            }

            if (b < 0)
            {
                return a;
            }

            return Math.Min(a, b);
        }
    }
}
