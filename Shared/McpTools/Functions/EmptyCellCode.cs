using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    /// <summary>
    /// 格几何码：T_表号_r行_c列。真空槽与序列歧义格盖戳、解析、剥标记集中在此。
    /// </summary>
    public static class EmptyCellCode
    {
        public const string MixError = "一格一条，禁止与 S_ 混排";
        public const string NoFormatError = "empty_cell_no_format: 空格码禁止取格式，写入请用默认宋体/12";
        public const string TableIdNotCellError = "table 须为光杆表号，不能用空格码";

        public static readonly Regex Pattern = new Regex(
            @"^T_[0-9A-Za-z]+_r[0-9]+_c[0-9]+$",
            RegexOptions.Compiled);

        public static readonly Regex MarkerPattern = new Regex(
            @"\[T_[0-9A-Za-z]+_r[0-9]+_c[0-9]+-(?:start|end)\]",
            RegexOptions.Compiled);

        private static readonly Regex SentenceStartInCell = new Regex(
            @"\[S_[0-9a-zA-Z]{5}-start(?: [^\]]*)?\]",
            RegexOptions.Compiled);

        private static readonly Regex SentenceStartCapture = new Regex(
            @"\[(S_[0-9a-zA-Z]{5})-start(?: [^\]]*)?\]",
            RegexOptions.Compiled);

        private static readonly Regex SentenceEndInCell = new Regex(
            @"\[S_[0-9a-zA-Z]{5}-end\]",
            RegexOptions.Compiled);

        private static readonly Regex ParagraphMarkerInCell = new Regex(
            @"\[P_[0-9A-Za-z_]+(?: start| end)\]",
            RegexOptions.Compiled);

        private static readonly Regex TableOpenWithId = new Regex(
            @"<table\s+id=""([^""]+)""",
            RegexOptions.Compiled | RegexOptions.IgnoreCase);

        public sealed class Parsed
        {
            public string TableId { get; set; }
            public int Row { get; set; }
            public int Col { get; set; }
            public string Code { get; set; }
        }

        public static bool IsEmptyCellCode(string code)
        {
            return !string.IsNullOrEmpty(code) && Pattern.IsMatch(code.Trim());
        }

        public static bool TryParse(string code, out Parsed parsed)
        {
            parsed = null;
            if (string.IsNullOrWhiteSpace(code))
            {
                return false;
            }

            string trimmed = code.Trim();
            Match m = Pattern.Match(trimmed);
            if (!m.Success)
            {
                return false;
            }

            int rAt = trimmed.LastIndexOf("_r", StringComparison.Ordinal);
            int cAt = trimmed.LastIndexOf("_c", StringComparison.Ordinal);
            if (rAt < 2 || cAt <= rAt)
            {
                return false;
            }

            string tableId = trimmed.Substring(0, rAt);
            if (!int.TryParse(trimmed.Substring(rAt + 2, cAt - (rAt + 2)), out int row)
                || !int.TryParse(trimmed.Substring(cAt + 2), out int col)
                || row < 1
                || col < 1)
            {
                return false;
            }

            parsed = new Parsed
            {
                TableId = tableId,
                Row = row,
                Col = col,
                Code = trimmed
            };
            return true;
        }

        public static string Format(string tableId, int row, int col)
        {
            return (tableId ?? "") + "_r" + row + "_c" + col;
        }

        public static string MarkerPair(string code)
        {
            return "[" + code + "-start][" + code + "-end]";
        }

        public static string StripMarkers(string text)
        {
            if (string.IsNullOrEmpty(text))
            {
                return text;
            }

            return MarkerPattern.Replace(text, "");
        }

        public static bool CsvContainsAny(string csv)
        {
            if (string.IsNullOrWhiteSpace(csv))
            {
                return false;
            }

            foreach (string part in csv.Split(','))
            {
                if (IsEmptyCellCode(part))
                {
                    return true;
                }
            }

            return false;
        }

        /// <summary>
        /// 按 display HTML 网格盖戳。行 = 顶层 &lt;tr&gt; 序号（1-based），
        /// 列 = 本行逻辑列（跳过 rowspan 占用，加上 colspan），与 WordReader 可见格 / TableRowSlot 同一套。
        /// 先真空插入，再按 SegSnapshot 把歧义格换成几何码。不 COM 读格。
        /// </summary>
        public static string Stamp(string display)
        {
            string segSnapshot = "";
            try
            {
                segSnapshot = DocumentState.SegSnapshot ?? "";
            }
            catch
            {
                segSnapshot = "";
            }

            return Stamp(display, segSnapshot);
        }

        public static string Stamp(string display, string segSnapshot)
        {
            return StampVacuum(display);
        }

        public static string Stamp(Word.Document doc, string display)
        {
            return Stamp(display);
        }

        public static void StampProcessingResult(Word.Document doc, WordDocumentExtractor.ProcessingResult result)
        {
            if (result == null)
            {
                return;
            }

            if (result.ChunkInfo != null)
            {
                foreach (var chunk in result.ChunkInfo)
                {
                    if (chunk == null)
                    {
                        continue;
                    }

                    chunk.DisplayContent = Stamp(chunk.DisplayContent);
                    if (!string.IsNullOrEmpty(chunk.ParagraphDisplayContent))
                    {
                        chunk.ParagraphDisplayContent = Stamp(chunk.ParagraphDisplayContent);
                    }
                }
            }

            IReadOnlyList<string> paras = DocumentState.ChunkParagraphDisplayContents;
            if (paras != null && paras.Count > 0)
            {
                var stamped = new List<string>(paras.Count);
                for (int i = 0; i < paras.Count; i++)
                {
                    stamped.Add(Stamp(paras[i]));
                }

                DocumentState.SetChunkParagraphDisplayContents(stamped);
            }
        }

        private static string StampVacuum(string display)
        {
            if (string.IsNullOrEmpty(display) || display.IndexOf("<table", StringComparison.OrdinalIgnoreCase) < 0)
            {
                return display;
            }

            var opens = new List<Match>();
            foreach (Match m in TableOpenWithId.Matches(display))
            {
                opens.Add(m);
            }

            if (opens.Count == 0)
            {
                return display;
            }

            string result = display;
            for (int i = opens.Count - 1; i >= 0; i--)
            {
                string tableId = opens[i].Groups[1].Value;
                if (string.IsNullOrEmpty(tableId) || IsEmptyCellCode(tableId))
                {
                    continue;
                }

                result = StampOneTable(result, tableId);
            }

            return result;
        }

        /// <summary>
        /// 可见格内 S_ 序列与某组原码一致则整格换成几何码（编码消歧 D8）。
        /// </summary>
        public static string WrapCellsMatchingCodeSets(
            string display,
            IReadOnlyList<IReadOnlyList<string>> codeSets)
        {
            if (string.IsNullOrEmpty(display)
                || codeSets == null
                || codeSets.Count == 0
                || display.IndexOf("<table", StringComparison.OrdinalIgnoreCase) < 0)
            {
                return display;
            }

            var needles = new List<List<string>>();
            foreach (IReadOnlyList<string> set in codeSets)
            {
                if (set != null && set.Count > 0)
                {
                    needles.Add(new List<string>(set));
                }
            }

            if (needles.Count == 0)
            {
                return display;
            }

            return ReplaceMatchingCells(display, needles);
        }

        private static string ReplaceMatchingCells(string display, List<List<string>> needles)
        {
            if (string.IsNullOrEmpty(display)
                || display.IndexOf("<table", StringComparison.OrdinalIgnoreCase) < 0)
            {
                return display;
            }

            var opens = new List<Match>();
            foreach (Match m in TableOpenWithId.Matches(display))
            {
                opens.Add(m);
            }

            if (opens.Count == 0)
            {
                return display;
            }

            var jobs = new List<ReplaceJob>();
            foreach (Match open in opens)
            {
                string tableId = open.Groups[1].Value;
                if (string.IsNullOrEmpty(tableId) || IsEmptyCellCode(tableId))
                {
                    continue;
                }

                if (!TryFindTableInner(display, tableId, out int innerStart, out int innerEnd))
                {
                    continue;
                }

                string inner = display.Substring(innerStart, innerEnd - innerStart);
                List<TdSpan> tds = CollectVisibleGridCells(inner);
                foreach (TdSpan td in tds)
                {
                    string body = inner.Substring(td.ContentStart, td.ContentEnd - td.ContentStart);
                    if (ContainsNestedTable(body))
                    {
                        continue;
                    }

                    List<string> codes = ExtractSentenceCodes(body);
                    if (codes.Count == 0 || !CodeSetMatches(needles, codes))
                    {
                        continue;
                    }

                    string code = Format(tableId, td.Row, td.Col);
                    jobs.Add(new ReplaceJob
                    {
                        Start = innerStart + td.ContentStart,
                        End = innerStart + td.ContentEnd,
                        Text = WrapAmbiguousCell(body, code)
                    });
                }
            }

            return ApplyReplacesFromBack(display, jobs);
        }

        private static bool CodeSetMatches(List<List<string>> needles, List<string> codes)
        {
            foreach (List<string> needle in needles)
            {
                if (needle.Count != codes.Count)
                {
                    continue;
                }

                bool same = true;
                for (int i = 0; i < codes.Count; i++)
                {
                    if (needle[i] != codes[i])
                    {
                        same = false;
                        break;
                    }
                }

                if (same)
                {
                    return true;
                }
            }

            return false;
        }

        private static List<string> ExtractSentenceCodes(string body)
        {
            var codes = new List<string>();
            if (string.IsNullOrEmpty(body))
            {
                return codes;
            }

            foreach (Match m in SentenceStartCapture.Matches(body))
            {
                codes.Add(m.Groups[1].Value);
            }

            return codes;
        }

        private static bool ContainsNestedTable(string body)
        {
            return !string.IsNullOrEmpty(body)
                && body.IndexOf("<table", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private static string WrapAmbiguousCell(string body, string code)
        {
            string text = body ?? "";
            text = SentenceStartInCell.Replace(text, "");
            text = SentenceEndInCell.Replace(text, "");
            text = ParagraphMarkerInCell.Replace(text, "");
            text = MarkerPattern.Replace(text, "");
            if (string.IsNullOrEmpty(Regex.Replace(text, @"\s+", "")))
            {
                return MarkerPair(code);
            }

            return "[" + code + "-start]" + text + "[" + code + "-end]";
        }

        private static string ApplyReplacesFromBack(string display, List<ReplaceJob> jobs)
        {
            if (jobs == null || jobs.Count == 0)
            {
                return display;
            }

            jobs.Sort((a, b) => b.Start.CompareTo(a.Start));
            string result = display;
            foreach (ReplaceJob job in jobs)
            {
                result = result.Substring(0, job.Start) + job.Text + result.Substring(job.End);
            }

            return result;
        }

        private sealed class ReplaceJob
        {
            public int Start;
            public int End;
            public string Text;
        }

        private static string StampOneTable(string display, string tableId)
        {
            if (!TryFindTableInner(display, tableId, out int innerStart, out int innerEnd))
            {
                return display;
            }

            string inner = display.Substring(innerStart, innerEnd - innerStart);
            List<TdSpan> tds = CollectVisibleGridCells(inner);
            var inserts = new List<Tuple<int, string>>();
            foreach (TdSpan td in tds)
            {
                string body = inner.Substring(td.ContentStart, td.ContentEnd - td.ContentStart);
                if (SentenceStartInCell.IsMatch(body) || MarkerPattern.IsMatch(body))
                {
                    continue;
                }

                if (!IsVisuallyEmptyTd(body))
                {
                    continue;
                }

                string code = Format(tableId, td.Row, td.Col);
                inserts.Add(Tuple.Create(innerStart + td.ContentStart, MarkerPair(code)));
            }

            return ApplyInsertsFromBack(display, inserts);
        }

        private static bool IsVisuallyEmptyTd(string body)
        {
            if (string.IsNullOrEmpty(body))
            {
                return true;
            }

            if (body.IndexOf("<table", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return false;
            }

            string stripped = SentenceStartInCell.Replace(body, "");
            stripped = MarkerPattern.Replace(stripped, "");
            stripped = Regex.Replace(stripped, @"\s+", "");
            return stripped.Length == 0;
        }

        private static string ApplyInsertsFromBack(string display, List<Tuple<int, string>> inserts)
        {
            if (inserts == null || inserts.Count == 0)
            {
                return display;
            }

            inserts.Sort((a, b) => b.Item1.CompareTo(a.Item1));
            string result = display;
            foreach (var item in inserts)
            {
                result = result.Insert(item.Item1, item.Item2);
            }

            return result;
        }

        private static bool TryFindTableInner(string display, string tableId, out int innerStart, out int innerEnd)
        {
            innerStart = 0;
            innerEnd = 0;
            string needle = "id=\"" + tableId + "\"";
            int search = 0;
            while (search < display.Length)
            {
                int idAt = IndexOfIgnoreCase(display, needle, search);
                if (idAt < 0)
                {
                    return false;
                }

                int tagStart = display.LastIndexOf('<', idAt);
                if (tagStart < 0)
                {
                    return false;
                }

                if (!StartsWithIgnoreCase(display, tagStart, "<table"))
                {
                    search = idAt + needle.Length;
                    continue;
                }

                int openEnd = display.IndexOf('>', idAt);
                if (openEnd < 0)
                {
                    return false;
                }

                int close = FindMatchingTableClose(display, openEnd + 1);
                if (close < 0)
                {
                    return false;
                }

                innerStart = openEnd + 1;
                innerEnd = close;
                return true;
            }

            return false;
        }

        private static int FindMatchingTableClose(string text, int from)
        {
            int depth = 1;
            int i = from;
            while (i < text.Length)
            {
                if (StartsWithIgnoreCase(text, i, "<table"))
                {
                    depth++;
                    i += 6;
                    continue;
                }

                if (StartsWithIgnoreCase(text, i, "</table"))
                {
                    depth--;
                    if (depth == 0)
                    {
                        return i;
                    }

                    i += 7;
                    continue;
                }

                i++;
            }

            return -1;
        }

        private sealed class TdSpan
        {
            public int Row;
            public int Col;
            public int ContentStart;
            public int ContentEnd;
        }

        /// <summary>
        /// 顶层可见格：&lt;tr&gt; 计行，rowspan 占用列跳过，colspan 占宽。嵌套表不计入外表。
        /// </summary>
        private static List<TdSpan> CollectVisibleGridCells(string inner)
        {
            var list = new List<TdSpan>();
            var remain = new List<int>();
            int depth = 0;
            int row = 0;
            int col = 1;
            int i = 0;
            while (i < inner.Length)
            {
                if (StartsWithIgnoreCase(inner, i, "<table"))
                {
                    depth++;
                    i += 6;
                    continue;
                }

                if (StartsWithIgnoreCase(inner, i, "</table"))
                {
                    if (depth > 0)
                    {
                        depth--;
                    }

                    i += 7;
                    continue;
                }

                if (depth != 0)
                {
                    i++;
                    continue;
                }

                if (StartsWithIgnoreCase(inner, i, "<tr"))
                {
                    char afterTr = i + 3 < inner.Length ? inner[i + 3] : ' ';
                    if (afterTr == '>' || char.IsWhiteSpace(afterTr) || afterTr == '/')
                    {
                        row++;
                        col = 1;
                    }

                    i += 3;
                    continue;
                }

                if (StartsWithIgnoreCase(inner, i, "</tr"))
                {
                    for (int c = 0; c < remain.Count; c++)
                    {
                        if (remain[c] > 0)
                        {
                            remain[c]--;
                        }
                    }

                    i += 4;
                    continue;
                }

                if (StartsWithIgnoreCase(inner, i, "<td") || StartsWithIgnoreCase(inner, i, "<th"))
                {
                    char after = i + 3 < inner.Length ? inner[i + 3] : ' ';
                    if (after != '>' && !char.IsWhiteSpace(after) && after != '/')
                    {
                        i++;
                        continue;
                    }

                    int openEnd = inner.IndexOf('>', i);
                    if (openEnd < 0)
                    {
                        break;
                    }

                    string openTag = inner.Substring(i, openEnd - i + 1);
                    int colspan = ReadSpanAttr(openTag, "colspan");
                    int rowspan = ReadSpanAttr(openTag, "rowspan");
                    while (col <= remain.Count && remain[col - 1] > 0)
                    {
                        col++;
                    }

                    int close = FindMatchingCellClose(inner, openEnd + 1);
                    if (close < 0)
                    {
                        break;
                    }

                    list.Add(new TdSpan
                    {
                        Row = row,
                        Col = col,
                        ContentStart = openEnd + 1,
                        ContentEnd = close
                    });

                    EnsureRemainSize(remain, col + colspan - 1);
                    for (int c = 0; c < colspan; c++)
                    {
                        remain[col - 1 + c] = rowspan;
                    }

                    col += colspan;
                    i = close;
                    continue;
                }

                i++;
            }

            return list;
        }

        private static void EnsureRemainSize(List<int> remain, int col)
        {
            while (remain.Count < col)
            {
                remain.Add(0);
            }
        }

        private static int ReadSpanAttr(string tag, string name)
        {
            if (string.IsNullOrEmpty(tag) || string.IsNullOrEmpty(name))
            {
                return 1;
            }

            Match m = Regex.Match(
                tag,
                name + @"\s*=\s*[""']?(\d+)[""']?",
                RegexOptions.IgnoreCase);
            if (!m.Success)
            {
                return 1;
            }

            int n;
            if (!int.TryParse(m.Groups[1].Value, out n) || n < 1)
            {
                return 1;
            }

            return n;
        }

        private static int FindMatchingCellClose(string text, int from)
        {
            int depth = 1;
            int i = from;
            while (i < text.Length)
            {
                if (StartsWithIgnoreCase(text, i, "<td") || StartsWithIgnoreCase(text, i, "<th"))
                {
                    char after = i + 3 < text.Length ? text[i + 3] : ' ';
                    if (after == '>' || char.IsWhiteSpace(after))
                    {
                        depth++;
                    }

                    i += 3;
                    continue;
                }

                if (StartsWithIgnoreCase(text, i, "</td") || StartsWithIgnoreCase(text, i, "</th"))
                {
                    depth--;
                    if (depth == 0)
                    {
                        return i;
                    }

                    i += 4;
                    continue;
                }

                i++;
            }

            return -1;
        }

        private static bool StartsWithIgnoreCase(string text, int index, string value)
        {
            if (index < 0 || index + value.Length > text.Length)
            {
                return false;
            }

            return string.Compare(text, index, value, 0, value.Length, StringComparison.OrdinalIgnoreCase) == 0;
        }

        private static int IndexOfIgnoreCase(string text, string value, int start)
        {
            return text.IndexOf(value, start, StringComparison.OrdinalIgnoreCase);
        }
    }
}
