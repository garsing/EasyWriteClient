using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace WordAddIn1
{
    /// <summary>
    /// 在 display 全文上从目标表往前找直接父格。不访问 COM。
    /// </summary>
    public static class NestedTableParentFinder
    {
        public enum Kind
        {
            TopLevel,
            Nested,
            Failed
        }

        public sealed class Result
        {
            public Kind Kind;
            public string ParentTableId;
            public int Row;
            public int Col;
            public string Error;

            public string GeometricCode
            {
                get
                {
                    if (Kind != Kind.Nested || string.IsNullOrEmpty(ParentTableId))
                    {
                        return "";
                    }

                    return ParentTableId + "_r" + Row + "_c" + Col;
                }
            }
        }

        public static Result Find(string display, string tableId)
        {
            if (string.IsNullOrEmpty(display) || string.IsNullOrEmpty(tableId))
            {
                return Fail("display 中没有目标表");
            }

            if (!TryFindTableOpen(display, tableId, out int targetLt))
            {
                return Fail("display 中没有目标表");
            }

            int cursor = targetLt - 1;
            int parentTd = -1;
            int attrRow = 0;
            int attrCol = 0;
            bool hasRc = false;
            while (cursor >= 0)
            {
                int lt = PreviousLt(display, cursor);
                if (lt < 0)
                {
                    break;
                }

                if (IsCloseTable(display, lt))
                {
                    if (!TrySkipClosedTable(display, lt, out cursor))
                    {
                        return Fail("嵌套表但拿不到父级单元格");
                    }

                    continue;
                }

                if (IsOpenTd(display, lt))
                {
                    parentTd = lt;
                    hasRc = TryReadRowCol(OpenTag(display, lt), out attrRow, out attrCol);
                    cursor = lt - 1;
                    break;
                }

                cursor = lt - 1;
            }

            if (parentTd < 0)
            {
                return new Result { Kind = Kind.TopLevel };
            }

            int parentOpen = -1;
            string parentId = null;
            while (cursor >= 0)
            {
                int lt = PreviousLt(display, cursor);
                if (lt < 0)
                {
                    break;
                }

                if (IsCloseTable(display, lt))
                {
                    if (!TrySkipClosedTable(display, lt, out cursor))
                    {
                        return Fail("嵌套表但拿不到父级单元格");
                    }

                    continue;
                }

                if (IsOpenTable(display, lt))
                {
                    parentOpen = lt;
                    parentId = ReadId(OpenTag(display, lt));
                    break;
                }

                cursor = lt - 1;
            }

            if (parentOpen < 0 || string.IsNullOrEmpty(parentId))
            {
                return Fail("嵌套表但拿不到父级单元格");
            }

            int row;
            int col;
            if (hasRc)
            {
                row = attrRow;
                col = attrCol;
            }
            else if (!TryCountCell(display, parentOpen, parentTd, out row, out col))
            {
                return Fail("嵌套表但拿不到父级单元格");
            }

            return new Result
            {
                Kind = Kind.Nested,
                ParentTableId = parentId,
                Row = row,
                Col = col
            };
        }

        private static Result Fail(string error)
        {
            return new Result { Kind = Kind.Failed, Error = error };
        }

        private static bool TryFindTableOpen(string display, string tableId, out int tagStart)
        {
            tagStart = -1;
            string needle = "id=\"" + tableId + "\"";
            int search = 0;
            while (search < display.Length)
            {
                int idAt = display.IndexOf(needle, search, StringComparison.OrdinalIgnoreCase);
                if (idAt < 0)
                {
                    return false;
                }

                int lt = display.LastIndexOf('<', idAt);
                if (lt >= 0 && IsOpenTable(display, lt))
                {
                    tagStart = lt;
                    return true;
                }

                search = idAt + needle.Length;
            }

            return false;
        }

        private static bool TrySkipClosedTable(string text, int closeLt, out int cursor)
        {
            int depth = 1;
            int i = closeLt - 1;
            while (i >= 0)
            {
                int lt = PreviousLt(text, i);
                if (lt < 0)
                {
                    break;
                }

                if (IsCloseTable(text, lt))
                {
                    depth++;
                    i = lt - 1;
                    continue;
                }

                if (IsOpenTable(text, lt))
                {
                    depth--;
                    if (depth == 0)
                    {
                        cursor = lt - 1;
                        return true;
                    }

                    i = lt - 1;
                    continue;
                }

                i = lt - 1;
            }

            cursor = -1;
            return false;
        }

        private static bool TryCountCell(string display, int parentOpen, int parentTd, out int row, out int col)
        {
            row = 0;
            col = 0;
            int openEnd = display.IndexOf('>', parentOpen);
            if (openEnd < 0 || openEnd >= parentTd)
            {
                return false;
            }

            var remain = new List<int>();
            int depth = 0;
            int logicalRow = 0;
            int logicalCol = 1;
            int i = openEnd + 1;
            while (i < display.Length && i <= parentTd)
            {
                if (StartsWith(display, i, "<table"))
                {
                    depth++;
                    i += 6;
                    continue;
                }

                if (StartsWith(display, i, "</table"))
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

                if (StartsWith(display, i, "<tr"))
                {
                    char after = i + 3 < display.Length ? display[i + 3] : ' ';
                    if (after == '>' || char.IsWhiteSpace(after) || after == '/')
                    {
                        int trEnd = display.IndexOf('>', i);
                        if (trEnd < 0)
                        {
                            return false;
                        }

                        logicalRow++;
                        logicalCol = 1;
                        i = trEnd + 1;
                        continue;
                    }
                }

                if (StartsWith(display, i, "</tr"))
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

                if (IsOpenTd(display, i))
                {
                    int tdEnd = display.IndexOf('>', i);
                    if (tdEnd < 0)
                    {
                        return false;
                    }

                    string openTag = display.Substring(i, tdEnd - i + 1);
                    int colspan = ReadSpan(openTag, "colspan");
                    int rowspan = ReadSpan(openTag, "rowspan");
                    while (logicalCol <= remain.Count && remain[logicalCol - 1] > 0)
                    {
                        logicalCol++;
                    }

                    if (i == parentTd)
                    {
                        row = logicalRow;
                        col = logicalCol;
                        return row >= 1 && col >= 1;
                    }

                    while (remain.Count < logicalCol + colspan - 1)
                    {
                        remain.Add(0);
                    }

                    for (int c = 0; c < colspan; c++)
                    {
                        remain[logicalCol - 1 + c] = rowspan;
                    }

                    logicalCol += colspan;
                    i = tdEnd + 1;
                    continue;
                }

                i++;
            }

            return false;
        }

        private static bool TryReadRowCol(string openTag, out int row, out int col)
        {
            row = 0;
            col = 0;
            if (!TryAttr(openTag, "r", out row) || !TryAttr(openTag, "c", out col))
            {
                return false;
            }

            return row >= 1 && col >= 1;
        }

        private static bool TryAttr(string tag, string name, out int value)
        {
            value = 0;
            if (string.IsNullOrEmpty(tag))
            {
                return false;
            }

            Match match = Regex.Match(
                tag,
                @"\b" + name + @"\s*=\s*[""']?(\d+)",
                RegexOptions.IgnoreCase);
            return match.Success && int.TryParse(match.Groups[1].Value, out value);
        }

        private static int ReadSpan(string tag, string name)
        {
            if (!TryAttr(tag, name, out int n) || n < 1)
            {
                return 1;
            }

            return n;
        }

        private static string ReadId(string openTag)
        {
            if (string.IsNullOrEmpty(openTag))
            {
                return null;
            }

            Match match = Regex.Match(openTag, @"\bid\s*=\s*""([^""]*)""", RegexOptions.IgnoreCase);
            if (!match.Success || string.IsNullOrEmpty(match.Groups[1].Value))
            {
                return null;
            }

            return match.Groups[1].Value;
        }

        private static string OpenTag(string text, int lt)
        {
            int end = text.IndexOf('>', lt);
            if (end < 0)
            {
                return "";
            }

            return text.Substring(lt, end - lt + 1);
        }

        private static int PreviousLt(string text, int fromInclusive)
        {
            for (int i = fromInclusive; i >= 0; i--)
            {
                if (text[i] == '<')
                {
                    return i;
                }
            }

            return -1;
        }

        private static bool IsCloseTable(string text, int i)
        {
            return StartsWith(text, i, "</table");
        }

        private static bool IsOpenTable(string text, int i)
        {
            if (!StartsWith(text, i, "<table"))
            {
                return false;
            }

            char after = i + 6 < text.Length ? text[i + 6] : ' ';
            return after == '>' || char.IsWhiteSpace(after) || after == '/';
        }

        private static bool IsOpenTd(string text, int i)
        {
            if (!StartsWith(text, i, "<td"))
            {
                return false;
            }

            char after = i + 3 < text.Length ? text[i + 3] : ' ';
            return after == '>' || char.IsWhiteSpace(after) || after == '/';
        }

        private static bool StartsWith(string text, int i, string prefix)
        {
            if (i < 0 || i + prefix.Length > text.Length)
            {
                return false;
            }

            return string.Compare(text, i, prefix, 0, prefix.Length, StringComparison.OrdinalIgnoreCase) == 0;
        }
    }
}
