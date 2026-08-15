using System;
using System.Text;
using System.Text.RegularExpressions;

namespace WordAddIn1.SpreadsheetHost
{
    internal static class A1Address
    {
        private static readonly Regex CellPattern = new Regex(
            @"^\s*\$?([A-Za-z]{1,3})\$?(\d{1,7})\s*$",
            RegexOptions.Compiled);

        public static string ColumnLetter(int column)
        {
            if (column <= 0)
            {
                return "";
            }

            var sb = new StringBuilder();
            int n = column;
            while (n > 0)
            {
                n--;
                sb.Insert(0, (char)('A' + (n % 26)));
                n /= 26;
            }

            return sb.ToString();
        }

        public static bool TryParseColumnLetter(string letters, out int column)
        {
            column = 0;
            if (string.IsNullOrWhiteSpace(letters))
            {
                return false;
            }

            letters = letters.Trim().ToUpperInvariant();
            if (letters.Length > 3)
            {
                return false;
            }

            int value = 0;
            for (int i = 0; i < letters.Length; i++)
            {
                char ch = letters[i];
                if (ch < 'A' || ch > 'Z')
                {
                    return false;
                }

                value = value * 26 + (ch - 'A' + 1);
            }

            if (value <= 0)
            {
                return false;
            }

            column = value;
            return true;
        }

        public static string Cell(int row, int column)
        {
            if (row <= 0 || column <= 0)
            {
                return "";
            }

            return ColumnLetter(column) + row;
        }

        public static string Range(int firstRow, int firstCol, int lastRow, int lastCol)
        {
            string a = Cell(firstRow, firstCol);
            string b = Cell(lastRow, lastCol);
            if (string.IsNullOrEmpty(a) || string.IsNullOrEmpty(b))
            {
                return "";
            }

            return string.Equals(a, b, StringComparison.Ordinal) ? a : a + ":" + b;
        }

        public static bool TryParseCell(string a1, out int row, out int column)
        {
            row = 0;
            column = 0;
            if (string.IsNullOrWhiteSpace(a1))
            {
                return false;
            }

            Match m = CellPattern.Match(a1.Trim());
            if (!m.Success)
            {
                return false;
            }

            if (!TryParseColumnLetter(m.Groups[1].Value, out column))
            {
                return false;
            }

            if (!int.TryParse(m.Groups[2].Value, out row) || row <= 0)
            {
                return false;
            }

            return true;
        }

        /// <summary>
        /// 解析纯 A1：C3 或 A1:G40。禁止含 !。
        /// </summary>
        public static bool TryParseRange(
            string a1,
            out int firstRow,
            out int firstCol,
            out int lastRow,
            out int lastCol,
            out string error)
        {
            firstRow = firstCol = lastRow = lastCol = 0;
            error = null;
            if (string.IsNullOrWhiteSpace(a1))
            {
                error = "非法 range: （空）";
                return false;
            }

            string raw = a1.Trim();
            if (raw.IndexOf('!') >= 0)
            {
                error = "range 须为纯 A1（如 A1:G40），表名请用 sheet 参数";
                return false;
            }

            int colon = raw.IndexOf(':');
            if (colon < 0)
            {
                if (!TryParseCell(raw, out firstRow, out firstCol))
                {
                    error = "非法 range: " + raw;
                    return false;
                }

                lastRow = firstRow;
                lastCol = firstCol;
                return true;
            }

            string left = raw.Substring(0, colon);
            string right = raw.Substring(colon + 1);
            if (!TryParseCell(left, out firstRow, out firstCol)
                || !TryParseCell(right, out lastRow, out lastCol))
            {
                error = "非法 range: " + raw;
                return false;
            }

            if (lastRow < firstRow || lastCol < firstCol)
            {
                error = "非法 range: 起止颠倒 " + raw;
                return false;
            }

            return true;
        }
    }
}
