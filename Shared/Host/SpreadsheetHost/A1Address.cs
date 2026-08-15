using System;
using System.Text;

namespace WordAddIn1.SpreadsheetHost
{
    internal static class A1Address
    {
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
    }
}
