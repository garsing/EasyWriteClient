using System;
using System.Collections.Generic;

namespace WordAddIn1.SpreadsheetHost
{
    internal static class SpreadsheetSheetNames
    {
        private static readonly char[] Illegal =
        {
            '\\', '/', '?', '*', '[', ']', ':'
        };

        public static bool TryValidateName(string name, out string error)
        {
            error = null;
            if (string.IsNullOrWhiteSpace(name))
            {
                error = "工作表名不能为空";
                return false;
            }

            string trimmed = name.Trim();
            if (trimmed.Length > 31)
            {
                error = "工作表名过长（最多 31 个字符）";
                return false;
            }

            if (trimmed.IndexOfAny(Illegal) >= 0)
            {
                error = "工作表名含非法字符（不能包含 \\ / ? * [ ] :）";
                return false;
            }

            if (trimmed.StartsWith("'", StringComparison.Ordinal)
                || trimmed.EndsWith("'", StringComparison.Ordinal))
            {
                error = "工作表名不能以单引号开头或结尾";
                return false;
            }

            return true;
        }
    }
}
