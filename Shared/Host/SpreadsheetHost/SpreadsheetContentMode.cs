namespace WordAddIn1.SpreadsheetHost
{
    /// <summary>
    /// 读区域格内容模式：显示值与公式二选一（I14）。
    /// </summary>
    internal enum SpreadsheetContentMode
    {
        Value = 0,
        Formula = 1
    }

    internal static class SpreadsheetContentModeUtil
    {
        public static bool TryParse(string raw, out SpreadsheetContentMode mode, out string error)
        {
            mode = SpreadsheetContentMode.Value;
            error = null;
            if (string.IsNullOrWhiteSpace(raw))
            {
                return true;
            }

            string s = raw.Trim().ToLowerInvariant();
            if (s == "value" || s == "display")
            {
                mode = SpreadsheetContentMode.Value;
                return true;
            }

            if (s == "formula")
            {
                mode = SpreadsheetContentMode.Formula;
                return true;
            }

            error = "content 须为 value 或 formula";
            return false;
        }

        public static string ToWire(SpreadsheetContentMode mode)
        {
            return mode == SpreadsheetContentMode.Formula ? "formula" : "value";
        }

        /// <summary>
        /// formula 模式且有公式串 → 用公式；否则用显示值。
        /// </summary>
        public static string ResolveText(string displayText, string formulaOrNull, SpreadsheetContentMode mode)
        {
            if (mode == SpreadsheetContentMode.Formula
                && !string.IsNullOrEmpty(formulaOrNull))
            {
                return formulaOrNull;
            }

            return displayText ?? "";
        }
    }
}
