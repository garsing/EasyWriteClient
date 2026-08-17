using System;
using System.Globalization;
using System.Text.RegularExpressions;

namespace WordAddIn1.PresentationHost
{
    /// <summary>约定 HTML 样式读写（B1 色 / B3 字号粗体 z / B3b 线）。</summary>
    internal static class PptHtmlStyleIo
    {
        private static readonly Regex HexColor = new Regex(
            @"^#([0-9A-Fa-f]{6})$",
            RegexOptions.CultureInvariant | RegexOptions.Compiled);

        private static readonly Regex StyleColor = new Regex(
            @"color\s*:\s*(#[0-9A-Fa-f]{6}|#[0-9A-Fa-f]{3})\b",
            RegexOptions.IgnoreCase | RegexOptions.CultureInvariant | RegexOptions.Compiled);

        /// <summary>规范化为 #RRGGBB 或 none；非法返回 false。</summary>
        public static bool TryNormalizeFillOrColor(string raw, bool allowNone, out string normalized, out string error)
        {
            normalized = null;
            error = null;
            if (string.IsNullOrWhiteSpace(raw))
            {
                error = "颜色值为空";
                return false;
            }

            string s = raw.Trim();
            if (allowNone && string.Equals(s, "none", StringComparison.OrdinalIgnoreCase))
            {
                normalized = "none";
                return true;
            }

            if (s.Length == 4 && s[0] == '#')
            {
                // #RGB → #RRGGBB
                s = "#" + s[1] + s[1] + s[2] + s[2] + s[3] + s[3];
            }

            Match m = HexColor.Match(s);
            if (!m.Success)
            {
                error = "非法颜色（须为 #RRGGBB" + (allowNone ? " 或 none" : "") + "）: " + raw;
                return false;
            }

            normalized = "#" + m.Groups[1].Value.ToUpperInvariant();
            return true;
        }

        /// <summary>从 style 中取 color:#…（无则 null）。</summary>
        public static string TryExtractStyleColor(string style)
        {
            if (string.IsNullOrWhiteSpace(style))
            {
                return null;
            }

            Match m = StyleColor.Match(style);
            if (!m.Success)
            {
                return null;
            }

            if (!TryNormalizeFillOrColor(m.Groups[1].Value, allowNone: false, out string n, out _))
            {
                return null;
            }

            return n;
        }

        public static string FormatOfficeRgb(int officeRgb)
        {
            int r = officeRgb & 0xFF;
            int g = (officeRgb >> 8) & 0xFF;
            int b = (officeRgb >> 16) & 0xFF;
            return "#"
                + r.ToString("X2", CultureInfo.InvariantCulture)
                + g.ToString("X2", CultureInfo.InvariantCulture)
                + b.ToString("X2", CultureInfo.InvariantCulture);
        }

        public static bool TryParseHexToOfficeRgb(string hex, out int officeRgb, out string error)
        {
            officeRgb = 0;
            if (!TryNormalizeFillOrColor(hex, allowNone: false, out string n, out error))
            {
                return false;
            }

            int r = int.Parse(n.Substring(1, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture);
            int g = int.Parse(n.Substring(3, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture);
            int b = int.Parse(n.Substring(5, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture);
            officeRgb = r + (g << 8) + (b << 16);
            return true;
        }

        public static bool TryParseFontSizePt(string raw, out double pt, out string error)
        {
            pt = 0;
            error = null;
            if (string.IsNullOrWhiteSpace(raw))
            {
                error = "data-font-size 为空";
                return false;
            }

            string s = raw.Trim();
            if (s.EndsWith("pt", StringComparison.OrdinalIgnoreCase))
            {
                s = s.Substring(0, s.Length - 2).Trim();
            }

            if (!double.TryParse(s, NumberStyles.Float, CultureInfo.InvariantCulture, out pt)
                || double.IsNaN(pt)
                || double.IsInfinity(pt)
                || pt <= 0
                || pt > 400)
            {
                error = "非法 data-font-size（须为 0–400 的 pt 数值）: " + raw;
                return false;
            }

            return true;
        }

        public static bool TryParseFontBold(string raw, out bool bold, out string error)
        {
            bold = false;
            error = null;
            if (string.IsNullOrWhiteSpace(raw))
            {
                error = "data-font-bold 为空";
                return false;
            }

            string s = raw.Trim();
            if (string.Equals(s, "true", StringComparison.OrdinalIgnoreCase)
                || s == "1")
            {
                bold = true;
                return true;
            }

            if (string.Equals(s, "false", StringComparison.OrdinalIgnoreCase)
                || s == "0")
            {
                bold = false;
                return true;
            }

            error = "非法 data-font-bold（须为 true/false）: " + raw;
            return false;
        }

        public static bool TryParseZ(string raw, out int z, out string error)
        {
            z = 0;
            error = null;
            if (string.IsNullOrWhiteSpace(raw))
            {
                error = "data-z 为空";
                return false;
            }

            if (!int.TryParse(raw.Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out z)
                || z < 0)
            {
                error = "非法 data-z（须为非负整数）: " + raw;
                return false;
            }

            return true;
        }

        public static bool TryParseLineWidthPt(string raw, out double pt, out string error)
        {
            pt = 0;
            error = null;
            if (string.IsNullOrWhiteSpace(raw))
            {
                error = "data-line-width 为空";
                return false;
            }

            string s = raw.Trim();
            if (s.EndsWith("pt", StringComparison.OrdinalIgnoreCase))
            {
                s = s.Substring(0, s.Length - 2).Trim();
            }

            if (!double.TryParse(s, NumberStyles.Float, CultureInfo.InvariantCulture, out pt)
                || double.IsNaN(pt)
                || double.IsInfinity(pt)
                || pt < 0
                || pt > 100)
            {
                error = "非法 data-line-width（须为 0–100 的 pt 数值）: " + raw;
                return false;
            }

            return true;
        }

        public static string FormatPt(double pt)
        {
            return pt.ToString("0.##", CultureInfo.InvariantCulture);
        }
    }
}
