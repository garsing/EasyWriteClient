using System;
using System.Globalization;
using System.Text.RegularExpressions;

namespace WordAddIn1.PresentationHost
{
    /// <summary>B1：约定 HTML 颜色读写（data-fill / data-font-color）。</summary>
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
    }
}
