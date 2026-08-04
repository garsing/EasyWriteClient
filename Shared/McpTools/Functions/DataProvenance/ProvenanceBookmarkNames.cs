using System;
using System.Security.Cryptography;
using System.Text;

namespace WordAddIn1
{
    public static class ProvenanceBookmarkNames
    {
        private const string Prefix = "yw_prov_";
        private const int MaxBookmarkLength = 40;

        public static string Marker(int n) => $"{Prefix}mk_{n}";

        public static string RefTitle() => $"{Prefix}ref_title";

        public static string RefBlock() => $"{Prefix}ref_block";

        public static string RefLine(int n) => $"{Prefix}ref_{n}";

        public static string TableLine(string code, int k) =>
            TruncateBookmark($"{Prefix}tbl_{SanitizeCode(code)}_{k}");

        public static string TableBlock(string code) =>
            TruncateBookmark($"{Prefix}tbl_block_{SanitizeCode(code)}");

        public static string ChartLine(string code, int k) =>
            TruncateBookmark($"{Prefix}cht_{SanitizeCode(code)}_{k}");

        public static string ChartBlock(string code) =>
            TruncateBookmark($"{Prefix}cht_block_{SanitizeCode(code)}");

        public static string ImageLine(string code, int k) =>
            TruncateBookmark($"{Prefix}img_{SanitizeCode(code)}_{k}");

        public static string ImageBlock(string code) =>
            TruncateBookmark($"{Prefix}img_block_{SanitizeCode(code)}");

        public static bool IsProvenanceBookmark(string name) =>
            !string.IsNullOrEmpty(name)
            && name.StartsWith(Prefix, StringComparison.Ordinal);

        private static string SanitizeCode(string code)
        {
            if (string.IsNullOrEmpty(code))
            {
                return "x";
            }

            var sb = new StringBuilder();
            foreach (char c in code)
            {
                if (char.IsLetterOrDigit(c) || c == '_')
                {
                    sb.Append(c);
                }
            }

            return sb.Length > 0 ? sb.ToString() : "x";
        }

        private static string TruncateBookmark(string name)
        {
            if (string.IsNullOrEmpty(name) || name.Length <= MaxBookmarkLength)
            {
                return name;
            }

            int hashLen = 7;
            string hash = ShortHash(name);
            int keep = MaxBookmarkLength - hashLen;
            if (keep < Prefix.Length + 2)
            {
                return name.Substring(0, MaxBookmarkLength);
            }

            return name.Substring(0, keep) + hash;
        }

        private static string ShortHash(string input)
        {
            using (var sha1 = SHA1.Create())
            {
                byte[] bytes = sha1.ComputeHash(Encoding.UTF8.GetBytes(input ?? ""));
                return BitConverter.ToString(bytes, 0, 3).Replace("-", "").ToLowerInvariant();
            }
        }
    }
}
