using System;
using System.Collections.Generic;
using System.Security.Cryptography;
using System.Text.RegularExpressions;

namespace WordAddIn1.SpreadsheetHost
{
    internal static class SpreadsheetImageStampHelper
    {
        public const string StampHead = "ew:";

        private const string Alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789";
        private static readonly Regex StampRegex = new Regex(
            @"^ew:I_([A-Za-z0-9]{8})",
            RegexOptions.Compiled);

        public static bool TryParse(string altText, out string imageId, out string userSuffix)
        {
            imageId = null;
            userSuffix = "";
            if (string.IsNullOrEmpty(altText))
            {
                return false;
            }

            Match m = StampRegex.Match(altText);
            if (!m.Success)
            {
                return false;
            }

            imageId = "I_" + m.Groups[1].Value;
            string rest = altText.Substring(m.Length);
            if (rest.StartsWith(" ", StringComparison.Ordinal))
            {
                rest = rest.Substring(1);
            }

            userSuffix = rest;
            return true;
        }

        public static string Format(string imageId, string userSuffix)
        {
            string stamp = StampHead + imageId;
            if (string.IsNullOrEmpty(userSuffix))
            {
                return stamp;
            }

            return stamp + " " + userSuffix;
        }

        public static string NewRandomImageId(ISet<string> occupied)
        {
            if (occupied == null)
            {
                occupied = new HashSet<string>(StringComparer.Ordinal);
            }

            byte[] buf = new byte[8];
            for (int attempt = 0; attempt < 32; attempt++)
            {
                using (var rng = RandomNumberGenerator.Create())
                {
                    rng.GetBytes(buf);
                }

                char[] chars = new char[8];
                for (int i = 0; i < 8; i++)
                {
                    chars[i] = Alphabet[buf[i] % Alphabet.Length];
                }

                string id = "I_" + new string(chars);
                if (occupied.Add(id))
                {
                    return id;
                }
            }

            return null;
        }
    }
}
