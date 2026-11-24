using System;
using System.Text;
using System.Text.RegularExpressions;

namespace WordAddIn1.DocumentMapping
{
    public static class AutonumConvert
    {
        public static readonly Regex Tag = new Regex(
            @"<autonum\b[^>]*/?>",
            RegexOptions.Compiled | RegexOptions.IgnoreCase);

        public static string MakeTag(string mark)
        {
            if (string.IsNullOrEmpty(mark))
            {
                return "";
            }

            return "<autonum v=\"" + EscapeAttr(mark) + "\"/>";
        }

        public static string StripTags(string text)
        {
            if (string.IsNullOrEmpty(text))
            {
                return text ?? "";
            }

            return Tag.Replace(text, "");
        }

        public static int SkipTag(string text, int i)
        {
            if (text == null || i < 0 || i >= text.Length || text[i] != '<')
            {
                return i;
            }

            if (i + 8 <= text.Length
                && string.Compare(text, i, "<autonum", 0, 8, StringComparison.OrdinalIgnoreCase) == 0)
            {
                int close = text.IndexOf('>', i);
                if (close >= 0)
                {
                    return close + 1;
                }
            }

            return i;
        }

        public static string ReplaceCodeInners(string display, Func<string, string> storyByCode)
        {
            if (string.IsNullOrEmpty(display) || storyByCode == null)
            {
                return display ?? "";
            }

            var sb = new StringBuilder();
            int i = 0;
            while (i < display.Length)
            {
                int open = display.IndexOf('[', i);
                if (open < 0)
                {
                    sb.Append(display.Substring(i));
                    break;
                }

                sb.Append(display.Substring(i, open - i));
                int close = display.IndexOf(']', open);
                if (close < 0)
                {
                    sb.Append(display.Substring(open));
                    break;
                }

                string tag = display.Substring(open, close - open + 1);
                string code;
                string endNeedle;
                if (TryStartTag(tag, out code, out endNeedle))
                {
                    int endAt = display.IndexOf(endNeedle, close + 1, StringComparison.Ordinal);
                    if (endAt < 0)
                    {
                        sb.Append(tag);
                        i = close + 1;
                        continue;
                    }

                    string story = storyByCode(code);
                    sb.Append(tag);
                    sb.Append(story ?? display.Substring(close + 1, endAt - (close + 1)));
                    sb.Append(endNeedle);
                    i = endAt + endNeedle.Length;
                    continue;
                }

                sb.Append(tag);
                i = close + 1;
            }

            return sb.ToString();
        }

        public static string EscapeAttr(string value)
        {
            if (string.IsNullOrEmpty(value))
            {
                return "";
            }

            return value
                .Replace("&", "&amp;")
                .Replace("\"", "&quot;")
                .Replace("<", "&lt;")
                .Replace(">", "&gt;");
        }

        private static bool TryStartTag(string tag, out string code, out string endNeedle)
        {
            code = null;
            endNeedle = null;
            if (string.IsNullOrEmpty(tag) || tag[0] != '[')
            {
                return false;
            }

            if (tag.StartsWith("[S_", StringComparison.Ordinal) && tag.IndexOf("-start", StringComparison.Ordinal) > 0)
            {
                int dash = tag.IndexOf("-start", StringComparison.Ordinal);
                code = tag.Substring(1, dash - 1);
                endNeedle = "[" + code + "-end]";
                return true;
            }

            if (tag.StartsWith("[P_", StringComparison.Ordinal) && tag.EndsWith(" start]", StringComparison.Ordinal))
            {
                code = tag.Substring(1, tag.Length - " start]".Length - 1);
                endNeedle = "[" + code + " end]";
                return true;
            }

            return false;
        }
    }
}
