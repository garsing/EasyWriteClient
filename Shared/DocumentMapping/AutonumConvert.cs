using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;
using WordAddIn1;

namespace WordAddIn1.DocumentMapping
{
    public static class AutonumConvert
    {
        public static readonly Regex Tag = new Regex(
            @"<autonum\b[^>]*/?>",
            RegexOptions.Compiled | RegexOptions.IgnoreCase);

        private static readonly Regex TagAttrV = new Regex(
            @"<autonum\b[^>]*\bv\s*=\s*""([^""]*)""",
            RegexOptions.Compiled | RegexOptions.IgnoreCase);

        private static readonly Regex DecimalDot = new Regex(@"^\d+\.\s?", RegexOptions.Compiled);
        private static readonly Regex LetterDot = new Regex(@"^[A-Za-z]\.\s?", RegexOptions.Compiled);
        private static readonly Regex LetterParen = new Regex(@"^[A-Za-z]\)\s?", RegexOptions.Compiled);
        private static readonly Regex Bullet = new Regex(
            @"^[\u2022\u00B7\u25E6\u25AA\u25AB\u25CB\u25CF\uF0B7o]\s?",
            RegexOptions.Compiled);

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

        public static string ToVisible(string tagged)
        {
            if (string.IsNullOrEmpty(tagged))
            {
                return tagged ?? "";
            }

            return Tag.Replace(tagged, m =>
            {
                Match attr = TagAttrV.Match(m.Value);
                string v = attr.Success ? UnescapeAttr(attr.Groups[1].Value) : "";
                return VisiblePrefix(v);
            });
        }

        public static string VisiblePrefix(string mark)
        {
            if (string.IsNullOrEmpty(mark))
            {
                return "";
            }

            if (char.IsWhiteSpace(mark[mark.Length - 1]))
            {
                return mark;
            }

            return mark + " ";
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

        public static string StripNewByOriginalTypes(IList<string> originalMarks, string newText)
        {
            if (string.IsNullOrEmpty(newText))
            {
                return newText ?? "";
            }

            var kinds = KindsFromMarks(originalMarks);
            if (kinds.Count == 0 && (originalMarks == null || originalMarks.Count == 0))
            {
                return newText;
            }

            string[] parts = newText.Split('\r');
            for (int i = 0; i < parts.Length; i++)
            {
                parts[i] = StripOnePrefix(parts[i], kinds, originalMarks);
            }

            return string.Join("\r", parts);
        }

        public static bool SegmentHasOriginalTypePrefix(string segment, IList<string> originalMarks)
        {
            if (string.IsNullOrEmpty(segment))
            {
                return false;
            }

            var kinds = KindsFromMarks(originalMarks);
            string stripped = StripOnePrefix(segment, kinds, originalMarks);
            return stripped != segment;
        }

        public static List<string> PrefixMarksFromPair(string displayInner, string content)
        {
            var marks = new List<string>();
            string[] d = (displayInner ?? "").Split('\r');
            string[] c = (content ?? "").Split('\r');
            int n = Math.Max(d.Length, c.Length);
            for (int i = 0; i < n; i++)
            {
                string ds = i < d.Length ? d[i] : "";
                string cs = i < c.Length ? c[i] : "";
                if (ds.Length > cs.Length && ds.EndsWith(cs, StringComparison.Ordinal))
                {
                    string prefix = ds.Substring(0, ds.Length - cs.Length);
                    if (prefix.Length > 0)
                    {
                        marks.Add(prefix.TrimEnd());
                    }
                }
            }

            return marks;
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

        public static string UnescapeAttr(string value)
        {
            if (string.IsNullOrEmpty(value))
            {
                return "";
            }

            return value
                .Replace("&quot;", "\"")
                .Replace("&lt;", "<")
                .Replace("&gt;", ">")
                .Replace("&amp;", "&");
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

        private enum Kind
        {
            DecimalDot,
            LetterDot,
            LetterParen,
            Bullet,
            Exact
        }

        private static HashSet<Kind> KindsFromMarks(IList<string> marks)
        {
            var set = new HashSet<Kind>();
            if (marks == null)
            {
                return set;
            }

            foreach (string m in marks)
            {
                set.Add(Classify(m));
            }

            return set;
        }

        private static Kind Classify(string mark)
        {
            string s = (mark ?? "").Trim();
            if (Regex.IsMatch(s, @"^\d+\.$"))
            {
                return Kind.DecimalDot;
            }

            if (Regex.IsMatch(s, @"^[A-Za-z]\.$"))
            {
                return Kind.LetterDot;
            }

            if (Regex.IsMatch(s, @"^[A-Za-z]\)$"))
            {
                return Kind.LetterParen;
            }

            if (Bullet.IsMatch(s) || (s.Length == 1 && !char.IsLetterOrDigit(s[0])))
            {
                return Kind.Bullet;
            }

            return Kind.Exact;
        }

        private static string StripOnePrefix(string seg, HashSet<Kind> kinds, IList<string> marks)
        {
            if (string.IsNullOrEmpty(seg) || kinds == null || kinds.Count == 0)
            {
                return seg ?? "";
            }

            if (kinds.Contains(Kind.DecimalDot) && DecimalDot.IsMatch(seg))
            {
                return DecimalDot.Replace(seg, "", 1);
            }

            if (kinds.Contains(Kind.LetterDot) && LetterDot.IsMatch(seg))
            {
                return LetterDot.Replace(seg, "", 1);
            }

            if (kinds.Contains(Kind.LetterParen) && LetterParen.IsMatch(seg))
            {
                return LetterParen.Replace(seg, "", 1);
            }

            if (kinds.Contains(Kind.Bullet) && Bullet.IsMatch(seg))
            {
                return Bullet.Replace(seg, "", 1);
            }

            if (kinds.Contains(Kind.Exact) && marks != null)
            {
                foreach (string m in marks)
                {
                    string p = VisiblePrefix((m ?? "").TrimEnd());
                    if (p.Length > 0 && seg.StartsWith(p, StringComparison.Ordinal))
                    {
                        return seg.Substring(p.Length);
                    }

                    string raw = (m ?? "").TrimEnd();
                    if (raw.Length > 0 && seg.StartsWith(raw, StringComparison.Ordinal))
                    {
                        int n = raw.Length;
                        if (n < seg.Length && seg[n] == ' ')
                        {
                            n++;
                        }

                        return seg.Substring(n);
                    }
                }
            }

            return seg;
        }
    }
}
