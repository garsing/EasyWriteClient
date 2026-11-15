using System;
using System.Collections.Generic;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    /// <summary>
    /// 主文档目录域钉 ewC_ 书签，并记下域内句子、段落码。切句标记不进正文。
    /// </summary>
    public static class TocFieldStamp
    {
        public const char Edge = '\uE001';
        public const string InteriorMessage = "这是目录域里面的编码，不能直接改。";

        private static readonly Regex BookmarkName = new Regex(
            @"^ewC_[A-Za-z0-9]{8,32}$",
            RegexOptions.Compiled | RegexOptions.CultureInvariant);

        private static readonly Regex TocInstr = new Regex(
            @"(^|[^A-Za-z0-9])TOC([^A-Za-z0-9]|$)",
            RegexOptions.Compiled | RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);

        private static readonly Regex SentenceBlock = new Regex(
            @"\[S_([0-9A-Za-z]{5})-start[^\]]*\][\s\S]*?\[S_\1-end\]",
            RegexOptions.Compiled);

        private static readonly object Gate = new object();
        private static List<string> _ids = new List<string>();
        private static readonly HashSet<string> Interior = new HashSet<string>(StringComparer.Ordinal);
        private static readonly Dictionary<string, string> SentenceId = new Dictionary<string, string>(StringComparer.Ordinal);
        private static int _fieldCursor;
        private static bool _inside;
        private static string _openId;

        public static bool IsTocInstruction(string text)
        {
            return !string.IsNullOrEmpty(text) && TocInstr.IsMatch(text);
        }

        public static void Clear()
        {
            lock (Gate)
            {
                _ids = new List<string>();
                Interior.Clear();
                SentenceId.Clear();
                _fieldCursor = 0;
                _inside = false;
                _openId = null;
            }
        }

        public static void ResetPairing()
        {
            lock (Gate)
            {
                Interior.Clear();
                SentenceId.Clear();
                _fieldCursor = 0;
                _inside = false;
                _openId = null;
            }
        }

        public static IReadOnlyList<string> Collect(Word.Document doc)
        {
            var ids = new List<string>();
            if (doc == null)
            {
                lock (Gate)
                {
                    _ids = ids;
                }

                return ids;
            }

            try
            {
                foreach (Word.Field field in doc.Fields)
                {
                    int type;
                    try
                    {
                        type = (int)field.Type;
                    }
                    catch (Exception)
                    {
                        continue;
                    }

                    if (type != (int)Word.WdFieldType.wdFieldTOC)
                    {
                        continue;
                    }

                    Word.Range point = field.Code.Duplicate;
                    point.Collapse(Word.WdCollapseDirection.wdCollapseStart);
                    string existing = FindBookmark(doc, point.Start);
                    if (!string.IsNullOrEmpty(existing))
                    {
                        ids.Add(existing);
                        continue;
                    }

                    string created = AddBookmark(doc, point);
                    if (!string.IsNullOrEmpty(created))
                    {
                        ids.Add(created);
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("[TocFieldStamp] " + ex.Message);
            }

            lock (Gate)
            {
                _ids = ids;
                _fieldCursor = 0;
                _inside = false;
                _openId = null;
            }

            return ids;
        }

        public static string FirstInteriorSentence()
        {
            lock (Gate)
            {
                foreach (KeyValuePair<string, string> pair in SentenceId)
                {
                    if (!string.IsNullOrEmpty(pair.Key))
                    {
                        return pair.Key;
                    }
                }
            }

            return null;
        }

        public static bool IsInterior(string codes)
        {
            if (string.IsNullOrWhiteSpace(codes))
            {
                return false;
            }

            lock (Gate)
            {
                foreach (string part in codes.Split(','))
                {
                    string code = part.Trim();
                    if (code.Length > 0 && Interior.Contains(code))
                    {
                        return true;
                    }
                }
            }

            return false;
        }

        public static void MarkSeg(string segText, IList<string> sentenceNames, IList<string> sentenceTexts)
        {
            if (string.IsNullOrEmpty(segText) || sentenceNames == null || sentenceTexts == null)
            {
                return;
            }

            int si = 0;
            int consumed = 0;
            lock (Gate)
            {
                for (int i = 0; i < segText.Length; i++)
                {
                    char ch = segText[i];
                    if (ch == Edge)
                    {
                        if (!_inside)
                        {
                            _inside = true;
                            _openId = _fieldCursor < _ids.Count ? _ids[_fieldCursor] : null;
                            _fieldCursor++;
                        }
                        else
                        {
                            _inside = false;
                            _openId = null;
                        }

                        continue;
                    }

                    if (si >= sentenceTexts.Count)
                    {
                        continue;
                    }

                    if (_inside && si < sentenceNames.Count && !string.IsNullOrEmpty(_openId))
                    {
                        string name = sentenceNames[si];
                        if (!string.IsNullOrEmpty(name))
                        {
                            Interior.Add(name);
                            SentenceId[name] = _openId;
                        }
                    }

                    consumed++;
                    if (consumed >= sentenceTexts[si].Length)
                    {
                        si++;
                        consumed = 0;
                    }
                }
            }
        }

        public static void Carry(IReadOnlyList<string> originals, string rewritten)
        {
            if (string.IsNullOrEmpty(rewritten) || originals == null || originals.Count == 0)
            {
                return;
            }

            lock (Gate)
            {
                if (SentenceId.ContainsKey(rewritten))
                {
                    return;
                }

                string id = null;
                foreach (string original in originals)
                {
                    if (string.IsNullOrEmpty(original) || !SentenceId.TryGetValue(original, out string found))
                    {
                        return;
                    }

                    if (id == null)
                    {
                        id = found;
                    }
                    else if (!string.Equals(id, found, StringComparison.Ordinal))
                    {
                        return;
                    }
                }

                Interior.Add(rewritten);
                SentenceId[rewritten] = id;
            }
        }

        public static void MarkCoveredParagraph(string paragraphName, IReadOnlyList<string> sentenceNames)
        {
            if (string.IsNullOrEmpty(paragraphName) || sentenceNames == null || sentenceNames.Count == 0)
            {
                return;
            }

            lock (Gate)
            {
                foreach (string sentence in sentenceNames)
                {
                    if (string.IsNullOrEmpty(sentence) || !Interior.Contains(sentence))
                    {
                        return;
                    }
                }

                Interior.Add(paragraphName);
            }
        }

        public static string CollapseDisplay(string display)
        {
            if (string.IsNullOrEmpty(display) || display.IndexOf("[S_", StringComparison.Ordinal) < 0)
            {
                return display ?? "";
            }

            var sb = new StringBuilder();
            int i = 0;
            string openId = null;
            while (i < display.Length)
            {
                Match m = SentenceBlock.Match(display, i);
        if (!m.Success || m.Index != i)
        {
            if (openId != null)
            {
                sb.Append("[目录域 ").Append(openId).Append(']');
                openId = null;
            }

            int next = NextSentence(display, i);
                    sb.Append(display, i, (next < 0 ? display.Length : next) - i);
                    if (next < 0)
                    {
                        break;
                    }

                    i = next;
                    openId = null;
                    continue;
                }

                string code = "S_" + m.Groups[1].Value;
                string id;
                bool interior;
                lock (Gate)
                {
                    interior = SentenceId.TryGetValue(code, out id);
                }

                if (!interior)
                {
                    if (openId != null)
                    {
                        sb.Append("[目录域 ").Append(openId).Append(']');
                        openId = null;
                    }

                    sb.Append(m.Value);
                    i = m.Index + m.Length;
                    continue;
                }

                if (openId != id)
                {
                    if (openId != null)
                    {
                        sb.Append("[目录域 ").Append(openId).Append(']');
                    }

                    openId = id;
                }

                i = m.Index + m.Length;
            }

            if (openId != null)
            {
                sb.Append("[目录域 ").Append(openId).Append(']');
            }

            return sb.ToString();
        }

        private static int NextSentence(string display, int from)
        {
            return display.IndexOf("[S_", from, StringComparison.Ordinal);
        }

        private static string FindBookmark(Word.Document doc, int start)
        {
            foreach (Word.Bookmark bookmark in doc.Bookmarks)
            {
                string name = bookmark.Name ?? "";
                if (!BookmarkName.IsMatch(name))
                {
                    continue;
                }

                Word.Range range = bookmark.Range;
                if (range != null && range.Start == start)
                {
                    return name;
                }
            }

            return null;
        }

        private static string AddBookmark(Word.Document doc, Word.Range point)
        {
            for (int attempt = 0; attempt < 8; attempt++)
            {
                string name = "ewC_" + RandomBody(8);
                if (doc.Bookmarks.Exists(name))
                {
                    continue;
                }

                object anchor = point;
                doc.Bookmarks.Add(name, ref anchor);
                return name;
            }

            return null;
        }

        private static string RandomBody(int length)
        {
            const string alphabet = "0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ";
            var bytes = new byte[length];
            using (RandomNumberGenerator rng = RandomNumberGenerator.Create())
            {
                rng.GetBytes(bytes);
            }

            var chars = new char[length];
            for (int i = 0; i < length; i++)
            {
                chars[i] = alphabet[bytes[i] % alphabet.Length];
            }

            return new string(chars);
        }
    }
}
