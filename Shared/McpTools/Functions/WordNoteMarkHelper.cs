using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    /// <summary>
    /// ReadText 将脚注/尾注引用抽成 ^f / ^e；写入文档时再还原为真正的 Note 引用。
    /// </summary>
    internal static class WordNoteMarkHelper
    {
        private static readonly Regex NoteMarkSplit = new Regex(@"(\^[fe])", RegexOptions.Compiled);

        public static bool ContainsNoteMarks(string text)
        {
            return !string.IsNullOrEmpty(text)
                && (text.IndexOf("^f", StringComparison.Ordinal) >= 0
                    || text.IndexOf("^e", StringComparison.Ordinal) >= 0);
        }

        /// <summary>
        /// 将 text 写入 targetRange；若含 ^f/^e，先保存原 Range 内脚注/尾注正文，再按标记重建引用。
        /// 返回写入后覆盖的 Range。
        /// </summary>
        public static Word.Range WriteTextWithNoteMarks(Word.Document doc, Word.Range targetRange, string text)
        {
            if (doc == null || targetRange == null)
            {
                return targetRange;
            }

            text = text ?? "";
            if (!ContainsNoteMarks(text))
            {
                targetRange.Text = text;
                return doc.Range(targetRange.Start, targetRange.End);
            }

            List<string> footnoteTexts = CaptureFootnoteTexts(targetRange);
            List<string> endnoteTexts = CaptureEndnoteTexts(targetRange);

            int start = targetRange.Start;
            targetRange.Text = "";
            int pos = start;
            int fi = 0;
            int ei = 0;

            foreach (string part in NoteMarkSplit.Split(text))
            {
                if (part == "^f")
                {
                    string note = fi < footnoteTexts.Count ? footnoteTexts[fi++] : "";
                    Word.Range at = doc.Range(pos, pos);
                    object missing = Type.Missing;
                    object noteObj = note ?? "";
                    Word.Footnote fn = doc.Footnotes.Add(at, ref missing, ref noteObj);
                    pos = fn.Reference.End;
                    System.Diagnostics.Debug.WriteLine(
                        $"[WordNoteMarkHelper] 写入脚注引用 pos={pos} reused={(fi <= footnoteTexts.Count)}");
                }
                else if (part == "^e")
                {
                    string note = ei < endnoteTexts.Count ? endnoteTexts[ei++] : "";
                    Word.Range at = doc.Range(pos, pos);
                    object missing = Type.Missing;
                    object noteObj = note ?? "";
                    Word.Endnote en = doc.Endnotes.Add(at, ref missing, ref noteObj);
                    pos = en.Reference.End;
                    System.Diagnostics.Debug.WriteLine(
                        $"[WordNoteMarkHelper] 写入尾注引用 pos={pos} reused={(ei <= endnoteTexts.Count)}");
                }
                else if (part.Length > 0)
                {
                    Word.Range at = doc.Range(pos, pos);
                    at.Text = part;
                    pos = at.End;
                }
            }

            return doc.Range(start, pos);
        }

        private static List<string> CaptureFootnoteTexts(Word.Range range)
        {
            var list = new List<string>();
            if (range == null)
            {
                return list;
            }

            try
            {
                foreach (Word.Footnote fn in range.Footnotes)
                {
                    list.Add(NormalizeNoteBody(fn.Range?.Text));
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[WordNoteMarkHelper] 采集脚注失败: {ex.Message}");
            }

            return list;
        }

        private static List<string> CaptureEndnoteTexts(Word.Range range)
        {
            var list = new List<string>();
            if (range == null)
            {
                return list;
            }

            try
            {
                foreach (Word.Endnote en in range.Endnotes)
                {
                    list.Add(NormalizeNoteBody(en.Range?.Text));
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[WordNoteMarkHelper] 采集尾注失败: {ex.Message}");
            }

            return list;
        }

        private static string NormalizeNoteBody(string text)
        {
            if (string.IsNullOrEmpty(text))
            {
                return "";
            }

            // 去掉 Word 注尾常见控制符 / 段标
            char last = text[text.Length - 1];
            if (last == '\u0007' || last == '\r' || last == '\n')
            {
                text = text.Substring(0, text.Length - 1);
            }

            return text;
        }
    }
}
