using System;
using System.Collections.Generic;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    /// <summary>
    /// 按符号流切断符写字：\f → 分页符，^b → 下一页分节。脚注 ^f/^e 仍走 WordNoteMarkHelper。
    /// </summary>
    internal static class WordBreakWriteHelper
    {
        private const int WdPageBreak = 7;
        private const int WdSectionBreakNextPage = 2;

        public static bool ContainsLayoutBreaks(string text)
        {
            return !string.IsNullOrEmpty(text)
                && (text.IndexOf('\f') >= 0
                    || text.IndexOf("^b", StringComparison.Ordinal) >= 0);
        }

        public static Word.Range WriteTextWithBreaks(Word.Document doc, Word.Range targetRange, string text)
        {
            if (doc == null || targetRange == null)
            {
                return targetRange;
            }

            text = text ?? "";
            if (!ContainsLayoutBreaks(text))
            {
                return WordNoteMarkHelper.WriteTextWithNoteMarks(doc, targetRange, text);
            }

            List<string> parts;
            List<bool> isPage;
            SplitByLayoutBreaks(text, out parts, out isPage);

            int start = targetRange.Start;
            Word.Range written = WordNoteMarkHelper.WriteTextWithNoteMarks(doc, targetRange, parts[0]);
            int end = written != null ? written.End : targetRange.End;

            for (int i = 0; i < isPage.Count; i++)
            {
                Word.Range at = doc.Range(end, end);
                if (isPage[i])
                {
                    InsertPageBreak(at);
                }
                else
                {
                    InsertSectionBreak(at);
                }

                try
                {
                    at.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                    end = at.End;
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine(
                        $"[WordBreakWriteHelper] 切断后定位失败: {ex.Message}");
                }

                if (i + 1 < parts.Count && !string.IsNullOrEmpty(parts[i + 1]))
                {
                    Word.Range next = doc.Range(end, end);
                    written = WordNoteMarkHelper.WriteTextWithNoteMarks(doc, next, parts[i + 1]);
                    if (written != null)
                    {
                        end = written.End;
                    }
                }
            }

            return doc.Range(start, end);
        }

        public static void InsertPageBreak(Word.Range range)
        {
            InsertBreak(range, Word.WdBreakType.wdPageBreak, WdPageBreak);
        }

        public static void InsertSectionBreak(Word.Range range)
        {
            InsertBreak(range, Word.WdBreakType.wdSectionBreakNextPage, WdSectionBreakNextPage);
        }

        private static void InsertBreak(Word.Range range, Word.WdBreakType kind, int fallback)
        {
            try
            {
                object breakType = kind;
                range.InsertBreak(ref breakType);
                return;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine(
                    $"[WordBreakWriteHelper] InsertBreak({kind}) 失败: {ex.Message}");
            }

            dynamic dyn = range;
            dyn.InsertBreak(fallback);
        }

        private static void SplitByLayoutBreaks(string text, out List<string> parts, out List<bool> isPage)
        {
            parts = new List<string>();
            isPage = new List<bool>();
            int start = 0;
            int i = 0;
            while (i < text.Length)
            {
                if (text[i] == '\f')
                {
                    parts.Add(text.Substring(start, i - start));
                    isPage.Add(true);
                    i++;
                    start = i;
                    continue;
                }

                if (i + 1 < text.Length && text[i] == '^' && text[i + 1] == 'b')
                {
                    parts.Add(text.Substring(start, i - start));
                    isPage.Add(false);
                    i += 2;
                    start = i;
                    continue;
                }

                i++;
            }

            parts.Add(text.Substring(start));
        }
    }
}
