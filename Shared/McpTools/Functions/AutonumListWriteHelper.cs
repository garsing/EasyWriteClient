using System;
using System.Collections.Generic;
using WordAddIn1.DocumentMapping;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    internal static class AutonumListWriteHelper
    {
        public static List<string> TypesFromRange(Word.Range range)
        {
            var marks = new List<string>();
            if (range == null)
            {
                return marks;
            }

            try
            {
                foreach (Word.Paragraph p in range.Paragraphs)
                {
                    string ls = ListStringOf(p);
                    if (!string.IsNullOrEmpty(ls))
                    {
                        marks.Add(ls.Trim());
                    }
                }
            }
            catch
            {
            }

            return marks;
        }

        public static List<string> TypesFromCodes(string codes)
        {
            var marks = new List<string>();
            if (string.IsNullOrEmpty(codes))
            {
                return marks;
            }

            foreach (string raw in codes.Split(new[] { ',', ';', ' ' }, StringSplitOptions.RemoveEmptyEntries))
            {
                string code = raw.Trim();
                string inner = DocumentState.GetDisplayInner(code);
                string content = code.StartsWith("P_", StringComparison.Ordinal)
                    ? DocumentState.GetParagraphContent(code)
                    : DocumentState.GetSentenceContent(code);
                foreach (string m in AutonumConvert.PrefixMarksFromPair(inner, content))
                {
                    if (!string.IsNullOrEmpty(m))
                    {
                        marks.Add(m);
                    }
                }
            }

            return marks;
        }

        public static bool RangeHasAutonum(Word.Range range)
        {
            return TypesFromRange(range).Count > 0;
        }

        public static Word.Range WriteReplace(
            Word.Document doc,
            Word.Range target,
            string modelNew,
            IList<string> originalMarks,
            bool numbered)
        {
            string story = AutonumConvert.StripNewByOriginalTypes(originalMarks, modelNew);
            Word.Range written = null;
            if (numbered)
            {
                written = TryFindReplace(target, story);
            }

            if (written == null)
            {
                written = WordBreakWriteHelper.WriteTextWithBreaks(doc, target, story);
            }

            return written;
        }

        public static Word.Range WriteInsert(
            Word.Document doc,
            Word.Range insertAt,
            string modelNew,
            IList<string> originalMarks)
        {
            string story = AutonumConvert.StripNewByOriginalTypes(originalMarks, modelNew);
            return WordBreakWriteHelper.WriteTextWithBreaks(doc, insertAt, story);
        }

        public static void ClearExtraListOnWritten(
            Word.Range written,
            string modelNew,
            IList<string> originalMarks)
        {
            if (written == null)
            {
                return;
            }

            string[] segs = (modelNew ?? "").Split('\r');
            int si = 0;
            try
            {
                Word.Document doc = written.Document;
                int start = written.Start;
                int end = written.End;
                for (int i = 1; i <= doc.Paragraphs.Count; i++)
                {
                    Word.Paragraph p = doc.Paragraphs[i];
                    Word.Range pr = p.Range;
                    if (pr.End <= start || pr.Start >= end)
                    {
                        continue;
                    }

                    if (pr.Start < start)
                    {
                        continue;
                    }

                    if (si >= segs.Length)
                    {
                        break;
                    }

                    string ls = ListStringOf(p);
                    bool want = AutonumConvert.SegmentHasOriginalTypePrefix(segs[si], originalMarks);
                    if (!string.IsNullOrEmpty(ls) && !want)
                    {
                        try
                        {
                            p.Range.ListFormat.RemoveNumbers();
                        }
                        catch
                        {
                        }
                    }

                    si++;
                }
            }
            catch
            {
            }
        }

        private static Word.Range TryFindReplace(Word.Range target, string story)
        {
            if (target == null)
            {
                return null;
            }

            try
            {
                string findText = (target.Text ?? "").Replace("\r", "^p");
                string repl = (story ?? "").Replace("\r", "^p");
                if (string.IsNullOrEmpty(findText))
                {
                    return null;
                }

                int start = target.Start;
                Word.Find find = target.Find;
                find.ClearFormatting();
                find.Replacement.ClearFormatting();
                find.Text = findText;
                find.Replacement.Text = repl;
                find.Forward = true;
                find.Wrap = Word.WdFindWrap.wdFindStop;
                find.Format = false;
                find.MatchCase = true;
                object replace = Word.WdReplace.wdReplaceOne;
                bool ok = find.Execute(
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    replace);
                if (!ok)
                {
                    return null;
                }

                Word.Document doc = target.Document;
                int end = start + (story ?? "").Length;
                if (end < start)
                {
                    end = start;
                }

                return doc.Range(start, end);
            }
            catch
            {
                return null;
            }
        }

        private static string ListStringOf(Word.Paragraph p)
        {
            try
            {
                return (p.Range.ListFormat.ListString ?? "").Trim();
            }
            catch
            {
                return "";
            }
        }
    }
}
