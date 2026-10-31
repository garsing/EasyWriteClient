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
            catch (Exception ex)
            {
                Dbg("TypesFromRange err=" + ex.Message);
            }

            Dbg("TypesFromRange " + AutonumConvert.DebugDescribeMarks(marks) + " " + RangeBrief(range));
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

                Dbg("TypesFromCodes " + code
                    + " inner=" + Snip(inner)
                    + " content=" + Snip(content)
                    + " pair=" + AutonumConvert.DebugDescribeMarks(
                        AutonumConvert.PrefixMarksFromPair(inner, content)));
            }

            Dbg("TypesFromCodes all " + AutonumConvert.DebugDescribeMarks(marks));
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
            Dbg("WriteReplace numbered=" + numbered
                + " marks=" + AutonumConvert.DebugDescribeMarks(originalMarks)
                + " modelNew=" + Snip(modelNew)
                + " story=" + Snip(story)
                + " " + RangeBrief(target));
            Word.Range written = null;
            if (numbered)
            {
                written = TryFindReplace(doc, target, story);
            }

            if (written == null)
            {
                Dbg("WriteReplace fallback WriteTextWithBreaks");
                written = WordBreakWriteHelper.WriteTextWithBreaks(doc, target, story);
            }

            Dbg("WriteReplace written " + RangeBrief(written));
            return written;
        }

        public static Word.Range WriteInsert(
            Word.Document doc,
            Word.Range insertAt,
            string modelNew,
            IList<string> originalMarks)
        {
            string story = AutonumConvert.StripNewByOriginalTypes(originalMarks, modelNew);
            Dbg("WriteInsert marks=" + AutonumConvert.DebugDescribeMarks(originalMarks)
                + " modelNew=" + Snip(modelNew)
                + " story=" + Snip(story)
                + " " + RangeBrief(insertAt));
            Word.Range written = WordBreakWriteHelper.WriteTextWithBreaks(doc, insertAt, story);
            Dbg("WriteInsert written " + RangeBrief(written));
            return written;
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

            Dbg("ClearExtra " + RangeBrief(written)
                + " modelNew=" + Snip(modelNew)
                + " marks=" + AutonumConvert.DebugDescribeMarks(originalMarks));
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
                    Dbg("ClearExtra i=" + i
                        + " ls=\"" + (ls ?? "") + "\""
                        + " seg=" + Snip(segs[si])
                        + " wantPrefix=" + want
                        + " remove=" + (!string.IsNullOrEmpty(ls) && !want)
                        + " story=" + Snip(SafeText(pr)));
                    if (!string.IsNullOrEmpty(ls) && !want)
                    {
                        try
                        {
                            p.Range.ListFormat.RemoveNumbers();
                        }
                        catch (Exception ex)
                        {
                            Dbg("ClearExtra RemoveNumbers err=" + ex.Message);
                        }
                    }

                    si++;
                }
            }
            catch
            {
            }
        }

        private static Word.Range TryFindReplace(Word.Document doc, Word.Range target, string story)
        {
            if (doc == null || target == null)
            {
                return null;
            }

            try
            {
                string before = SafeText(target);
                int start = target.Start;
                int origEnd = target.End;
                string findText = (before ?? "").Replace("\r", "^p");
                string repl = (story ?? "").Replace("\r", "^p");
                Dbg("TryFindReplace before=" + Snip(before)
                    + " start=" + start + " end=" + origEnd
                    + " find=" + Snip(findText)
                    + " repl=" + Snip(repl)
                    + " storyLen=" + (story ?? "").Length);
                if (string.IsNullOrEmpty(findText))
                {
                    Dbg("TryFindReplace skip empty find");
                    return null;
                }

                Word.Range search = doc.Range(start, doc.Content.End);
                Word.Find find = search.Find;
                find.ClearFormatting();
                find.Replacement.ClearFormatting();
                object findArg = findText;
                object falseArg = false;
                object trueArg = true;
                object wrapArg = Word.WdFindWrap.wdFindStop;
                object replaceWithArg = repl;
                object replaceArg = Word.WdReplace.wdReplaceOne;
                bool ok = find.Execute(
                    findArg,
                    falseArg,
                    falseArg,
                    falseArg,
                    falseArg,
                    falseArg,
                    trueArg,
                    wrapArg,
                    falseArg,
                    replaceWithArg,
                    replaceArg);
                string afterOrig = SafeSlice(doc, start, origEnd);
                bool mutated = !string.Equals(before, afterOrig, StringComparison.Ordinal);
                Dbg("TryFindReplace execute=" + ok
                    + " mutated=" + mutated
                    + " afterOrigSpan=" + Snip(afterOrig)
                    + " emptySuccess=" + (ok && !mutated));
                if (!ok || !mutated)
                {
                    return null;
                }

                int end = start + (story ?? "").Length;
                if (end < start)
                {
                    end = start;
                }

                Word.Range synth = doc.Range(start, end);
                Dbg("TryFindReplace synth start=" + start + " end=" + end
                    + " synthText=" + Snip(SafeText(synth)));
                return synth;
            }
            catch (Exception ex)
            {
                Dbg("TryFindReplace err=" + ex.Message);
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

        private static void Dbg(string message)
        {
            EasyWriteDiagnostics.Log(DebugCategory.Autonum, "[autonum] " + (message ?? ""));
        }

        private static string Snip(string text)
        {
            string s = TextUtils.EscapeNewlinesForDb(text ?? "");
            if (s.Length <= 120)
            {
                return s;
            }

            return s.Substring(0, 120) + "…";
        }

        private static string RangeBrief(Word.Range range)
        {
            if (range == null)
            {
                return "range=null";
            }

            try
            {
                string t = range.Text ?? "";
                bool cr = t.Length > 0 && t[t.Length - 1] == '\r';
                return "start=" + range.Start + " end=" + range.End + " cr=" + cr + " text=" + Snip(t);
            }
            catch (Exception ex)
            {
                return "range-err=" + ex.Message;
            }
        }

        private static string SafeText(Word.Range range)
        {
            if (range == null)
            {
                return "";
            }

            try
            {
                return range.Text ?? "";
            }
            catch
            {
                return "";
            }
        }

        private static string SafeSlice(Word.Document doc, int start, int end)
        {
            if (doc == null)
            {
                return "";
            }

            try
            {
                int cap = doc.Content.End;
                if (start < 0)
                {
                    start = 0;
                }

                if (end > cap)
                {
                    end = cap;
                }

                if (end < start)
                {
                    end = start;
                }

                return doc.Range(start, end).Text ?? "";
            }
            catch
            {
                return "";
            }
        }

        private static int SafeStart(Word.Range range)
        {
            try
            {
                return range == null ? -1 : range.Start;
            }
            catch
            {
                return -1;
            }
        }

        private static int SafeEnd(Word.Range range)
        {
            try
            {
                return range == null ? -1 : range.End;
            }
            catch
            {
                return -1;
            }
        }
    }
}
