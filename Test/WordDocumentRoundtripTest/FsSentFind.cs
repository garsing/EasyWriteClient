using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using WordAddIn1;
using WordAddIn1.DocumentMapping.CodeResolve;
using Word = Microsoft.Office.Interop.Word;

namespace WordDocumentRoundtripTest
{
    /// <summary>
    /// 真稿：每个 S_ 或 P_ 落到自己的 Range 后，在该 Range 内 Find。
    /// 命中不得落在 Range 起点之前或终点之后。直连 COM，不走 Desktop。
    /// </summary>
    internal static class FsSentFind
    {
        public static int Run(string host, string docPath)
        {
            return Run(host, docPath, false);
        }

        public static int Run(string host, string docPath, bool paragraph)
        {
            string caseId = paragraph ? "fs-para-find" : "fs-sent-find";
            Console.WriteLine(paragraph
                ? "真稿段落范围内 Find（不得落到 Range 前）"
                : "真稿句子范围内 Find（不得落到 Range 前）");
            Console.WriteLine(string.Equals(host, "wps", StringComparison.OrdinalIgnoreCase)
                ? "宿主 WPS 文字"
                : "宿主 Word");
            if (string.IsNullOrWhiteSpace(docPath) || !System.IO.File.Exists(docPath))
            {
                Console.WriteLine("稿不存在: " + docPath);
                return 1;
            }

            Word.Application app = AttachApp(host, out bool created);
            if (app == null)
            {
                Console.WriteLine(string.Equals(host, "wps", StringComparison.OrdinalIgnoreCase) ? "无 WPS" : "无 Word");
                return 1;
            }

            Word.WdAlertLevel oldAlerts = app.DisplayAlerts;
            bool oldUpdating = app.ScreenUpdating;
            app.DisplayAlerts = Word.WdAlertLevel.wdAlertsNone;
            app.ScreenUpdating = false;
            Word.Document doc = null;
            bool weOpened = false;
            try
            {
                doc = OpenReadOnly(app, docPath, out weOpened);
                Console.WriteLine("稿=" + docPath);
                var sw = System.Diagnostics.Stopwatch.StartNew();
                WordDocumentExtractor.ProcessDocument(
                    doc,
                    ProcessDocumentOptions.ForProcessActions(caseId));
                sw.Stop();
                Console.WriteLine("读文 " + sw.ElapsedMilliseconds + " ms");

                List<string> codes = paragraph
                    ? ParagraphCodes()
                    : SentenceCodes(DocumentState.SegSnapshot);
                Console.WriteLine((paragraph ? "段落码 " : "句子码 ") + codes.Count);
                TableScopeIndex tableScope = paragraph
                    ? null
                    : TableSentenceScopeHelper.BuildIndexFromProcessCache(DocumentState.LastProcessCache);

                int inside = 0;
                int miss = 0;
                int before = 0;
                int after = 0;
                int ambiguous = 0;
                int noRange = 0;
                int noNeedle = 0;
                int findError = 0;
                var problems = new List<string>();

                for (int i = 0; i < codes.Count; i++)
                {
                    string code = codes[i];
                    Word.Range range;
                    if (paragraph)
                    {
                        int times = ParagraphCodeAmbiguityHelper.CountParagraphCodeInDisplay(code);
                        if (times != 1)
                        {
                            ambiguous++;
                            Note(problems, "歧义 " + code + " matches=" + times);
                            continue;
                        }

                        range = ParagraphRangeLocator.TryResolve(doc, code, caseId);
                    }
                    else
                    {
                        DisplaySequenceLocateResult locate = DisplaySequenceLocator.TryResolve(
                            new[] { code },
                            new DisplaySequenceLocateOptions
                            {
                                Snapshot = DocumentState.Snapshot,
                                SegSnapshot = DocumentState.SegSnapshot,
                                TableScope = tableScope
                            });
                        if (locate.MatchCount != 1 || !locate.DisplayStartPosition.HasValue)
                        {
                            ambiguous++;
                            Note(problems, "歧义 " + code + " matches=" + locate.MatchCount);
                            continue;
                        }

                        range = DisplayPositionRangeResolver.ResolveSpan(
                            doc,
                            new[] { code },
                            locate.DisplayStartPosition.Value,
                            null,
                            tableScope,
                            locate.OrderedDomainCodes,
                            caseId);
                    }
                    if (range == null)
                    {
                        noRange++;
                        Note(problems, "无Range " + code);
                        continue;
                    }

                    int start = range.Start;
                    int end = range.End;
                    if (paragraph)
                    {
                        end = ShrinkOffCellEnd(doc, start, end);
                    }

                    string needle = Needle(paragraph
                        ? DocumentState.GetParagraphContent(code)
                        : DocumentState.GetSentenceContent(code));
                    if (needle.Length == 0)
                    {
                        noNeedle++;
                        Note(problems, "无字 " + code + " window=" + start + "-" + end);
                        continue;
                    }

                    Word.Range window = doc.Range(start, end);
                    string windowText = window.Text ?? "";
                    string preview = windowText;
                    if (preview.Length > 24)
                    {
                        preview = preview.Substring(0, 24);
                    }

                    if (string.Equals(windowText, needle, StringComparison.Ordinal))
                    {
                        inside++;
                        if ((i + 1) % 50 == 0)
                        {
                            Console.WriteLine("已查 " + (i + 1) + "/" + codes.Count + " 窗前=" + before);
                        }

                        continue;
                    }

                    Word.Find find = window.Find;
                    find.ClearFormatting();
                    find.Text = needle;
                    find.MatchCase = true;
                    find.MatchWholeWord = false;
                    find.MatchWildcards = false;
                    find.Wrap = Word.WdFindWrap.wdFindStop;
                    bool hit;
                    try
                    {
                        hit = find.Execute();
                    }
                    catch (Exception ex)
                    {
                        findError++;
                        Note(problems, "Find异常 " + code + " " + ex.Message);
                        continue;
                    }

                    if (!hit)
                    {
                        miss++;
                        Note(problems, "未命中 " + code + " window=" + start + "-" + end + " needle=" + needle);
                        continue;
                    }

                    if (window.Start < start)
                    {
                        before++;
                        string line = "BEFORE " + code
                            + " window=" + start + "-" + end
                            + " found=" + window.Start + "-" + window.End
                            + " needle=U+" + ((int)needle[0]).ToString("X4")
                            + " preview=" + Quote(preview)
                            + " " + DescribeJump(doc, start, end, window.Start);
                        Console.WriteLine(line);
                        Note(problems, line);
                    }
                    else if (window.Start >= end)
                    {
                        after++;
                        string line = "AFTER " + code
                            + " window=" + start + "-" + end
                            + " found=" + window.Start + "-" + window.End
                            + " needle=U+" + ((int)needle[0]).ToString("X4")
                            + " preview=" + Quote(preview);
                        Console.WriteLine(line);
                        Note(problems, line);
                    }
                    else
                    {
                        inside++;
                    }

                    if ((i + 1) % 50 == 0)
                    {
                        Console.WriteLine("已查 " + (i + 1) + "/" + codes.Count + " 窗前=" + before);
                    }
                }

                bool pass = before == 0 && after == 0 && miss == 0 && ambiguous == 0 && noRange == 0 && findError == 0;
                Console.WriteLine(
                    caseId + " " + (pass ? "过" : "败")
                    + (paragraph ? " 段内=" : " 句内=") + inside
                    + " 窗前=" + before
                    + " 窗后=" + after
                    + " 未命中=" + miss
                    + " 歧义=" + ambiguous
                    + " 无Range=" + noRange
                    + " 无字=" + noNeedle
                    + " Find异常=" + findError);
                foreach (string line in problems)
                {
                    Console.WriteLine(line);
                }

                return pass ? 0 : 1;
            }
            catch (Exception ex)
            {
                Console.WriteLine("失败: " + ex.GetType().Name + " " + ex.Message);
                return 1;
            }
            finally
            {
                if (weOpened)
                {
                    CloseDoc(doc);
                }

                try
                {
                    app.ScreenUpdating = oldUpdating;
                    app.DisplayAlerts = oldAlerts;
                }
                catch
                {
                }

                if (created)
                {
                    try
                    {
                        app.Quit(Word.WdSaveOptions.wdDoNotSaveChanges);
                    }
                    catch
                    {
                    }
                }
            }
        }

        private static void Note(List<string> problems, string line)
        {
            if (problems.Count < 30)
            {
                problems.Add(line);
            }
        }

        private static Word.Document OpenReadOnly(Word.Application app, string path, out bool weOpened)
        {
            weOpened = false;
            string norm = System.IO.Path.GetFullPath(path);
            foreach (Word.Document open in app.Documents)
            {
                try
                {
                    if (string.Equals(
                        System.IO.Path.GetFullPath(open.FullName),
                        norm,
                        StringComparison.OrdinalIgnoreCase))
                    {
                        return open;
                    }
                }
                catch
                {
                }
            }

            object fileName = path;
            object readOnly = true;
            object missing = System.Reflection.Missing.Value;
            Word.Document doc = app.Documents.Open(ref fileName, ref missing, ref readOnly);
            weOpened = true;
            return doc;
        }

        private static Word.Application AttachApp(string host, out bool created)
        {
            created = false;
            if (string.Equals(host, "wps", StringComparison.OrdinalIgnoreCase))
            {
                string[] ids = { "Kwps.Application", "kwps.Application", "wps.Application" };
                foreach (string id in ids)
                {
                    try
                    {
                        object obj = Marshal.GetActiveObject(id);
                        if (obj != null)
                        {
                            return (Word.Application)obj;
                        }
                    }
                    catch
                    {
                    }
                }

                foreach (string id in ids)
                {
                    try
                    {
                        Type t = Type.GetTypeFromProgID(id);
                        if (t == null)
                        {
                            continue;
                        }

                        Word.Application createdApp = (Word.Application)Activator.CreateInstance(t);
                        createdApp.Visible = true;
                        created = true;
                        return createdApp;
                    }
                    catch
                    {
                    }
                }

                return null;
            }

            try
            {
                return (Word.Application)Marshal.GetActiveObject("Word.Application");
            }
            catch
            {
                Word.Application createdApp = new Word.Application();
                createdApp.Visible = true;
                created = true;
                return createdApp;
            }
        }

        private static void CloseDoc(Word.Document doc)
        {
            if (doc == null)
            {
                return;
            }

            try
            {
                object save = false;
                doc.Close(ref save);
            }
            catch
            {
            }
        }

        private static int ShrinkOffCellEnd(Word.Document doc, int start, int end)
        {
            if (end <= start + 1)
            {
                return end;
            }

            try
            {
                Word.Range probe = doc.Range(start, end);
                bool inTable = (bool)probe.Information[Word.WdInformation.wdWithInTable];
                if (!inTable)
                {
                    return end;
                }

                Word.Range cell = probe.Cells[1].Range;
                if (cell.End != end)
                {
                    return end;
                }

                return end - 1;
            }
            catch
            {
                return end;
            }
        }

        private static string DescribeJump(Word.Document doc, int start, int end, int foundStart)
        {
            try
            {
                Word.Range para = doc.Range(start, end);
                bool inTable = (bool)para.Information[Word.WdInformation.wdWithInTable];
                int cellStart = -1;
                int cellEnd = -1;
                if (inTable)
                {
                    Word.Range cell = para.Cells[1].Range;
                    cellStart = cell.Start;
                    cellEnd = cell.End;
                }

                int lead = Math.Max(0, start - 20);
                string beforeText = Quote(doc.Range(lead, start).Text ?? "");
                string foundText = Quote(doc.Range(foundStart, foundStart + 1).Text ?? "");
                return "inTable=" + inTable
                    + " cell=" + cellStart + "-" + cellEnd
                    + " before20=" + beforeText
                    + " foundChar=" + foundText;
            }
            catch (Exception ex)
            {
                return "ctx=" + ex.GetType().Name;
            }
        }

        private static string Quote(string text)
        {
            if (string.IsNullOrEmpty(text))
            {
                return "";
            }

            var sb = new System.Text.StringBuilder();
            foreach (char c in text)
            {
                if (c == '\r')
                {
                    sb.Append("\\r");
                }
                else if (c == '\n')
                {
                    sb.Append("\\n");
                }
                else if (c == '\a')
                {
                    sb.Append("\\a");
                }
                else if (c < ' ')
                {
                    sb.Append("\\x").Append(((int)c).ToString("X2"));
                }
                else
                {
                    sb.Append(c);
                }
            }

            return sb.ToString();
        }

        private static List<string> ParagraphCodes()
        {
            var codes = new List<string>();
            var seen = new HashSet<string>(StringComparer.Ordinal);
            foreach (string code in ParagraphCodeAmbiguityHelper.GetOrderedParagraphCodesFromDisplay())
            {
                if (seen.Add(code))
                {
                    codes.Add(code);
                }
            }

            return codes;
        }

        private static List<string> SentenceCodes(string segSnapshot)
        {
            var codes = new List<string>();
            if (string.IsNullOrEmpty(segSnapshot))
            {
                return codes;
            }

            foreach (string seg in segSnapshot.Split(';'))
            {
                foreach (string raw in seg.Split(','))
                {
                    string code = raw.Trim();
                    if (code.StartsWith("S_", StringComparison.Ordinal))
                    {
                        codes.Add(code);
                    }
                }
            }

            return codes;
        }

        private static string Needle(string text)
        {
            if (string.IsNullOrEmpty(text))
            {
                return "";
            }

            foreach (char c in text)
            {
                if (c == '\r' || c == '\n' || c == '\a' || c == '\v' || c == '\f' || c == '^' || char.IsWhiteSpace(c))
                {
                    continue;
                }

                return c.ToString();
            }

            return "";
        }
    }
}
