using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using Word = Microsoft.Office.Interop.Word;
using WordAddIn1;
using WordAddIn1.DocumentMapping;

namespace WordDocumentRoundtripTest
{
    /// <summary>
    /// 探针：游标逐位 Find，看 hit.Start &lt; cursor（stale）是不是只在进表后出现。
    /// 直连本机 Word COM，不走 Desktop 管道。
    /// </summary>
    internal static class FindStaleCases
    {
        public const int Target = 20;

        private const string Needle = "歧义甲。";

        /// <summary>fs-dalign 默认稿；可用 --doc= 换文件。</summary>
        internal static string AlignDocPath = "";

        public static void Run(TestRun run, string host, string caseFilter = "", string docPath = "")
        {
            AlignDocPath = docPath ?? "";
            Word.Application app = null;
            bool createdApp = false;
            try
            {
                app = AttachApp(host, out createdApp);
                if (app == null)
                {
                    throw new InvalidOperationException(
                        string.Equals(host, "wps", StringComparison.OrdinalIgnoreCase)
                            ? "无 WPS"
                            : "无 Word");
                }

                bool onlyReal = !string.IsNullOrEmpty(caseFilter)
                    && caseFilter.IndexOf("real", StringComparison.OrdinalIgnoreCase) >= 0;
                bool onlyPage = !string.IsNullOrEmpty(caseFilter)
                    && caseFilter.IndexOf("page", StringComparison.OrdinalIgnoreCase) >= 0;
                bool onlyPlus = !string.IsNullOrEmpty(caseFilter)
                    && caseFilter.IndexOf("plus", StringComparison.OrdinalIgnoreCase) >= 0;
                bool onlyParaText = !string.IsNullOrEmpty(caseFilter)
                    && caseFilter.IndexOf("pratext", StringComparison.OrdinalIgnoreCase) >= 0;
                bool onlyDalign = !string.IsNullOrEmpty(caseFilter)
                    && caseFilter.IndexOf("dalign", StringComparison.OrdinalIgnoreCase) >= 0;
                bool onlyPwalk = !string.IsNullOrEmpty(caseFilter)
                    && caseFilter.IndexOf("pwalk", StringComparison.OrdinalIgnoreCase) >= 0;
                bool onlyPsplit = !string.IsNullOrEmpty(caseFilter)
                    && caseFilter.IndexOf("psplit", StringComparison.OrdinalIgnoreCase) >= 0;
                bool onlyPnth = !string.IsNullOrEmpty(caseFilter)
                    && caseFilter.IndexOf("pnth", StringComparison.OrdinalIgnoreCase) >= 0;
                bool onlyPchars = !string.IsNullOrEmpty(caseFilter)
                    && caseFilter.IndexOf("pchars", StringComparison.OrdinalIgnoreCase) >= 0;
                if (onlyPnth)
                {
                    RunPublicCodeNth(run);
                }
                else if (onlyPsplit)
                {
                    RunSplitConvertedVsParas(run, app);
                }
                else if (onlyPwalk)
                {
                    RunPublicCodeNth(run);
                    RunConvertParaWalk(run, app);
                }
                else if (onlyPchars)
                {
                    RunRealDocParaChars(run, app);
                }
                else if (onlyDalign)
                {
                    RunRealDocDisplayAlign(run, app);
                }
                else if (onlyParaText)
                {
                    RunRealDocParaTexts(run, app);
                }
                else if (onlyPlus)
                {
                    RunCellPlusCr(run, app);
                }
                else if (onlyPage)
                {
                    RunPageBreakChar(run, app);
                }
                else if (onlyReal)
                {
                    RunRealDocCharsAbut(run, app);
                }
                else
                {
                    RunPlain(run, app);
                    RunTable(run, app);
                    RunNestedCellText(run, app);
                    RunCellPreAmbigFind(run, app);
                    RunCellPlainAmbigFind(run, app);
                    RunCellPlainWalk(run, app);
                    RunCellCharacters(run, app);
                    RunDocCharacters(run, app);
                    RunCharsAbut(run, app);
                    RunPageBreakChar(run, app);
                }
            }
            catch (Exception ex)
            {
                run.CaseFail("fs-plain-write", ex.GetType().Name + ": " + ex.Message);
                run.CaseFail("fs-plain-zero-stale", "无 Word");
                run.CaseFail("fs-table-write", "无 Word");
                run.CaseFail("fs-table-next-count", "无 Word");
                run.CaseFail("fs-stale-only-in-table", "无 Word");
                run.CaseFail("fs-cell21-scope", "无 Word");
                run.CaseFail("fs-table-20-37", "无 Word");
                run.CaseFail("fs-para-2nd-find", "无 Word");
                run.CaseFail("fs-nested-cell-text", "无 Word");
                run.CaseFail("fs-cell-pre-first", "无 Word");
                run.CaseFail("fs-cell-pre-second", "无 Word");
                run.CaseFail("fs-cell-plain-first", "无 Word");
                run.CaseFail("fs-cell-plain-second", "无 Word");
                run.CaseFail("fs-cell-plain-walk", "无 Word");
                run.CaseFail("fs-cell-chars", "无 Word");
                run.CaseFail("fs-doc-chars", "无 Word");
                run.CaseFail("fs-chars-abut", "无 Word");
                run.CaseFail("fs-real-chars-abut", "无 Word");
                run.CaseFail("fs-text-align", "无 Word");
                run.CaseFail("fs-page-char", "无 Word");
            }
            finally
            {
                if (createdApp)
                {
                    // 自己拉起的 Word 只关稿，不 Quit。
                }
            }
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

        private static void RunPlain(TestRun run, Word.Application app)
        {
            Word.Document doc = null;
            try
            {
                doc = app.Documents.Add();
                doc.Content.Text = Needle + "\r间隔乙。\r" + Needle + "\r" + Needle + "\r" + Needle;
                run.CaseOk("fs-plain-write");

                StaleWalk walk = Walk(doc, Needle);
                Console.WriteLine(walk.FormatLog("无表对照"));
                Ok(run, "fs-plain-zero-stale", walk.StaleTotal == 0,
                    "无表也 stale " + walk.StaleTotal
                    + " 次（before=" + walk.StaleBefore
                    + " in=" + walk.StaleIn
                    + " after=" + walk.StaleAfter + "）");
            }
            catch (Exception ex)
            {
                run.CaseFail("fs-plain-write", ex.Message);
                run.CaseFail("fs-plain-zero-stale", "写/探针失败");
            }
            finally
            {
                CloseDoc(doc);
            }
        }

        private static void RunTable(TestRun run, Word.Application app)
        {
            Word.Document doc = null;
            try
            {
                doc = app.Documents.Add();
                if (!PlantTableFixture(doc))
                {
                    run.CaseFail("fs-table-write", "表没建出来");
                    run.CaseFail("fs-table-next-count", "无稿");
                    run.CaseFail("fs-stale-only-in-table", "无稿");
                    run.CaseFail("fs-cell21-scope", "无稿");
                    run.CaseFail("fs-table-20-37", "无稿");
                    run.CaseFail("fs-para-2nd-find", "无稿");
                    return;
                }

                run.CaseOk("fs-table-write");
                StaleWalk walk = Walk(doc, Needle);
                Console.WriteLine(walk.FormatLog("表前同文 + 表内同文"));
                Ok(run, "fs-table-next-count", walk.NextHits >= 4,
                    "Find Next 命中 " + walk.NextHits + "（期望 ≥4：表前2 + 表内2）");
                bool onlyInTable = walk.StaleBefore == 0 && walk.StaleAfter == 0;
                Ok(run, "fs-stale-only-in-table", onlyInTable,
                    "stale 不在「仅进表」：before=" + walk.StaleBefore
                    + " in=" + walk.StaleIn
                    + " after=" + walk.StaleAfter
                    + " total=" + walk.StaleTotal
                    + (walk.StaleTotal == 0 ? "（本夹具一次 stale 都没有）" : ""));
                ProbeSecondCellScope(run, doc);
                ProbeTableTailScope(run, doc);
                ProbeParagraphList(run, doc);
            }
            catch (Exception ex)
            {
                run.CaseFail("fs-table-write", ex.Message);
                run.CaseFail("fs-table-next-count", "写/探针失败");
                run.CaseFail("fs-stale-only-in-table", "写/探针失败");
                run.CaseFail("fs-cell21-scope", "写/探针失败");
                run.CaseFail("fs-table-20-37", "写/探针失败");
                run.CaseFail("fs-para-2nd-find", "写/探针失败");
            }
            finally
            {
                CloseDoc(doc);
            }
        }

        private static void RunNestedCellText(TestRun run, Word.Application app)
        {
            Word.Document doc = null;
            try
            {
                doc = app.Documents.Add();
                doc.Content.Text = "表前。\r";
                Word.Table outer = doc.Tables.Add(doc.Paragraphs[1].Range, 1, 2);
                outer.Cell(1, 1).Range.Text = "外格甲。";
                Word.Cell host = outer.Cell(1, 2);
                host.Range.Text = "外前甲。";
                Word.Range nestAt = host.Range;
                nestAt.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                nestAt.MoveEnd(Word.WdUnits.wdCharacter, -1);
                nestAt.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                Word.Table inner = doc.Tables.Add(nestAt, 1, 2);
                inner.Cell(1, 1).Range.Text = "内格乙。";
                inner.Cell(1, 2).Range.Text = "内格丙。";

                Word.Range outerCell = host.Range;
                Console.WriteLine("  --- 外层 [1,2] 内套表 ---");
                Console.WriteLine("    doc.Tables=" + doc.Tables.Count
                    + "  外表 " + outer.Range.Start + "-" + outer.Range.End
                    + "  Cell.Tables=" + host.Tables.Count
                    + "  Range.Tables=" + outerCell.Tables.Count);
                int ti = 0;
                foreach (Word.Table t in doc.Tables)
                {
                    ti++;
                    Console.WriteLine("    doc.Tables[" + ti + "] " + t.Range.Start + "-" + t.Range.End);
                }

                Console.WriteLine("    外格[1,2] " + outerCell.Start + "-" + outerCell.End);
                Console.WriteLine("    外格.Text=" + QuoteSpecial(outerCell.Text));
                Console.WriteLine("    外格[1,1].Text=" + QuoteSpecial(outer.Cell(1, 1).Range.Text));

                int p = 0;
                foreach (Word.Paragraph para in outerCell.Paragraphs)
                {
                    p++;
                    Console.WriteLine("    外格段P" + p + " " + para.Range.Start + "-" + para.Range.End
                        + "  " + QuoteSpecial(para.Range.Text));
                }

                if (host.Tables.Count > 0)
                {
                    Word.Table nest = host.Tables[1];
                    Console.WriteLine("    Cell.Tables[1] " + nest.Range.Start + "-" + nest.Range.End);
                    Console.WriteLine("    内[1,1] " + nest.Cell(1, 1).Range.Start + "-"
                        + nest.Cell(1, 1).Range.End + "  " + QuoteSpecial(nest.Cell(1, 1).Range.Text));
                    Console.WriteLine("    内[1,2] " + nest.Cell(1, 2).Range.Start + "-"
                        + nest.Cell(1, 2).Range.End + "  " + QuoteSpecial(nest.Cell(1, 2).Range.Text));
                }

                string cellText = outerCell.Text ?? "";
                bool hasBoth = cellText.IndexOf("外前甲", StringComparison.Ordinal) >= 0
                    && cellText.IndexOf("内格乙", StringComparison.Ordinal) >= 0;
                Ok(run, "fs-nested-cell-text", hasBoth, "外格 Text 未见「外前+内表」: " + QuoteSpecial(cellText));
            }
            catch (Exception ex)
            {
                run.CaseFail("fs-nested-cell-text", ex.Message);
            }
            finally
            {
                CloseDoc(doc);
            }
        }

        private static void RunCellPreAmbigFind(TestRun run, Word.Application app)
        {
            Word.Document doc = null;
            try
            {
                doc = app.Documents.Add();
                Word.Table outer = doc.Tables.Add(doc.Range(0, 0), 1, 1);
                Word.Cell host = outer.Cell(1, 1);
                host.Range.Text = Needle + "间隔乙。" + Needle;
                Word.Range nestAt = host.Range;
                nestAt.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                nestAt.MoveEnd(Word.WdUnits.wdCharacter, -1);
                nestAt.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                Word.Table inner = doc.Tables.Add(nestAt, 1, 2);
                inner.Cell(1, 1).Range.Text = Needle;
                inner.Cell(1, 2).Range.Text = "内格丙。";

                Word.Range cell = host.Range;
                Console.WriteLine("  --- 格内表前两处同文 + 内表 ---");
                Console.WriteLine("    cell=" + cell.Start + "-" + cell.End
                    + "  Text=" + QuoteSpecial(cell.Text)
                    + "  Cell.Tables=" + host.Tables.Count);

                bool firstOk = TryFind(doc, cell.Start, cell.End, Needle, out int h1s, out int h1e);
                bool firstIn = firstOk && h1s >= cell.Start && h1e <= cell.End && h1s >= cell.Start;
                Console.WriteLine("    #1 整格搜 Execute=" + firstOk
                    + " hit=" + h1s + "-" + h1e
                    + (firstOk && h1s < cell.Start ? " 弹回格外" : (firstIn ? " 落在格内" : "")));
                Ok(run, "fs-cell-pre-first", firstOk && firstIn && h1s >= cell.Start,
                    firstOk
                        ? ("hit=" + h1s + "-" + h1e + " cell=" + cell.Start + "-" + cell.End)
                        : "整格搜第一处失败");

                if (!firstOk)
                {
                    run.CaseFail("fs-cell-pre-second", "没有第一处，无法推进");
                    return;
                }

                bool secondOk = TryFind(doc, h1e, cell.End, Needle, out int h2s, out int h2e);
                bool bounced = secondOk && h2s < h1e;
                bool inside = secondOk && h2s >= h1e && h2e <= cell.End;
                Console.WriteLine("    #2 从 " + h1e + " 收到格末 " + cell.End
                    + " Execute=" + secondOk
                    + " hit=" + h2s + "-" + h2e
                    + (bounced ? " STALE弹回第一处" : (inside ? " 落在第二处窗内" : "")));
                Ok(run, "fs-cell-pre-second", secondOk && inside && !bounced,
                    secondOk
                        ? ("hit=" + h2s + "-" + h2e
                            + " window=" + h1e + "-" + cell.End
                            + (bounced ? " 弹回" : " 未落窗内"))
                        : "第二处 Execute=False（窗内含内表）");
            }
            catch (Exception ex)
            {
                run.CaseFail("fs-cell-pre-first", ex.Message);
                run.CaseFail("fs-cell-pre-second", "写/探针失败");
            }
            finally
            {
                CloseDoc(doc);
            }
        }

        private static void RunCellPlusCr(TestRun run, Word.Application app)
        {
            Word.Document doc = null;
            try
            {
                doc = app.Documents.Add();
                Word.Table table = doc.Tables.Add(doc.Range(0, 0), 1, 1);
                Word.Cell cell = table.Cell(1, 1);
                cell.Range.Text = "封面甲。";
                Word.Range cr = cell.Range;
                string cellText = cr.Text ?? "";
                string body = cellText.TrimEnd('\r', '\a');
                string withCr = body + "\r";
                string findCode = WordAddIn1.WordRangeFinder.ConvertNewlinesToWordCodes(withCr);

                Console.WriteLine("  --- 格文 + \\r / ^p ---");
                Console.WriteLine("    Cell.Text=" + QuoteSpecial(cellText)
                    + "  当 display 的格文=" + QuoteSpecial(body));
                Word.Characters chars = cr.Characters;
                if (chars.Count > 0)
                {
                    Word.Range last = chars[chars.Count];
                    Console.WriteLine("    末项 Characters[" + chars.Count + "] "
                        + last.Start + "-" + last.End + "  " + QuoteSpecial(last.Text));
                }

                bool memCr = cellText.IndexOf(withCr, StringComparison.Ordinal) >= 0;
                Console.WriteLine("    内存含「格文+\\r」? " + memCr
                    + "  含「格文+\\r\\a」? "
                    + (cellText.IndexOf(body + "\r\a", StringComparison.Ordinal) >= 0));

                bool raw = TryFind(doc, cr.Start, cr.End, withCr, out int rs, out int re);
                Console.WriteLine("    Find 原文「格文+\\r」 Execute=" + raw
                    + (raw ? (" hit=" + rs + "-" + re) : ""));

                bool caret = TryFind(doc, cr.Start, cr.End, findCode, out int cs, out int ce);
                Console.WriteLine("    Find 「" + findCode + "」 Execute=" + caret
                    + (caret ? (" hit=" + cs + "-" + ce) : ""));

                Ok(run, "fs-cell-plus-p", memCr && caret,
                    caret
                        ? ("^p hit=" + cs + "-" + ce + " 内存+\\r=" + memCr)
                        : ("^p 未命中 内存+\\r=" + memCr + " raw=" + raw));
            }
            catch (Exception ex)
            {
                run.CaseFail("fs-cell-plus-p", ex.Message);
            }
            finally
            {
                CloseDoc(doc);
            }
        }

        private static void RunCellPlainAmbigFind(TestRun run, Word.Application app)
        {
            Word.Document doc = null;
            try
            {
                doc = app.Documents.Add();
                Word.Table outer = doc.Tables.Add(doc.Range(0, 0), 1, 1);
                Word.Cell host = outer.Cell(1, 1);
                host.Range.Text = Needle + "间隔乙。" + Needle;
                Word.Range cell = host.Range;
                Console.WriteLine("  --- 格内无内表，仅「歧义甲。间隔乙。歧义甲。」 ---");
                Console.WriteLine("    cell=" + cell.Start + "-" + cell.End
                    + "  Text=" + QuoteSpecial(cell.Text)
                    + "  Cell.Tables=" + host.Tables.Count);

                bool firstOk = TryFind(doc, cell.Start, cell.End, Needle, out int h1s, out int h1e);
                bool firstIn = firstOk && h1s >= cell.Start && h1e <= cell.End;
                Console.WriteLine("    #1 整格搜 Execute=" + firstOk
                    + " hit=" + h1s + "-" + h1e
                    + (firstIn ? " 落在格内" : ""));
                Ok(run, "fs-cell-plain-first", firstOk && firstIn,
                    firstOk ? ("hit=" + h1s + "-" + h1e) : "整格搜第一处失败");

                if (!firstOk)
                {
                    run.CaseFail("fs-cell-plain-second", "没有第一处，无法推进");
                    return;
                }

                bool secondOk = TryFind(doc, h1e, cell.End, Needle, out int h2s, out int h2e);
                bool bounced = secondOk && h2s < h1e;
                bool inside = secondOk && h2s >= h1e && h2e <= cell.End;
                Console.WriteLine("    #2 从 " + h1e + " 收到格末 " + cell.End
                    + " Execute=" + secondOk
                    + " hit=" + h2s + "-" + h2e
                    + (bounced ? " STALE弹回第一处" : (inside ? " 落在第二处窗内" : "")));
                Ok(run, "fs-cell-plain-second", secondOk && inside && !bounced,
                    secondOk
                        ? ("hit=" + h2s + "-" + h2e
                            + " window=" + h1e + "-" + cell.End
                            + (bounced ? " 弹回" : " 未落窗内"))
                        : "第二处 Execute=False");
            }
            catch (Exception ex)
            {
                run.CaseFail("fs-cell-plain-first", ex.Message);
                run.CaseFail("fs-cell-plain-second", "写/探针失败");
            }
            finally
            {
                CloseDoc(doc);
            }
        }

        private static void RunCellPlainWalk(TestRun run, Word.Application app)
        {
            Word.Document doc = null;
            try
            {
                doc = app.Documents.Add();
                Word.Table outer = doc.Tables.Add(doc.Range(0, 0), 1, 1);
                Word.Cell host = outer.Cell(1, 1);
                host.Range.Text = Needle + "间隔乙。" + Needle;
                Word.Range cell = host.Range;
                Console.WriteLine("  --- 格内两处同文，逐字推进 ---");
                Console.WriteLine("    cell=" + cell.Start + "-" + cell.End
                    + "  Text=" + QuoteSpecial(cell.Text));

                int firstSecondAt = -1;
                int lastStaleAt = -1;
                var sb = new StringBuilder();
                for (int c = cell.Start; c < cell.End; c++)
                {
                    bool ok = TryFind(doc, c, cell.End, Needle, out int hs, out int he);
                    string tag;
                    if (!ok)
                    {
                        tag = "MISS";
                    }
                    else if (hs < c)
                    {
                        tag = "STALE hit=" + hs + "-" + he;
                        lastStaleAt = c;
                    }
                    else
                    {
                        tag = "ok hit=" + hs + "-" + he;
                        if (firstSecondAt < 0 && hs > cell.Start)
                        {
                            firstSecondAt = c;
                        }
                    }

                    sb.Append("    cursor=").Append(c).Append("  ").AppendLine(tag);
                }

                Console.WriteLine(sb.ToString());
                Console.WriteLine("    首次命中第二处的 cursor="
                    + (firstSecondAt < 0 ? "从未" : firstSecondAt.ToString())
                    + "  最后一次 STALE cursor=" + lastStaleAt);
                Ok(run, "fs-cell-plain-walk", firstSecondAt >= 0,
                    firstSecondAt < 0
                        ? "逐字推到格末也从未命中第二处"
                        : ("cursor=" + firstSecondAt + " 起才能中第二处"));
            }
            catch (Exception ex)
            {
                run.CaseFail("fs-cell-plain-walk", ex.Message);
            }
            finally
            {
                CloseDoc(doc);
            }
        }

        private static void RunCellCharacters(TestRun run, Word.Application app)
        {
            Word.Document doc = null;
            try
            {
                doc = app.Documents.Add();
                Word.Table outer = doc.Tables.Add(doc.Range(0, 0), 1, 1);
                Word.Cell host = outer.Cell(1, 1);
                host.Range.Text = Needle + "间隔乙。" + Needle;
                Word.Range cell = host.Range;
                Word.Characters chars = cell.Characters;
                int n = chars.Count;
                Console.WriteLine("  --- Cell.Range.Characters ---");
                Console.WriteLine("    cell=" + cell.Start + "-" + cell.End
                    + "  Characters.Count=" + n
                    + "  Text=" + QuoteSpecial(cell.Text));
                var sb = new StringBuilder();
                for (int i = 1; i <= n; i++)
                {
                    Word.Range ch = chars[i];
                    sb.Append("    [").Append(i).Append("] ")
                        .Append(ch.Start).Append("-").Append(ch.End)
                        .Append("  ").AppendLine(QuoteSpecial(ch.Text));
                }

                Console.WriteLine(sb.ToString());
                Ok(run, "fs-cell-chars", n > 0, "Characters.Count=" + n);
            }
            catch (Exception ex)
            {
                run.CaseFail("fs-cell-chars", ex.Message);
            }
            finally
            {
                CloseDoc(doc);
            }
        }

        private static void RunDocCharacters(TestRun run, Word.Application app)
        {
            Word.Document doc = null;
            try
            {
                doc = app.Documents.Add();
                doc.Content.Text = Needle + "\r间隔乙。\r" + Needle + "\r";
                int end = doc.Content.End;
                Word.Range at = doc.Range(end - 1, end - 1);
                Word.Table table = doc.Tables.Add(at, 1, 2);
                table.Cell(1, 1).Range.Text = Needle;
                table.Cell(1, 2).Range.Text = "表内丙。";

                Word.Characters chars = doc.Content.Characters;
                int n = chars.Count;
                Console.WriteLine("  --- Document.Content.Characters ---");
                Console.WriteLine("    Content.End=" + doc.Content.End + "  Characters.Count=" + n);
                int show = n < 8 ? n : 8;
                for (int i = 1; i <= show; i++)
                {
                    Word.Range ch = chars[i];
                    Console.WriteLine("    [" + i + "] " + ch.Start + "-" + ch.End
                        + "  " + QuoteSpecial(ch.Text));
                }

                if (n > show)
                {
                    Word.Range last = chars[n];
                    Console.WriteLine("    …");
                    Console.WriteLine("    [" + n + "] " + last.Start + "-" + last.End
                        + "  " + QuoteSpecial(last.Text));
                }

                Ok(run, "fs-doc-chars", n > 0, "Count=" + n);
            }
            catch (Exception ex)
            {
                run.CaseFail("fs-doc-chars", ex.Message);
            }
            finally
            {
                CloseDoc(doc);
            }
        }

        private static void RunCharsAbut(TestRun run, Word.Application app)
        {
            Word.Document doc = null;
            try
            {
                doc = app.Documents.Add();
                doc.Content.Text = Needle + "\r间隔乙。\r" + Needle + "\r";
                int end = doc.Content.End;
                Word.Range at = doc.Range(end - 1, end - 1);
                Word.Table table = doc.Tables.Add(at, 1, 2);
                table.Cell(1, 1).Range.Text = Needle;
                table.Cell(1, 2).Range.Text = "表内丙。";

                string cellReport = CheckAbut("格", table.Cell(1, 1).Range);
                string docReport = CheckAbut("整篇", doc.Content);
                Console.WriteLine("  --- Characters 是否收尾相连 ---");
                Console.WriteLine(cellReport);
                Console.WriteLine(docReport);
                bool ok = cellReport.IndexOf("不相连: GAP=0 OVER=0", StringComparison.Ordinal) >= 0
                    && docReport.IndexOf("不相连: GAP=0 OVER=0", StringComparison.Ordinal) >= 0;
                Ok(run, "fs-chars-abut", ok, "有缺口或重叠，见日志");
            }
            catch (Exception ex)
            {
                run.CaseFail("fs-chars-abut", ex.Message);
            }
            finally
            {
                CloseDoc(doc);
            }
        }

        private const string RealDocPath =
            @"D:\documents\易写\易写产品试用2\2. 体验素材\1.重庆工程学院课程思政示范课程建设项目结题验收报告.doc";

        private const string RealDocxPath =
            @"D:\documents\易写\易写产品试用2\2. 体验素材\1.重庆工程学院课程思政示范课程建设项目结题验收报告.docx";

        private static void RunPageBreakChar(TestRun run, Word.Application app)
        {
            Word.Document doc = null;
            try
            {
                doc = app.Documents.Add();
                doc.Content.Text = "封面甲。";
                Word.Range at = doc.Range(doc.Content.End - 1, doc.Content.End - 1);
                object breakType = Word.WdBreakType.wdPageBreak;
                at.InsertBreak(ref breakType);
                Word.Range after = doc.Range(doc.Content.End - 1, doc.Content.End - 1);
                after.InsertAfter("正文乙。");

                string text = doc.Content.Text ?? "";
                Console.WriteLine("  --- 分页符在 Content.Text / Characters ---");
                Console.WriteLine("    Content.End=" + doc.Content.End
                    + "  Text.Length=" + text.Length
                    + "  Text=" + QuoteSpecial(text));

                Word.Characters chars = doc.Content.Characters;
                int n = chars.Count;
                Console.WriteLine("    Characters.Count=" + n);
                int ffIndex = -1;
                for (int i = 1; i <= n; i++)
                {
                    Word.Range ch = chars[i];
                    string t = ch.Text ?? "";
                    Console.WriteLine("    [" + i + "] " + ch.Start + "-" + ch.End
                        + "  " + QuoteSpecial(t));
                    if (t.IndexOf('\f') >= 0)
                    {
                        ffIndex = i;
                    }
                }

                bool inText = text.IndexOf('\f') >= 0;
                Ok(run, "fs-page-char", inText && ffIndex > 0,
                    inText
                        ? ("Text 有 \\f，Characters[" + ffIndex + "]")
                        : "Content.Text 里没有 \\f");
            }
            catch (Exception ex)
            {
                run.CaseFail("fs-page-char", ex.Message);
            }
            finally
            {
                CloseDoc(doc);
            }
        }

        private static void RunRealDocDisplayAlign(TestRun run, Word.Application app)
        {
            string path = !string.IsNullOrEmpty(AlignDocPath)
                ? AlignDocPath
                : (System.IO.File.Exists(RealDocxPath) ? RealDocxPath : RealDocPath);
            if (!System.IO.File.Exists(path))
            {
                run.CaseSkip("fs-display-align", "文件不在: " + path);
                return;
            }

            Word.Document doc = null;
            bool weOpened = false;
            bool oldUpdate = true;
            try
            {
                string norm = System.IO.Path.GetFullPath(path);
                foreach (Word.Document d in app.Documents)
                {
                    try
                    {
                        if (string.Equals(
                            System.IO.Path.GetFullPath(d.FullName),
                            norm,
                            StringComparison.OrdinalIgnoreCase))
                        {
                            doc = d;
                            break;
                        }
                    }
                    catch
                    {
                    }
                }

                if (doc == null)
                {
                    object fileName = path;
                    object readOnly = true;
                    object missing = System.Reflection.Missing.Value;
                    doc = app.Documents.Open(
                        ref fileName,
                        ref missing,
                        ref readOnly);
                    weOpened = true;
                }

                oldUpdate = app.ScreenUpdating;
                app.ScreenUpdating = false;
                try
                {
                    app.DisplayAlerts = Word.WdAlertLevel.wdAlertsNone;
                }
                catch
                {
                }

                Console.WriteLine("  --- display 转换后对 Content.Text ---");
                Console.WriteLine("    path=" + path);
                Console.Out.Flush();

                var sw = System.Diagnostics.Stopwatch.StartNew();
                string story = doc.Content.Text ?? "";
                Console.WriteLine("    Content.Text Length=" + story.Length
                    + "  用时=" + sw.Elapsed.TotalSeconds.ToString("F2") + "s");
                Console.Out.Flush();

                sw.Restart();
                Console.WriteLine("    ProcessDocument 开始…");
                Console.Out.Flush();
                var options = ProcessDocumentOptions.ForGetDocumentContent("fs-dalign");
                options.ForceRefresh = true;
                WordDocumentExtractor.ProcessingResult processed =
                    WordDocumentExtractor.ProcessDocument(doc, options);
                Console.WriteLine("    ProcessDocument 结束 "
                    + sw.Elapsed.TotalSeconds.ToString("F2") + "s");
                Console.Out.Flush();
                string display = "";
                var paras = DocumentState.ChunkParagraphDisplayContents;
                if (paras != null)
                {
                    var sb = new StringBuilder();
                    foreach (string p in paras)
                    {
                        sb.Append(p ?? "");
                    }

                    display = sb.ToString();
                }

                if (display.Length == 0 && processed != null && processed.ChunkInfo != null)
                {
                    var sb = new StringBuilder();
                    foreach (WordDocumentExtractor.ChunkInfo chunk in processed.ChunkInfo)
                    {
                        if (!string.IsNullOrEmpty(chunk.ParagraphDisplayContent))
                        {
                            sb.Append(chunk.ParagraphDisplayContent);
                        }
                        else
                        {
                            sb.Append(chunk.DisplayContent ?? "");
                        }
                    }

                    display = sb.ToString();
                }

                Console.WriteLine("    paragraph display Length=" + display.Length
                    + "  ProcessDocument=" + sw.Elapsed.TotalSeconds.ToString("F2") + "s"
                    + "  method=" + (processed != null ? processed.ProcessingMethod : "?"));
                string readText = processed != null ? (processed.ReadText ?? "") : "";
                string sentDisplay = "";
                if (processed != null && processed.ChunkInfo != null)
                {
                    var ssb = new StringBuilder();
                    foreach (WordDocumentExtractor.ChunkInfo chunk in processed.ChunkInfo)
                    {
                        ssb.Append(chunk.DisplayContent ?? "");
                    }

                    sentDisplay = ssb.ToString();
                }

                Console.WriteLine("    ReadText 含 <image=" + (readText.IndexOf("<image", StringComparison.OrdinalIgnoreCase) >= 0)
                    + "  含 <Chart=" + (readText.IndexOf("<Chart", StringComparison.OrdinalIgnoreCase) >= 0)
                    + "  Length=" + readText.Length);
                Console.WriteLine("    sentence display 含 <image=" + (sentDisplay.IndexOf("<image", StringComparison.OrdinalIgnoreCase) >= 0)
                    + "  Length=" + sentDisplay.Length);
                Console.WriteLine("    paragraph display 含 <image=" + (display.IndexOf("<image", StringComparison.OrdinalIgnoreCase) >= 0));
                Console.Out.Flush();

                string converted = DisplayToStory(display);
                string displayMarks = ConvertTables(StripDisplayWrappers(display))
                    .Replace("\n", "\x0B");
                int dispCr, dispB, dispF, dispCrA;
                int dispEnds = WalkConvertedParaEnds(
                    displayMarks, out dispCr, out dispB, out dispF, out dispCrA);
                int nPara = doc.Paragraphs.Count;
                Console.WriteLine("    从文首转换 段号=" + dispEnds
                    + "  \\r=" + dispCr
                    + "  其中\\r\\a=" + dispCrA
                    + "  ^b=" + dispB
                    + "  \\f=" + dispF
                    + "  Paragraphs.Count=" + nPara
                    + (dispEnds == nPara ? "  相等" : "  对不上"));
                Ok(run, "fs-display-pcount", dispEnds == nPara,
                    "display段尾=" + dispEnds + " Count=" + nPara
                    + " (\\r=" + dispCr + " ^b=" + dispB + " \\f=" + dispF + ")");
                DumpPSlotVsCount(run, display, dispEnds, nPara);
                int slashRemoved;
                string storyNoImg = StripComPictureChars(doc, story ?? "", out slashRemoved);
                int leftoverSoH = 0;
                foreach (char ch in storyNoImg)
                {
                    if (ch == '\x01')
                    {
                        leftoverSoH++;
                    }
                }

                Console.WriteLine("    转换后(去 image 标签) Length=" + converted.Length);
                Console.WriteLine("    story 只去图占位 / 和 \\x01 Length=" + storyNoImg.Length
                    + "  原=" + story.Length
                    + "  去 /=" + slashRemoved
                    + "  仍剩 \\x01=" + leftoverSoH);
                ReportStoryDiff("双方去图", converted, storyNoImg);
                DumpImageGap(display, converted, story);
                bool thesisDump = string.Equals(
                    System.IO.Path.GetFullPath(path),
                    System.IO.Path.GetFullPath(RealDocxPath),
                    StringComparison.OrdinalIgnoreCase);
                if (thesisDump)
                {
                    DumpFig1Com(doc, story);
                    DumpLecturerGap(display, story, doc);
                }

                bool ok = string.Equals(converted, storyNoImg, StringComparison.Ordinal);
                Ok(run, "fs-display-align", ok,
                    ok ? "转换后与 Content.Text 一致" : "仍有对不齐，见日志首差");
                CheckEncodedSentencesNoImage(run, sentDisplay);
            }
            catch (Exception ex)
            {
                run.CaseFail("fs-display-align", ex.GetType().Name + ": " + ex.Message);
            }
            finally
            {
                try
                {
                    app.ScreenUpdating = oldUpdate;
                }
                catch
                {
                }

                if (weOpened && doc != null)
                {
                    CloseDoc(doc);
                }
            }
        }

        private static readonly Regex DisplayCodeStart = new Regex(
            @"\[(?:P|S)_[^\]]+?(?:-start| start)\]|\[T_[0-9A-Za-z]+_r[0-9]+_c[0-9]+-start\]",
            RegexOptions.Compiled);
        private static readonly Regex DisplayCodeEnd = new Regex(
            @"\[(?:P|S)_[^\]]+?(?:-end| end)\]|\[T_[0-9A-Za-z]+_r[0-9]+_c[0-9]+-end\]",
            RegexOptions.Compiled);
        private static readonly Regex DisplayInTable = new Regex(
            @"\[in_table=[^\]]+\]", RegexOptions.Compiled);
        private static readonly Regex DisplayImage = new Regex(
            @"<image\b[^>]*/?>", RegexOptions.Compiled | RegexOptions.IgnoreCase);
        private static readonly Regex DisplayChart = new Regex(
            @"<Chart\b[^>]*>[\s\S]*?</Chart\s*>|<Chart\b[^>]*/>",
            RegexOptions.Compiled | RegexOptions.IgnoreCase);
        private static readonly Regex DisplayAttrInt = new Regex(
            @"(colspan|rowspan)\s*=\s*""?(\d+)""?",
            RegexOptions.Compiled | RegexOptions.IgnoreCase);

        private static string StripComPictureChars(Word.Document doc, string story, out int slashRemoved)
        {
            slashRemoved = 0;
            if (string.IsNullOrEmpty(story))
            {
                return "";
            }

            string s = story.Replace("\x01", "");
            Word.InlineShapes pics = doc.InlineShapes;
            int n = pics.Count;
            for (int i = 1; i <= n; i++)
            {
                Word.Range picRange;
                try
                {
                    picRange = pics[i].Range;
                }
                catch
                {
                    continue;
                }

                string picText = picRange.Text ?? "";
                if (picText.IndexOf('/') < 0)
                {
                    continue;
                }

                int from = picRange.Start - 16;
                if (from < doc.Content.Start)
                {
                    from = doc.Content.Start;
                }

                string prefix;
                try
                {
                    prefix = (from < picRange.Start)
                        ? (doc.Range(from, picRange.Start).Text ?? "")
                        : "";
                }
                catch
                {
                    prefix = "";
                }

                prefix = prefix.Replace("\x01", "");
                string needle = prefix + "/";
                int at = string.IsNullOrEmpty(prefix)
                    ? s.IndexOf('/')
                    : s.IndexOf(needle, StringComparison.Ordinal);
                if (at < 0)
                {
                    continue;
                }

                int slashAt = string.IsNullOrEmpty(prefix) ? at : at + prefix.Length;
                if (slashAt >= 0 && slashAt < s.Length && s[slashAt] == '/')
                {
                    s = s.Remove(slashAt, 1);
                    slashRemoved++;
                }
            }

            return s;
        }

        private static string DisplayToStory(string display)
        {
            if (string.IsNullOrEmpty(display))
            {
                return "";
            }

            string s = StripDisplayWrappers(display);
            return ConvertTables(s).Replace("\n", "\x0B").Replace("^b", "\f");
        }

        private static readonly Regex DisplayPBlock = new Regex(
            @"\[P_[^\]]+?(?:-start| start)\](.*?)\[P_[^\]]+?(?:-end| end)\]",
            RegexOptions.Compiled | RegexOptions.Singleline);

        private static void DumpPSlotVsCount(
            TestRun run, string paragraphDisplay, int dispEnds, int nPara)
        {
            var blocks = DisplayPBlock.Matches(paragraphDisplay ?? "");
            int pSlots = blocks.Count;
            int caretB = 0;
            int caretBGlued = 0;
            int ff = 0;
            int noCr = 0;
            for (int i = 0; i < blocks.Count; i++)
            {
                string inner = blocks[i].Groups[1].Value ?? "";
                if (inner.IndexOf("^b", StringComparison.Ordinal) >= 0)
                {
                    caretB++;
                    if (!inner.EndsWith("^b", StringComparison.Ordinal)
                        && !inner.EndsWith("^b\r", StringComparison.Ordinal))
                    {
                        caretBGlued++;
                        if (caretBGlued <= 3)
                        {
                            Console.WriteLine("    P_ 槽内 ^b 粘上下段  #" + (i + 1)
                                + "  " + QuoteSpecial(inner.Length > 40 ? inner.Substring(0, 40) : inner));
                        }
                    }
                }

                if (inner.IndexOf('\f') >= 0)
                {
                    ff++;
                }

                if (inner.IndexOf('\r') < 0)
                {
                    noCr++;
                }
            }

            int splitOnCr = SplitTextByParagraphs.Split(
                ConvertTables(StripDisplayWrappers(paragraphDisplay ?? ""))
                    .Replace("\n", "\x0B")).Count;

            Console.WriteLine("    P_ 槽数=" + pSlots
                + "  槽内含^b=" + caretB
                + "  ^b粘上下段=" + caretBGlued
                + "  槽内含\\f=" + ff
                + "  槽无\\r=" + noCr
                + "  转换后再按\\r切=" + splitOnCr);
            Console.WriteLine("    对照  P_槽=" + pSlots
                + "  display段尾(\\r+^b)=" + dispEnds
                + "  Paragraphs.Count=" + nPara
                + (pSlots == dispEnds ? "  P_与数段一致" : "  P_与数段不一致"));
            Ok(run, "fs-p-slot-vs-count", pSlots == dispEnds,
                "P_槽=" + pSlots + " 数段=" + dispEnds + " Count=" + nPara);
        }

        private static int WalkConvertedParaEnds(
            string text, out int cr, out int caretB, out int ff, out int crA)
        {
            cr = 0;
            caretB = 0;
            ff = 0;
            crA = 0;
            int paraNo = 0;
            if (string.IsNullOrEmpty(text))
            {
                return 0;
            }

            for (int i = 0; i < text.Length; i++)
            {
                if (text[i] == '\r')
                {
                    cr++;
                    paraNo++;
                    if (i + 1 < text.Length && text[i + 1] == '\a')
                    {
                        crA++;
                    }
                }
                else if (text[i] == '\f')
                {
                    ff++;
                    if (i + 1 >= text.Length || text[i + 1] != '\r')
                    {
                        paraNo++;
                    }
                }
                else if (text[i] == '^'
                    && i + 1 < text.Length
                    && text[i + 1] == 'b')
                {
                    caretB++;
                    paraNo++;
                    i++;
                }
            }

            return paraNo;
        }

        private static string StripDisplayWrappers(string display)
        {
            string s = DisplayCodeStart.Replace(display, "");
            s = DisplayCodeEnd.Replace(s, "");
            s = DisplayInTable.Replace(s, "");
            s = DisplayChart.Replace(s, "");
            s = DisplayImage.Replace(s, "");
            return FormFeedText.UnescapeToChar12(s);
        }

        private static string ConvertTables(string text)
        {
            var sb = new StringBuilder();
            int i = 0;
            while (i < text.Length)
            {
                int tableAt = IndexOfIgnoreCase(text, "<table", i);
                if (tableAt < 0)
                {
                    sb.Append(text.Substring(i));
                    break;
                }

                sb.Append(text.Substring(i, tableAt - i));
                int openEnd = text.IndexOf('>', tableAt);
                if (openEnd < 0)
                {
                    sb.Append(text.Substring(tableAt));
                    break;
                }

                int close = FindMatchingClose(text, openEnd + 1, "<table", "</table>");
                if (close < 0)
                {
                    sb.Append(text.Substring(tableAt));
                    break;
                }

                string inner = text.Substring(openEnd + 1, close - (openEnd + 1));
                sb.Append(ConvertOneTable(inner));
                int after = text.IndexOf('>', close);
                i = after >= 0 ? after + 1 : close + 8;
            }

            return sb.ToString();
        }

        private static string ConvertOneTable(string inner)
        {
            var rows = ParseRows(inner);
            if (rows.Count == 0)
            {
                return "";
            }

            int extra = 0;
            foreach (List<DispCell> row in rows)
            {
                foreach (DispCell cell in row)
                {
                    extra = Math.Max(extra, cell.RowSpan - 1);
                }
            }

            int nrows = rows.Count + extra;
            var carry = new int[nrows];
            for (int r = 0; r < rows.Count; r++)
            {
                int used = 0;
                foreach (DispCell cell in rows[r])
                {
                    used += cell.ColSpan;
                    for (int dr = 1; dr < cell.RowSpan && r + dr < nrows; dr++)
                    {
                        carry[r + dr] += cell.ColSpan;
                    }
                }

                carry[r] += used;
            }

            int ncols = 0;
            for (int r = 0; r < nrows; r++)
            {
                if (carry[r] > ncols)
                {
                    ncols = carry[r];
                }
            }
            var origin = new int[nrows, ncols];
            for (int r = 0; r < nrows; r++)
            {
                for (int c = 0; c < ncols; c++)
                {
                    origin[r, c] = -1;
                }
            }

            for (int r = 0; r < rows.Count; r++)
            {
                int c = 0;
                foreach (DispCell cell in rows[r])
                {
                    while (c < ncols && origin[r, c] != -1)
                    {
                        c++;
                    }

                    if (c < ncols)
                    {
                        origin[r, c] = 0;
                    }

                    for (int dc = 1; dc < cell.ColSpan && c + dc < ncols; dc++)
                    {
                        origin[r, c + dc] = -2;
                    }

                    for (int dr = 1; dr < cell.RowSpan && r + dr < nrows; dr++)
                    {
                        if (c < ncols)
                        {
                            origin[r + dr, c] = cell.ColSpan;
                        }

                        for (int dc = 1; dc < cell.ColSpan && c + dc < ncols; dc++)
                        {
                            origin[r + dr, c + dc] = -2;
                        }
                    }

                    c += cell.ColSpan;
                }
            }

            var sb = new StringBuilder();
            for (int r = 0; r < rows.Count; r++)
            {
                int c = 0;
                int cellIdx = 0;
                while (c < ncols)
                {
                    if (origin[r, c] > 0)
                    {
                        sb.Append('\r').Append('\a');
                        c += origin[r, c];
                        continue;
                    }

                    if (origin[r, c] == -2)
                    {
                        c++;
                        continue;
                    }

                    if (cellIdx < rows[r].Count)
                    {
                        DispCell cell = rows[r][cellIdx++];
                        sb.Append(ConvertTables(cell.Inner));
                        sb.Append('\r').Append('\a');
                        c += cell.ColSpan;
                    }
                    else
                    {
                        c++;
                    }
                }

                sb.Append('\r').Append('\a');
            }

            return sb.ToString();
        }

        private sealed class DispCell
        {
            public int ColSpan = 1;
            public int RowSpan = 1;
            public string Inner = "";
        }

        private static List<List<DispCell>> ParseRows(string inner)
        {
            var rows = new List<List<DispCell>>();
            int i = 0;
            while (i < inner.Length)
            {
                int tr = IndexOfIgnoreCase(inner, "<tr", i);
                if (tr < 0)
                {
                    break;
                }

                int trEnd = inner.IndexOf('>', tr);
                if (trEnd < 0)
                {
                    break;
                }

                int trClose = FindMatchingClose(inner, trEnd + 1, "<tr", "</tr>");
                if (trClose < 0)
                {
                    break;
                }

                string rowHtml = inner.Substring(trEnd + 1, trClose - (trEnd + 1));
                rows.Add(ParseCells(rowHtml));
                int after = inner.IndexOf('>', trClose);
                i = after >= 0 ? after + 1 : trClose + 5;
            }

            return rows;
        }

        private static List<DispCell> ParseCells(string rowHtml)
        {
            var cells = new List<DispCell>();
            int i = 0;
            while (i < rowHtml.Length)
            {
                int td = IndexOfIgnoreCase(rowHtml, "<td", i);
                int th = IndexOfIgnoreCase(rowHtml, "<th", i);
                int open = MinPos(td, th);
                if (open < 0)
                {
                    break;
                }

                bool isTh = th >= 0 && (td < 0 || th < td);
                int openEnd = rowHtml.IndexOf('>', open);
                if (openEnd < 0)
                {
                    break;
                }

                string tag = rowHtml.Substring(open, openEnd - open);
                int close = FindMatchingClose(
                    rowHtml,
                    openEnd + 1,
                    isTh ? "<th" : "<td",
                    isTh ? "</th>" : "</td>");
                if (close < 0)
                {
                    break;
                }

                var cell = new DispCell
                {
                    Inner = rowHtml.Substring(openEnd + 1, close - (openEnd + 1)),
                    ColSpan = ReadSpan(tag, "colspan"),
                    RowSpan = ReadSpan(tag, "rowspan")
                };
                cells.Add(cell);
                int after = rowHtml.IndexOf('>', close);
                i = after >= 0 ? after + 1 : close + 5;
            }

            return cells;
        }

        private static int ReadSpan(string tag, string name)
        {
            Match m = DisplayAttrInt.Match(tag);
            while (m.Success)
            {
                if (string.Equals(m.Groups[1].Value, name, StringComparison.OrdinalIgnoreCase))
                {
                    int n;
                    if (int.TryParse(m.Groups[2].Value, out n) && n > 0)
                    {
                        return n;
                    }
                }

                m = m.NextMatch();
            }

            return 1;
        }

        private static int FindMatchingClose(string text, int from, string open, string close)
        {
            int depth = 1;
            int i = from;
            while (i < text.Length && depth > 0)
            {
                int nextOpen = IndexOfIgnoreCase(text, open, i);
                int nextClose = IndexOfIgnoreCase(text, close, i);
                if (nextClose < 0)
                {
                    return -1;
                }

                if (nextOpen >= 0 && nextOpen < nextClose)
                {
                    depth++;
                    i = nextOpen + open.Length;
                }
                else
                {
                    depth--;
                    if (depth == 0)
                    {
                        return nextClose;
                    }

                    i = nextClose + close.Length;
                }
            }

            return -1;
        }

        private static int IndexOfIgnoreCase(string text, string value, int start)
        {
            return text.IndexOf(value, start, StringComparison.OrdinalIgnoreCase);
        }

        private static int MinPos(int a, int b)
        {
            if (a < 0)
            {
                return b;
            }

            if (b < 0)
            {
                return a;
            }

            return Math.Min(a, b);
        }

        private static readonly Regex SentencePair = new Regex(
            @"\[S_([0-9a-zA-Z]{5})-start\](.*?)\[S_\1-end\]",
            RegexOptions.Compiled | RegexOptions.Singleline);

        private static void CheckEncodedSentencesNoImage(TestRun run, string sentDisplay)
        {
            Console.WriteLine("  --- 编码句子是否含 <image ---");
            int n = 0;
            int bad = 0;
            if (!string.IsNullOrEmpty(sentDisplay))
            {
                foreach (Match m in SentencePair.Matches(sentDisplay))
                {
                    n++;
                    string code = "S_" + m.Groups[1].Value;
                    string inner = m.Groups[2].Value ?? "";
                    if (inner.IndexOf("<image", StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        bad++;
                        if (bad <= 8)
                        {
                            Console.WriteLine("    display " + code + " 含 image  "
                                + QuoteSpecial(Peek(inner, 0, 80)));
                        }
                    }
                }
            }

            int mapN = 0;
            int mapBad = 0;
            foreach (IReadOnlyList<string> row in DocumentState.SentenceNameMapping)
            {
                if (row == null || row.Count < 3)
                {
                    continue;
                }

                mapN++;
                string content = row[2] ?? "";
                if (content.IndexOf("<image", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    mapBad++;
                    if (mapBad <= 8)
                    {
                        Console.WriteLine("    mapping " + row[0] + " 含 image  "
                            + QuoteSpecial(Peek(content, 0, 80)));
                    }
                }
            }

            Console.WriteLine("    display S_ 句数=" + n + "  含 image=" + bad
                + "  mapping 句数=" + mapN + "  含 image=" + mapBad);
            bool ok = bad == 0 && mapBad == 0 && n > 0;
            Ok(run, "fs-sent-no-image", ok,
                ok
                    ? ("S_ 均无 image，display=" + n + " mapping=" + mapN)
                    : ("display 含 image=" + bad + " mapping 含 image=" + mapBad
                        + " display句=" + n));
        }

        private static void DumpImageGap(string display, string converted, string story)
        {
            int nTag = DisplayImage.Matches(display).Count;
            int nSoH = 0;
            foreach (char ch in story)
            {
                if (ch == '\x01')
                {
                    nSoH++;
                }
            }

            Console.WriteLine("    display <image> 数=" + nTag + "  story \\x01 数=" + nSoH);
            int sohAt = (story ?? "").IndexOf('\x01');
            if (sohAt >= 0)
            {
                string around = Peek(story, Math.Max(0, sohAt - 8), 20);
                Console.WriteLine("    \\x01 下标=" + sohAt + "  附近=" + QuoteSpecial(around));
                var codes = new StringBuilder();
                int from = Math.Max(0, sohAt - 2);
                int to = Math.Min(story.Length, sohAt + 6);
                for (int i = from; i < to; i++)
                {
                    if (codes.Length > 0)
                    {
                        codes.Append(' ');
                    }

                    codes.Append(i == sohAt ? "[" : "")
                        .Append(((int)story[i]).ToString("X2"))
                        .Append(i == sohAt ? "]" : "");
                }

                Console.WriteLine("    \\x01 邻码=" + codes);
            }

            bool hasSoHSlashCr = (story ?? "").IndexOf("\x01/\r", StringComparison.Ordinal) >= 0;
            bool hasSlashCrAfterXlsx = false;
            int xlsxAt = (story ?? "").IndexOf("知识图谱.xlsx", StringComparison.Ordinal);
            if (xlsxAt >= 0)
            {
                string tail = Peek(story, xlsxAt, 40);
                hasSlashCrAfterXlsx = tail.IndexOf("/\r", StringComparison.Ordinal) >= 0;
                Console.WriteLine("    story 图附近: " + QuoteSpecial(tail));
            }

            Console.WriteLine("    含 \\x01/\\r=" + hasSoHSlashCr
                + "  图附近含 /\\r=" + hasSlashCrAfterXlsx);
            int at = display.IndexOf("知识图谱.xlsx", StringComparison.Ordinal);
            if (at >= 0)
            {
                Console.WriteLine("    display 图附近: "
                    + QuoteSpecial(Peek(display, Math.Max(0, at - 40), 160)));
            }

            int cAt = converted.IndexOf("知识图谱.xlsx", StringComparison.Ordinal);
            if (cAt >= 0)
            {
                Console.WriteLine("    conv 图附近: "
                    + QuoteSpecial(Peek(converted, Math.Max(0, cAt - 8), 40)));
            }
        }

        private static void DumpFig1Com(Word.Document doc, string story)
        {
            Console.WriteLine("  --- 图1 前一段 COM ---");
            Word.Paragraphs paras = doc.Paragraphs;
            int n = paras.Count;
            int shown = 0;
            for (int i = 1; i <= n && shown < 2; i++)
            {
                string t = paras[i].Range.Text ?? "";
                if (t.IndexOf("图1.", StringComparison.Ordinal) < 0)
                {
                    continue;
                }

                if (i > 1)
                {
                    DumpOneParaCom(paras[i - 1], "图1前");
                }

                DumpOneParaCom(paras[i], "图1段");
                shown++;
            }

            int slashAt = (story ?? "").IndexOf("/\r图1", StringComparison.Ordinal);
            Console.WriteLine("    story 中 /\\r图1 下标=" + slashAt);
        }

        private static void DumpOneParaCom(Word.Paragraph para, string title)
        {
            Word.Range rng = para.Range;
            Console.WriteLine("    " + title + " Range=" + rng.Start + "-" + rng.End
                + "  Text=" + QuoteSpecial(rng.Text)
                + "  InlineShapes=" + rng.InlineShapes.Count
                + "  Fields=" + rng.Fields.Count);
            if (rng.InlineShapes.Count > 0)
            {
                Word.InlineShape pic = rng.InlineShapes[1];
                Console.WriteLine("    " + title + " InlineShape Type=" + pic.Type
                    + "  Width=" + pic.Width.ToString("F1"));
            }

            Word.Characters chars = rng.Characters;
            int take = Math.Min(4, chars.Count);
            for (int i = 1; i <= take; i++)
            {
                string ct = chars[i].Text ?? "";
                int code = ct.Length > 0 ? (int)ct[0] : -1;
                Console.WriteLine("    " + title + " Characters[" + i + "] code="
                    + code.ToString("X2") + "  " + QuoteSpecial(ct)
                    + "  span=" + (chars[i].End - chars[i].Start));
            }
        }

        private static void DumpLecturerGap(string display, string story, Word.Document doc)
        {
            Console.WriteLine("  --- 讲师附近：display 表结构 vs COM ---");
            int dAt = display.IndexOf("讲师", StringComparison.Ordinal);
            if (dAt >= 0)
            {
                int from = Math.Max(0, dAt - 180);
                int take = Math.Min(420, display.Length - from);
                Console.WriteLine("    display: " + QuoteSpecial(display.Substring(from, take)));
            }

            int sAt = story.IndexOf("讲师", StringComparison.Ordinal);
            if (sAt >= 0)
            {
                Console.WriteLine("    story : " + QuoteSpecial(Peek(story, Math.Max(0, sAt - 20), 80)));
            }

            if (doc.Tables.Count > 0)
            {
                Word.Table table = doc.Tables[1];
                Console.WriteLine("    COM 表1 " + table.Rows.Count + "×" + table.Columns.Count
                    + "  Range=" + table.Range.Start + "-" + table.Range.End);
                Console.WriteLine("    COM 表1.Range.Text=" + QuoteSpecial(table.Range.Text ?? ""));
                int rMax = Math.Min(6, table.Rows.Count);
                int cMax = Math.Min(6, table.Columns.Count);
                for (int r = 1; r <= rMax; r++)
                {
                    for (int c = 1; c <= cMax; c++)
                    {
                        try
                        {
                            Word.Range cr = table.Cell(r, c).Range;
                            Console.WriteLine("    COM[" + r + "," + c + "] "
                                + cr.Start + "-" + cr.End + "  "
                                + QuoteSpecial(cr.Text));
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine("    COM[" + r + "," + c + "] "
                                + ex.GetType().Name);
                        }
                    }
                }
            }
        }

        private static void ReportStoryDiff(string title, string converted, string story)
        {
            Console.WriteLine("    --- " + title + " ---");
            if (string.Equals(converted, story, StringComparison.Ordinal))
            {
                Console.WriteLine("    完全一致 Length=" + converted.Length);
                return;
            }

            int i = FirstDiff(converted, story);
            Console.WriteLine("    不一致  conv=" + converted.Length
                + "  story=" + story.Length
                + "  首差=" + i);
            Console.WriteLine("    conv  " + QuoteSpecial(Peek(converted, Math.Max(0, i - 12), 36)));
            Console.WriteLine("    story " + QuoteSpecial(Peek(story, Math.Max(0, i - 12), 36)));
        }

        private static void RunSplitConvertedVsParas(TestRun run, Word.Application app)
        {
            string path = ResolveAlignDocPath();
            if (!System.IO.File.Exists(path))
            {
                run.CaseSkip("fs-p-split", "文件不在: " + path);
                return;
            }

            Word.Document doc = null;
            bool weOpened = false;
            bool oldUpdate = true;
            try
            {
                string norm = System.IO.Path.GetFullPath(path);
                foreach (Word.Document d in app.Documents)
                {
                    try
                    {
                        if (string.Equals(
                            System.IO.Path.GetFullPath(d.FullName),
                            norm,
                            StringComparison.OrdinalIgnoreCase))
                        {
                            doc = d;
                            break;
                        }
                    }
                    catch
                    {
                    }
                }

                if (doc == null)
                {
                    object fileName = path;
                    object readOnly = true;
                    object missing = System.Reflection.Missing.Value;
                    doc = app.Documents.Open(
                        ref fileName,
                        ref missing,
                        ref readOnly);
                    weOpened = true;
                }

                oldUpdate = app.ScreenUpdating;
                app.ScreenUpdating = false;
                try
                {
                    app.DisplayAlerts = Word.WdAlertLevel.wdAlertsNone;
                }
                catch
                {
                }

                Console.WriteLine("  --- 转换后按段尾切开 vs Paragraphs ---");
                Console.WriteLine("    path=" + path);
                Console.Out.Flush();

                var options = ProcessDocumentOptions.ForGetDocumentContent("fs-psplit");
                options.ForceRefresh = true;
                var sw = System.Diagnostics.Stopwatch.StartNew();
                WordDocumentExtractor.ProcessingResult processed =
                    WordDocumentExtractor.ProcessDocument(doc, options);
                Console.WriteLine("    ProcessDocument " + sw.Elapsed.TotalSeconds.ToString("F2") + "s");
                Console.Out.Flush();

                string display = LoadParagraphDisplay(processed);
                string marks = ConvertTables(StripDisplayWrappers(display)).Replace("\n", "\x0B");
                List<string> pieces = SplitConvertedByParaEnds(marks);
                int nPara = doc.Paragraphs.Count;
                Console.WriteLine("    切开=" + pieces.Count + "  Paragraphs.Count=" + nPara
                    + (pieces.Count == nPara ? "  条数相等" : "  条数对不上"));
                Console.Out.Flush();

                int match = 0;
                int mismatch = 0;
                int n = pieces.Count < nPara ? pieces.Count : nPara;
                sw.Restart();
                for (int i = 0; i < n; i++)
                {
                    string com = "";
                    try
                    {
                        com = doc.Paragraphs[i + 1].Range.Text ?? "";
                    }
                    catch (Exception ex)
                    {
                        mismatch++;
                        if (mismatch <= 6)
                        {
                            Console.WriteLine("    取段失败 [" + (i + 1) + "] " + ex.Message);
                        }

                        continue;
                    }

                    com = com.Replace("\x01", "");
                    string got = pieces[i].Replace("^b", "\f");
                    if (string.Equals(got, com, StringComparison.Ordinal))
                    {
                        match++;
                    }
                    else
                    {
                        mismatch++;
                        if (mismatch <= 8)
                        {
                            Console.WriteLine("    对不上 片[" + (i + 1) + "]");
                            Console.WriteLine("      切  " + QuoteSpecial(Peek(got, 0, 40)));
                            Console.WriteLine("      COM " + QuoteSpecial(Peek(com, 0, 40)));
                        }
                    }
                }

                if (pieces.Count != nPara)
                {
                    mismatch += Math.Abs(pieces.Count - nPara);
                }

                Console.WriteLine("    核对  用时=" + sw.Elapsed.TotalSeconds.ToString("F2") + "s"
                    + "  对上=" + match
                    + "  对不上=" + mismatch);
                bool ok = pieces.Count == nPara && mismatch == 0 && match == nPara;
                Ok(run, "fs-p-split", ok,
                    "切开=" + pieces.Count + " Count=" + nPara
                    + "  对上=" + match + "  对不上=" + mismatch);
            }
            catch (Exception ex)
            {
                run.CaseFail("fs-p-split", ex.GetType().Name + ": " + ex.Message);
            }
            finally
            {
                try
                {
                    app.ScreenUpdating = oldUpdate;
                }
                catch
                {
                }

                if (weOpened && doc != null)
                {
                    CloseDoc(doc);
                }
            }
        }

        private static List<string> SplitConvertedByParaEnds(string text)
        {
            var pieces = new List<string>();
            if (string.IsNullOrEmpty(text))
            {
                return pieces;
            }

            var sb = new StringBuilder();
            for (int i = 0; i < text.Length; i++)
            {
                if (text[i] == '\r')
                {
                    sb.Append('\r');
                    if (i + 1 < text.Length && text[i + 1] == '\a')
                    {
                        sb.Append('\a');
                        i++;
                    }

                    pieces.Add(sb.ToString());
                    sb.Length = 0;
                    continue;
                }

                if (text[i] == '^'
                    && i + 1 < text.Length
                    && text[i + 1] == 'b')
                {
                    sb.Append('^').Append('b');
                    pieces.Add(sb.ToString());
                    sb.Length = 0;
                    i++;
                    continue;
                }

                sb.Append(text[i]);
            }

            if (sb.Length > 0)
            {
                pieces.Add(sb.ToString());
            }

            return pieces;
        }

        private static void RunPublicCodeNth(TestRun run)
        {
            try
            {
                string display = BuildParagraphDisplayContent.Build(
                    new[] { "" },
                    new List<IReadOnlyList<string>>
                    {
                        new List<string> { "P_00003", "P_00001", "P_00003" }
                    },
                    new Dictionary<string, string>
                    {
                        ["P_00003"] = "重复段。\r",
                        ["P_00001"] = "中间段。\r",
                    });

                bool wroteN = display.IndexOf("[P_00003_1 start]", StringComparison.Ordinal) >= 0
                    && display.IndexOf("[P_00003_2 start]", StringComparison.Ordinal) >= 0
                    && display.IndexOf("[P_00001 start]", StringComparison.Ordinal) >= 0
                    && display.IndexOf("[P_00003-start]", StringComparison.Ordinal) < 0;

                int first = DisplayParagraphIndexMap.ResolveParaIndex(display, "P_00003_1");
                int second = DisplayParagraphIndexMap.ResolveParaIndex(display, "P_00003_2");
                int mid = DisplayParagraphIndexMap.ResolveParaIndex(display, "P_00001");
                int bare = DisplayParagraphIndexMap.ResolveParaIndex(display, "P_00003");

                bool ok = wroteN && first == 1 && mid == 2 && second == 3 && bare == 0;
                Ok(run, "fs-p-nth", ok,
                    "写出=" + wroteN
                    + "  _1=" + first
                    + "  中=" + mid
                    + "  _2=" + second
                    + "  光杆=" + bare);
            }
            catch (Exception ex)
            {
                run.CaseFail("fs-p-nth", ex.GetType().Name + ": " + ex.Message);
            }
        }

        private static void RunConvertParaWalk(TestRun run, Word.Application app)
        {
            string path = ResolveAlignDocPath();
            if (!System.IO.File.Exists(path))
            {
                run.CaseSkip("fs-p-map", "文件不在: " + path);
                return;
            }

            Word.Document doc = null;
            bool weOpened = false;
            bool oldUpdate = true;
            try
            {
                string norm = System.IO.Path.GetFullPath(path);
                foreach (Word.Document d in app.Documents)
                {
                    try
                    {
                        if (string.Equals(
                            System.IO.Path.GetFullPath(d.FullName),
                            norm,
                            StringComparison.OrdinalIgnoreCase))
                        {
                            doc = d;
                            break;
                        }
                    }
                    catch
                    {
                    }
                }

                if (doc == null)
                {
                    object fileName = path;
                    object readOnly = true;
                    object missing = System.Reflection.Missing.Value;
                    doc = app.Documents.Open(
                        ref fileName,
                        ref missing,
                        ref readOnly);
                    weOpened = true;
                }

                oldUpdate = app.ScreenUpdating;
                app.ScreenUpdating = false;
                try
                {
                    app.DisplayAlerts = Word.WdAlertLevel.wdAlertsNone;
                }
                catch
                {
                }

                Console.WriteLine("  --- display 边转边映射 P_ → 段号 ---");
                Console.WriteLine("    path=" + path);
                Console.Out.Flush();

                var options = ProcessDocumentOptions.ForGetDocumentContent("fs-pwalk");
                options.ForceRefresh = true;
                var sw = System.Diagnostics.Stopwatch.StartNew();
                WordDocumentExtractor.ProcessingResult processed =
                    WordDocumentExtractor.ProcessDocument(doc, options);
                Console.WriteLine("    ProcessDocument " + sw.Elapsed.TotalSeconds.ToString("F2") + "s"
                    + "  method=" + (processed != null ? processed.ProcessingMethod : "?"));
                Console.Out.Flush();

                string display = LoadParagraphDisplay(processed);
                string marks = DisplayStoryConvert.ToMarks(display);
                List<string> pieces = DisplayStoryConvert.SplitByParaEnds(marks);
                int walkEnds = pieces.Count;
                List<DisplayParagraphIndexMap.Hit> hits = DisplayParagraphIndexMap.Build(display);
                int nPara = doc.Paragraphs.Count;
                Console.WriteLine("    转换后段号=" + walkEnds
                    + "  P_映射=" + hits.Count
                    + "  Paragraphs.Count=" + nPara
                    + (walkEnds == nPara ? "  段数相等" : "  段数对不上"));
                Console.Out.Flush();

                int unassigned = 0;
                int oob = 0;
                int mismatch = 0;
                int match = 0;
                sw.Restart();
                for (int i = 0; i < hits.Count; i++)
                {
                    DisplayParagraphIndexMap.Hit hit = hits[i];
                    if (hit.ParaNo <= 0)
                    {
                        unassigned++;
                        if (unassigned <= 3)
                        {
                            Console.WriteLine("    未赋段号 " + hit.Code
                                + "  inner=" + QuoteSpecial(Peek(hit.Inner, 0, 24)));
                        }

                        continue;
                    }

                    if (hit.ParaNo > nPara)
                    {
                        oob++;
                        continue;
                    }

                    string com = "";
                    try
                    {
                        com = doc.Paragraphs[hit.ParaNo].Range.Text ?? "";
                    }
                    catch (Exception ex)
                    {
                        oob++;
                        if (oob <= 3)
                        {
                            Console.WriteLine("    取段失败 " + hit.Code
                                + " → " + hit.ParaNo + "  " + ex.Message);
                        }

                        continue;
                    }

                    com = NormParaForMap(com.Replace("\x01", ""));
                    string expect = NormParaForMap(hit.Inner.Replace("^b", "\f"));
                    if (string.Equals(com, expect, StringComparison.Ordinal))
                    {
                        match++;
                    }
                    else
                    {
                        mismatch++;
                        if (mismatch <= 8)
                        {
                            Console.WriteLine("    对不上 " + hit.Code
                                + " → 段[" + hit.ParaNo + "]");
                            Console.WriteLine("      P_   " + QuoteSpecial(Peek(expect, 0, 36)));
                            Console.WriteLine("      COM  " + QuoteSpecial(Peek(com, 0, 36)));
                        }
                    }
                }

                Console.WriteLine("    核对  用时=" + sw.Elapsed.TotalSeconds.ToString("F2") + "s"
                    + "  对上=" + match
                    + "  对不上=" + mismatch
                    + "  未赋=" + unassigned
                    + "  越界=" + oob);
                bool ok = walkEnds == nPara && mismatch == 0 && unassigned == 0 && oob == 0
                    && hits.Count > 0;
                Ok(run, "fs-p-map", ok,
                    "段号=" + walkEnds + "/" + nPara
                    + "  映射=" + hits.Count
                    + "  对上=" + match
                    + "  对不上=" + mismatch
                    + "  未赋=" + unassigned);
            }
            catch (Exception ex)
            {
                run.CaseFail("fs-p-map", ex.GetType().Name + ": " + ex.Message);
            }
            finally
            {
                try
                {
                    app.ScreenUpdating = oldUpdate;
                }
                catch
                {
                }

                if (weOpened && doc != null)
                {
                    CloseDoc(doc);
                }
            }
        }

        private static string LoadParagraphDisplay(WordDocumentExtractor.ProcessingResult processed)
        {
            var paras = DocumentState.ChunkParagraphDisplayContents;
            if (paras != null)
            {
                var sb = new StringBuilder();
                foreach (string p in paras)
                {
                    sb.Append(p ?? "");
                }

                if (sb.Length > 0)
                {
                    return sb.ToString();
                }
            }

            if (processed == null || processed.ChunkInfo == null)
            {
                return "";
            }

            var alt = new StringBuilder();
            foreach (WordDocumentExtractor.ChunkInfo chunk in processed.ChunkInfo)
            {
                if (!string.IsNullOrEmpty(chunk.ParagraphDisplayContent))
                {
                    alt.Append(chunk.ParagraphDisplayContent);
                }
                else
                {
                    alt.Append(chunk.DisplayContent ?? "");
                }
            }

            return alt.ToString();
        }

        private static string StripDisplayKeepP(string display)
        {
            if (string.IsNullOrEmpty(display))
            {
                return "";
            }

            string s = DisplayInTable.Replace(display, "");
            s = DisplayChart.Replace(s, "");
            s = DisplayImage.Replace(s, "");
            s = Regex.Replace(s, @"\[S_[^\]]+-(?:start|end)\]", "");
            s = Regex.Replace(s, @"\[T_[0-9A-Za-z]+_r[0-9]+_c[0-9]+-(?:start|end)\]", "");
            return s;
        }

        private static string NormParaForMap(string text)
        {
            if (string.IsNullOrEmpty(text))
            {
                return "";
            }

            text = FormFeedText.UnescapeToChar12(text);
            int end = text.Length;
            while (end > 0)
            {
                char c = text[end - 1];
                if (c != '\r' && c != '\a')
                {
                    break;
                }

                end--;
            }

            return end == text.Length ? text : text.Substring(0, end);
        }

        private sealed class PWalkHit
        {
            public string Code;
            public int ParaNo;
            public string Inner;
        }

        private static List<PWalkHit> WalkBuildPtoPara(string text, out int paraNo)
        {
            var hits = new List<PWalkHit>();
            paraNo = 0;
            if (string.IsNullOrEmpty(text))
            {
                return hits;
            }

            string openCode = null;
            var inner = new StringBuilder();
            bool assigned = false;
            int assignedPara = 0;
            var pending = new List<int>();
            int i = 0;
            while (i < text.Length)
            {
                if (text[i] == '[')
                {
                    int close = text.IndexOf(']', i);
                    if (close < 0)
                    {
                        break;
                    }

                    string tag = text.Substring(i + 1, close - i - 1);
                    if (tag.StartsWith("P_", StringComparison.Ordinal)
                        && (tag.EndsWith("-start", StringComparison.Ordinal)
                            || tag.EndsWith(" start", StringComparison.Ordinal)))
                    {
                        openCode = tag.Substring(0, tag.Length - 6);
                        inner.Length = 0;
                        assigned = false;
                        assignedPara = 0;
                    }
                    else if (tag.StartsWith("P_", StringComparison.Ordinal)
                        && (tag.EndsWith("-end", StringComparison.Ordinal)
                            || tag.EndsWith(" end", StringComparison.Ordinal)))
                    {
                        string code = tag.Substring(0, tag.Length - 4);
                        hits.Add(new PWalkHit
                        {
                            Code = code,
                            ParaNo = assigned ? assignedPara : 0,
                            Inner = inner.ToString()
                        });
                        if (!assigned)
                        {
                            pending.Add(hits.Count - 1);
                        }

                        openCode = null;
                    }

                    i = close + 1;
                    continue;
                }

                if (text[i] == '\r')
                {
                    paraNo++;
                    if (openCode != null)
                    {
                        assigned = true;
                        assignedPara = paraNo;
                    }

                    for (int p = 0; p < pending.Count; p++)
                    {
                        hits[pending[p]].ParaNo = paraNo;
                    }

                    pending.Clear();
                    if (openCode != null)
                    {
                        inner.Append('\r');
                    }

                    i++;
                    continue;
                }

                if (text[i] == '\f')
                {
                    if (openCode != null)
                    {
                        inner.Append('\f');
                    }

                    if (i + 1 >= text.Length || text[i + 1] != '\r')
                    {
                        paraNo++;
                        if (openCode != null)
                        {
                            assigned = true;
                            assignedPara = paraNo;
                        }

                        for (int p = 0; p < pending.Count; p++)
                        {
                            hits[pending[p]].ParaNo = paraNo;
                        }

                        pending.Clear();
                    }

                    i++;
                    continue;
                }

                if (text[i] == '^'
                    && i + 1 < text.Length
                    && text[i + 1] == 'b')
                {
                    paraNo++;
                    if (openCode != null)
                    {
                        assigned = true;
                        assignedPara = paraNo;
                    }

                    for (int p = 0; p < pending.Count; p++)
                    {
                        hits[pending[p]].ParaNo = paraNo;
                    }

                    pending.Clear();
                    if (openCode != null)
                    {
                        inner.Append('^').Append('b');
                    }

                    i += 2;
                    continue;
                }

                if (openCode != null)
                {
                    inner.Append(text[i]);
                }

                i++;
            }

            return hits;
        }

        private static string ResolveAlignDocPath()
        {
            if (!string.IsNullOrEmpty(AlignDocPath))
            {
                return AlignDocPath;
            }

            return System.IO.File.Exists(RealDocxPath) ? RealDocxPath : RealDocPath;
        }

        private static void RunRealDocParaChars(TestRun run, Word.Application app)
        {
            string path = ResolveAlignDocPath();
            if (!System.IO.File.Exists(path))
            {
                run.CaseSkip("fs-para-chars-time", "文件不在: " + path);
                return;
            }

            Word.Document doc = null;
            bool weOpened = false;
            bool oldUpdate = true;
            try
            {
                string norm = System.IO.Path.GetFullPath(path);
                foreach (Word.Document d in app.Documents)
                {
                    try
                    {
                        if (string.Equals(
                            System.IO.Path.GetFullPath(d.FullName),
                            norm,
                            StringComparison.OrdinalIgnoreCase))
                        {
                            doc = d;
                            break;
                        }
                    }
                    catch
                    {
                    }
                }

                if (doc == null)
                {
                    object fileName = path;
                    object readOnly = true;
                    object missing = System.Reflection.Missing.Value;
                    doc = app.Documents.Open(
                        ref fileName,
                        ref missing,
                        ref readOnly);
                    weOpened = true;
                }

                oldUpdate = app.ScreenUpdating;
                app.ScreenUpdating = false;
                try
                {
                    app.DisplayAlerts = Word.WdAlertLevel.wdAlertsNone;
                }
                catch
                {
                }

                Console.WriteLine("  --- 按段扫 Characters.Text ---");
                Console.WriteLine("    path=" + path);
                Console.Out.Flush();

                var sw = System.Diagnostics.Stopwatch.StartNew();
                int nPara = doc.Paragraphs.Count;
                Console.WriteLine("    Paragraphs.Count=" + nPara
                    + "  取Count=" + sw.Elapsed.TotalSeconds.ToString("F2") + "s");
                Console.Out.Flush();

                sw.Restart();
                int nChar = 0;
                int i = 0;
                foreach (Word.Paragraph para in doc.Paragraphs)
                {
                    Word.Characters chars = para.Range.Characters;
                    int n = chars.Count;
                    for (int c = 1; c <= n; c++)
                    {
                        string t = chars[c].Text ?? "";
                        nChar += t.Length;
                    }

                    i++;
                    if (i % 20 == 0 || i == nPara)
                    {
                        Console.WriteLine("    段 " + i + "/" + nPara
                            + "  字 " + nChar
                            + "  " + sw.Elapsed.TotalSeconds.ToString("F1") + "s");
                        Console.Out.Flush();
                    }
                }

                sw.Stop();
                Console.WriteLine("    扫完 段=" + i + "  字=" + nChar
                    + "  用时=" + sw.Elapsed.TotalSeconds.ToString("F1") + "s");
                Ok(run, "fs-para-chars-time", i == nPara,
                    "段=" + i + " 字=" + nChar
                    + " " + sw.Elapsed.TotalSeconds.ToString("F1") + "s");
            }
            catch (Exception ex)
            {
                run.CaseFail("fs-para-chars-time", ex.GetType().Name + ": " + ex.Message);
            }
            finally
            {
                try
                {
                    app.ScreenUpdating = oldUpdate;
                }
                catch
                {
                }

                if (weOpened && doc != null)
                {
                    CloseDoc(doc);
                }
            }
        }

        private static void RunRealDocParaTexts(TestRun run, Word.Application app)
        {
            if (!System.IO.File.Exists(RealDocPath))
            {
                run.CaseSkip("fs-para-text-time", "文件不在: " + RealDocPath);
                return;
            }

            Word.Document doc = null;
            bool weOpened = false;
            bool oldUpdate = true;
            try
            {
                string norm = System.IO.Path.GetFullPath(RealDocPath);
                foreach (Word.Document d in app.Documents)
                {
                    try
                    {
                        if (string.Equals(
                            System.IO.Path.GetFullPath(d.FullName),
                            norm,
                            StringComparison.OrdinalIgnoreCase))
                        {
                            doc = d;
                            break;
                        }
                    }
                    catch
                    {
                    }
                }

                if (doc == null)
                {
                    object fileName = RealDocPath;
                    object readOnly = true;
                    object missing = System.Reflection.Missing.Value;
                    doc = app.Documents.Open(
                        ref fileName,
                        ref missing,
                        ref readOnly);
                    weOpened = true;
                }

                oldUpdate = app.ScreenUpdating;
                app.ScreenUpdating = false;
                try
                {
                    app.DisplayAlerts = Word.WdAlertLevel.wdAlertsNone;
                }
                catch
                {
                }

                Console.WriteLine("  --- 真稿只扫 Paragraph.Range.Text ---");
                Console.WriteLine("    path=" + RealDocPath);

                var sw = System.Diagnostics.Stopwatch.StartNew();
                int n = doc.Paragraphs.Count;
                Console.WriteLine("    Paragraphs.Count=" + n
                    + "  取Count用时=" + sw.Elapsed.TotalSeconds.ToString("F2") + "s");
                Console.Out.Flush();

                sw.Restart();
                int chars = 0;
                int i = 0;
                foreach (Word.Paragraph para in doc.Paragraphs)
                {
                    string t = para.Range.Text ?? "";
                    chars += t.Length;
                    i++;
                    if (i % 100 == 0)
                    {
                        Console.WriteLine("    段 " + i + "/" + n
                            + "  字 " + chars
                            + "  " + sw.Elapsed.TotalSeconds.ToString("F2") + "s");
                        Console.Out.Flush();
                    }
                }

                sw.Stop();
                Console.WriteLine("    扫完 " + i + " 段  Text合计=" + chars
                    + "  用时=" + sw.Elapsed.TotalSeconds.ToString("F2") + "s");
                Ok(run, "fs-para-text-time", i == n,
                    i + " 段 " + chars + " 字 " + sw.Elapsed.TotalSeconds.ToString("F2") + "s");
            }
            catch (Exception ex)
            {
                run.CaseFail("fs-para-text-time", ex.Message);
            }
            finally
            {
                try
                {
                    app.ScreenUpdating = oldUpdate;
                }
                catch
                {
                }

                if (weOpened && doc != null)
                {
                    CloseDoc(doc);
                }
            }
        }

        private static void RunRealDocCharsAbut(TestRun run, Word.Application app)
        {
            if (!System.IO.File.Exists(RealDocPath))
            {
                run.CaseSkip("fs-real-chars-abut", "文件不在: " + RealDocPath);
                run.CaseSkip("fs-text-align", "文件不在: " + RealDocPath);
                return;
            }

            Word.Document doc = null;
            bool weOpened = false;
            bool oldUpdate = true;
            try
            {
                string norm = System.IO.Path.GetFullPath(RealDocPath);
                foreach (Word.Document d in app.Documents)
                {
                    try
                    {
                        if (string.Equals(
                            System.IO.Path.GetFullPath(d.FullName),
                            norm,
                            StringComparison.OrdinalIgnoreCase))
                        {
                            doc = d;
                            break;
                        }
                    }
                    catch
                    {
                    }
                }

                if (doc == null)
                {
                    object fileName = RealDocPath;
                    object readOnly = true;
                    object missing = System.Reflection.Missing.Value;
                    doc = app.Documents.Open(
                        ref fileName,
                        ref missing,
                        ref readOnly);
                    weOpened = true;
                }

                oldUpdate = app.ScreenUpdating;
                app.ScreenUpdating = false;
                try
                {
                    app.DisplayAlerts = Word.WdAlertLevel.wdAlertsNone;
                }
                catch
                {
                }

                Console.WriteLine("  --- 真稿按段扫 Characters 收尾 ---");
                Console.WriteLine("    path=" + RealDocPath);

                var sw = System.Diagnostics.Stopwatch.StartNew();
                int contentStart = doc.Content.Start;
                int contentEnd = doc.Content.End;
                int expect = doc.Content.Characters.Count;
                int paraCount = doc.Paragraphs.Count;
                Console.WriteLine("    Content=" + contentStart + "-" + contentEnd
                    + "  Characters.Count=" + expect
                    + "  Paragraphs=" + paraCount
                    + "  用时=" + sw.Elapsed.TotalSeconds.ToString("F2") + "s");
                Console.Out.Flush();

                sw.Restart();
                string contentText = doc.Content.Text ?? "";
                var items = new List<CharItem>(expect);
                int gaps = 0;
                int overs = 0;
                int prevEnd = int.MinValue;
                int firstStart = -1;
                int paraIdx = 0;
                var samples = new StringBuilder();
                foreach (Word.Paragraph para in doc.Paragraphs)
                {
                    paraIdx++;
                    Word.Characters chars = para.Range.Characters;
                    int n = chars.Count;
                    for (int i = 1; i <= n; i++)
                    {
                        Word.Range ch = chars[i];
                        var item = new CharItem
                        {
                            Start = ch.Start,
                            End = ch.End,
                            Text = ch.Text ?? ""
                        };
                        items.Add(item);
                        if (firstStart < 0)
                        {
                            firstStart = item.Start;
                            Console.WriteLine("    首字 " + item.Start + "-" + item.End
                                + "  " + QuoteSpecial(item.Text));
                            Console.Out.Flush();
                        }

                        if (prevEnd != int.MinValue && prevEnd != item.Start)
                        {
                            string kind;
                            if (prevEnd < item.Start)
                            {
                                gaps++;
                                kind = "GAP";
                            }
                            else
                            {
                                overs++;
                                kind = "OVER";
                            }

                            if (gaps + overs <= 12)
                            {
                                samples.Append("    ").Append(kind)
                                    .Append(" 段").Append(paraIdx)
                                    .Append("[").Append(i).Append("] ")
                                    .Append(prevEnd).Append(" → ")
                                    .Append(item.Start).Append("-").Append(item.End)
                                    .Append("  ").Append(QuoteSpecial(item.Text))
                                    .AppendLine();
                            }
                        }

                        prevEnd = item.End;
                    }

                    if (paraIdx % 30 == 0)
                    {
                        Console.WriteLine("    段 " + paraIdx + "/" + paraCount
                            + "  字 " + items.Count
                            + "  " + sw.Elapsed.TotalSeconds.ToString("F1") + "s");
                        Console.Out.Flush();
                    }
                }

                sw.Stop();
                int done = items.Count;
                bool cover = firstStart == contentStart && prevEnd == contentEnd;
                bool countMatch = done == expect;
                Console.WriteLine("    扫了=" + done + "  期望Count=" + expect
                    + "  首Start=" + firstStart
                    + "  末End=" + prevEnd
                    + "  GAP=" + gaps
                    + "  OVER=" + overs
                    + "  用时=" + sw.Elapsed.TotalSeconds.ToString("F1") + "s");
                if (samples.Length > 0)
                {
                    Console.Write(samples.ToString());
                }

                Ok(run, "fs-real-chars-abut",
                    gaps == 0 && overs == 0 && cover && countMatch,
                    "GAP=" + gaps + " OVER=" + overs
                    + (cover ? "" : " 未盖满 Content")
                    + (countMatch ? "" : " 条数≠Count"));

                AlignContentText(run, contentText, items, contentStart, contentEnd);
            }
            catch (Exception ex)
            {
                run.CaseFail("fs-real-chars-abut", ex.Message);
                run.CaseFail("fs-text-align", ex.Message);
            }
            finally
            {
                try
                {
                    app.ScreenUpdating = oldUpdate;
                }
                catch
                {
                }

                if (weOpened && doc != null)
                {
                    CloseDoc(doc);
                }
            }
        }

        private struct CharItem
        {
            public int Start;
            public int End;
            public string Text;
        }

        private struct AlignRule
        {
            public string Pattern;
            public int Span;
        }

        private static void AlignContentText(
            TestRun run,
            string text,
            List<CharItem> items,
            int contentStart,
            int contentEnd)
        {
            Console.WriteLine("  --- Content.Text 与 Characters 对齐 ---");
            Console.WriteLine("    Text.Length=" + text.Length
                + "  items=" + items.Count
                + "  Content=" + contentStart + "-" + contentEnd);

            var concat = new StringBuilder(text.Length);
            foreach (CharItem item in items)
            {
                concat.Append(item.Text);
            }

            bool concatEq = string.Equals(concat.ToString(), text, StringComparison.Ordinal);
            Console.WriteLine("    concat(Characters.Text)==Content.Text ? " + concatEq);
            if (!concatEq)
            {
                int miss = FirstDiff(concat.ToString(), text);
                Console.WriteLine("    concat 首差下标=" + miss);
            }

            var rules = new List<AlignRule>();
            string fail = null;
            int rounds = 0;
            const int maxRounds = 40;
            while (rounds < maxRounds)
            {
                rounds++;
                AlignMismatch miss = FindAlignMismatch(text, items, rules, contentStart);
                if (miss == null)
                {
                    Console.WriteLine("    第" + rounds + "轮对齐成功  规则数=" + rules.Count);
                    break;
                }

                Console.WriteLine("    第" + rounds + "轮对不上  Characters[" + miss.ItemIndex + "]"
                    + " 实际=" + miss.Item.Start + "-" + miss.Item.End
                    + " Text=" + QuoteSpecial(miss.Item.Text)
                    + "  span=" + (miss.Item.End - miss.Item.Start)
                    + "  文本下标=" + miss.TextIndex
                    + "  预测Start=" + miss.PredictedStart
                    + "  预测消耗=" + miss.PredictedConsume + "/" + miss.PredictedSpan
                    + "  原文=" + QuoteSpecial(Peek(text, miss.TextIndex, 8)));

                if (miss.TextIndex < 0
                    || miss.TextIndex > text.Length
                    || !StartsWithAt(text, miss.TextIndex, miss.Item.Text))
                {
                    fail = "Characters.Text 对不上 Content.Text 下标 " + miss.TextIndex;
                    break;
                }

                string pattern = miss.Item.Text;
                int span = miss.Item.End - miss.Item.Start;
                AlignRule existing;
                int existAt = IndexOfRule(rules, pattern);
                if (existAt >= 0)
                {
                    existing = rules[existAt];
                    if (existing.Span != span)
                    {
                        fail = "同文不同跨度 " + QuoteSpecial(pattern)
                            + " span=" + existing.Span + "/" + span;
                        break;
                    }

                    fail = "规则已在仍对不上 " + QuoteSpecial(pattern);
                    break;
                }

                rules.Add(new AlignRule { Pattern = pattern, Span = span });
                Console.WriteLine("    加规则 " + QuoteSpecial(pattern) + " → span=" + span
                    + "  重头再对");
            }

            if (fail == null && rounds >= maxRounds)
            {
                fail = "超过 " + maxRounds + " 轮";
            }

            Console.WriteLine("    最终规则:");
            if (rules.Count == 0)
            {
                Console.WriteLine("      （无，默认 1 字 = 1 格）");
            }
            else
            {
                var hist = new Dictionary<string, int>(StringComparer.Ordinal);
                foreach (AlignRule rule in rules)
                {
                    string key = QuoteSpecial(rule.Pattern) + " → span=" + rule.Span;
                    int n;
                    hist.TryGetValue(key, out n);
                    hist[key] = n + 1;
                }

                foreach (KeyValuePair<string, int> kv in hist)
                {
                    Console.WriteLine("      " + kv.Key);
                }
            }

            Ok(run, "fs-text-align", fail == null,
                fail == null ? ("规则 " + rules.Count + " 条，" + rounds + " 轮对齐") : fail);
        }

        private sealed class AlignMismatch
        {
            public int ItemIndex;
            public CharItem Item;
            public int TextIndex;
            public int PredictedStart;
            public int PredictedConsume;
            public int PredictedSpan;
        }

        private static AlignMismatch FindAlignMismatch(
            string text,
            List<CharItem> items,
            List<AlignRule> rules,
            int predicted)
        {
            int ti = 0;
            for (int ci = 0; ci < items.Count; ci++)
            {
                CharItem item = items[ci];
                int consume;
                int span;
                PredictOne(text, ti, rules, out consume, out span);
                bool textOk = StartsWithAt(text, ti, item.Text);
                int actualSpan = item.End - item.Start;
                if (!textOk || consume != item.Text.Length || span != actualSpan || predicted != item.Start)
                {
                    return new AlignMismatch
                    {
                        ItemIndex = ci + 1,
                        Item = item,
                        TextIndex = ti,
                        PredictedStart = predicted,
                        PredictedConsume = consume,
                        PredictedSpan = span
                    };
                }

                ti += item.Text.Length;
                predicted += actualSpan;
            }

            if (ti != text.Length)
            {
                return new AlignMismatch
                {
                    ItemIndex = items.Count,
                    Item = items.Count > 0 ? items[items.Count - 1] : default(CharItem),
                    TextIndex = ti,
                    PredictedStart = predicted,
                    PredictedConsume = 0,
                    PredictedSpan = 0
                };
            }

            return null;
        }

        private static void PredictOne(
            string text,
            int ti,
            List<AlignRule> rules,
            out int consume,
            out int span)
        {
            int bestLen = -1;
            int bestSpan = 1;
            foreach (AlignRule rule in rules)
            {
                if (rule.Pattern.Length > bestLen && StartsWithAt(text, ti, rule.Pattern))
                {
                    bestLen = rule.Pattern.Length;
                    bestSpan = rule.Span;
                }
            }

            if (bestLen > 0)
            {
                consume = bestLen;
                span = bestSpan;
                return;
            }

            consume = ti < text.Length ? 1 : 0;
            span = consume;
        }

        private static int IndexOfRule(List<AlignRule> rules, string pattern)
        {
            for (int i = 0; i < rules.Count; i++)
            {
                if (string.Equals(rules[i].Pattern, pattern, StringComparison.Ordinal))
                {
                    return i;
                }
            }

            return -1;
        }

        private static bool StartsWithAt(string text, int index, string value)
        {
            if (value == null || index < 0 || index + value.Length > text.Length)
            {
                return false;
            }

            return string.CompareOrdinal(text, index, value, 0, value.Length) == 0;
        }

        private static int FirstDiff(string a, string b)
        {
            int n = Math.Min(a.Length, b.Length);
            for (int i = 0; i < n; i++)
            {
                if (a[i] != b[i])
                {
                    return i;
                }
            }

            return a.Length == b.Length ? -1 : n;
        }

        private static string Peek(string text, int index, int take)
        {
            if (index < 0 || index >= text.Length)
            {
                return "";
            }

            int n = Math.Min(take, text.Length - index);
            return text.Substring(index, n);
        }

        private static string CheckAbut(string title, Word.Range scope)
        {
            Word.Characters chars = scope.Characters;
            int n = chars.Count;
            var sb = new StringBuilder();
            sb.Append("    ").Append(title)
                .Append(" Range=").Append(scope.Start).Append("-").Append(scope.End)
                .Append(" Count=").Append(n);
            if (n > 0)
            {
                sb.Append("  firstStart=").Append(chars[1].Start)
                    .Append(" lastEnd=").Append(chars[n].End);
            }

            sb.AppendLine();
            int gaps = 0;
            int overs = 0;
            for (int i = 1; i < n; i++)
            {
                if (i == 1 || i % 2000 == 0)
                {
                    Console.WriteLine("    扫到 " + i + "/" + n);
                }

                Word.Range a = chars[i];
                Word.Range b = chars[i + 1];
                if (a.End == b.Start)
                {
                    continue;
                }

                string kind = a.End < b.Start ? "GAP" : "OVER";
                if (a.End < b.Start)
                {
                    gaps++;
                }
                else
                {
                    overs++;
                }

                if (gaps + overs <= 8)
                {
                    sb.Append("    ").Append(kind)
                        .Append(" [").Append(i).Append("] ").Append(a.Start).Append("-").Append(a.End)
                        .Append(" ").Append(QuoteSpecial(a.Text))
                        .Append(" → [").Append(i + 1).Append("] ").Append(b.Start).Append("-").Append(b.End)
                        .Append(" ").Append(QuoteSpecial(b.Text))
                        .AppendLine();
                }
            }

            sb.Append("    不相连: GAP=").Append(gaps).Append(" OVER=").Append(overs);
            return sb.ToString();
        }

        private static string QuoteSpecial(string text)
        {
            if (text == null)
            {
                return "(null)";
            }

            var sb = new StringBuilder("\"");
            foreach (char ch in text)
            {
                if (ch == '\r')
                {
                    sb.Append("\\r");
                }
                else if (ch == '\n')
                {
                    sb.Append("\\n");
                }
                else if (ch == '\a')
                {
                    sb.Append("\\a");
                }
                else if (ch == '\t')
                {
                    sb.Append("\\t");
                }
                else if (ch == '\f')
                {
                    sb.Append("\\f");
                }
                else if (ch < 32)
                {
                    sb.Append("\\x").Append(((int)ch).ToString("X2"));
                }
                else
                {
                    sb.Append(ch);
                }
            }

            sb.Append("\"");
            return sb.ToString();
        }

        private static bool PlantTableFixture(Word.Document doc)
        {
            doc.Content.Text = Needle + "\r间隔乙。\r" + Needle + "\r";
            int end = doc.Content.End;
            Word.Range at = doc.Range(end - 1, end - 1);
            Word.Table table = doc.Tables.Add(at, 2, 2);
            table.Cell(1, 1).Range.Text = Needle;
            table.Cell(1, 2).Range.Text = "表内丙。";
            table.Cell(2, 1).Range.Text = Needle;
            table.Cell(2, 2).Range.Text = "表内丁。";
            Word.Range after = doc.Range(doc.Content.End - 1, doc.Content.End - 1);
            after.InsertAfter(Needle);
            return doc.Tables.Count >= 1;
        }

        private static void ProbeSecondCellScope(TestRun run, Word.Document doc)
        {
            Word.Range cell = doc.Tables[1].Cell(2, 1).Range;
            int scopeStart = cell.Start;
            int scopeEnd = cell.End;
            Word.Range search = doc.Range(scopeStart, scopeEnd);
            Word.Find find = search.Find;
            find.ClearFormatting();
            find.Text = Needle;
            find.MatchCase = true;
            find.MatchWholeWord = false;
            find.Wrap = Word.WdFindWrap.wdFindStop;
            bool ok = find.Execute();
            int hitStart = ok ? search.Start : -1;
            int hitEnd = ok ? search.End : -1;
            bool bounced = ok && hitStart < scopeStart;
            bool inside = ok && hitStart >= scopeStart && hitEnd <= scopeEnd;
            Console.WriteLine("  --- 限定 [2,1] " + scopeStart + "-" + scopeEnd + " 搜「" + Needle + "」 ---");
            Console.WriteLine("    Execute=" + ok
                + "  hit=" + hitStart + "-" + hitEnd
                + (bounced ? "  弹回格外" : (inside ? "  落在本格" : "  未落本格")));
            Ok(run, "fs-cell21-scope", ok && inside && !bounced,
                ok
                    ? ("hit=" + hitStart + "-" + hitEnd
                        + " scope=" + scopeStart + "-" + scopeEnd
                        + (bounced ? " 弹回第一格" : " 未落在 26–31"))
                    : "本格 Find 未命中");
        }

        private static void ProbeTableTailScope(TestRun run, Word.Document doc)
        {
            Word.Range table = doc.Tables[1].Range;
            int firstCellEnd = doc.Tables[1].Cell(1, 1).Range.End;
            int scopeStart = firstCellEnd;
            int scopeEnd = table.End;
            Word.Range search = doc.Range(scopeStart, scopeEnd);
            Word.Find find = search.Find;
            find.ClearFormatting();
            find.Text = Needle;
            find.MatchCase = true;
            find.MatchWholeWord = false;
            find.Wrap = Word.WdFindWrap.wdFindStop;
            bool ok = find.Execute();
            int hitStart = ok ? search.Start : -1;
            int hitEnd = ok ? search.End : -1;
            bool bounced = ok && hitStart < scopeStart;
            bool inside = ok && hitStart >= scopeStart && hitEnd <= scopeEnd;
            Console.WriteLine("  --- 限定表尾 " + scopeStart + "-" + scopeEnd
                + "（不含 [1,1]）搜「" + Needle + "」 ---");
            Console.WriteLine("    Execute=" + ok
                + "  hit=" + hitStart + "-" + hitEnd
                + (bounced ? "  弹回第一格" : (inside ? "  落在窗内（第二处）" : "  未落窗内")));
            Ok(run, "fs-table-20-37", ok && inside && !bounced,
                ok
                    ? ("hit=" + hitStart + "-" + hitEnd
                        + " scope=" + scopeStart + "-" + scopeEnd
                        + (bounced ? " 弹回 15–20" : " 未落在第二处"))
                    : "20–37 内 Find 未命中");
        }

        private static void ProbeParagraphList(TestRun run, Word.Document doc)
        {
            Console.WriteLine("  --- 全部 Paragraph.Range ---");
            var hits = new System.Collections.Generic.List<Word.Range>();
            var tableHits = new System.Collections.Generic.List<Word.Range>();
            int i = 0;
            foreach (Word.Paragraph para in doc.Paragraphs)
            {
                i++;
                Word.Range pr = para.Range;
                string raw = pr.Text ?? "";
                string shown = raw.Replace("\r", "\\r").Replace("\a", "\\a").Replace("\n", "\\n");
                if (shown.Length > 24)
                {
                    shown = shown.Substring(0, 24) + "…";
                }

                bool inTable = false;
                try
                {
                    inTable = pr.Tables.Count > 0;
                }
                catch
                {
                }

                bool exact = string.Equals(NormParaText(raw), Needle, StringComparison.Ordinal);
                Console.WriteLine("    P" + i + " " + pr.Start + "-" + pr.End
                    + (inTable ? " table" : " body")
                    + (exact ? " EXACT" : "")
                    + "  \"" + shown + "\"");
                if (exact)
                {
                    hits.Add(pr);
                    if (inTable)
                    {
                        tableHits.Add(pr);
                    }
                }
            }

            bool secondOk = FindInsideParagraph(hits.Count >= 2 ? hits[1] : null, "全文第2精确段");
            bool tableSecondOk = FindInsideParagraph(
                tableHits.Count >= 2 ? tableHits[1] : null,
                "表内第2精确段");
            Ok(run, "fs-para-2nd-find", secondOk && tableSecondOk,
                "全文第2=" + secondOk + " 表内第2=" + tableSecondOk
                + " exact=" + hits.Count + " tableExact=" + tableHits.Count);
        }

        private static bool FindInsideParagraph(Word.Range para, string label)
        {
            if (para == null)
            {
                Console.WriteLine("    " + label + " 没有");
                return false;
            }

            Word.Range search = para.Document.Range(para.Start, para.End);
            Word.Find find = search.Find;
            find.ClearFormatting();
            find.Text = Needle;
            find.MatchCase = true;
            find.MatchWholeWord = false;
            find.Wrap = Word.WdFindWrap.wdFindStop;
            bool ok = find.Execute();
            int hitStart = ok ? search.Start : -1;
            int hitEnd = ok ? search.End : -1;
            bool inside = ok && hitStart >= para.Start && hitEnd <= para.End;
            bool bounced = ok && hitStart < para.Start;
            Console.WriteLine("    " + label + " " + para.Start + "-" + para.End
                + " 内 Find → Execute=" + ok + " hit=" + hitStart + "-" + hitEnd
                + (bounced ? " 弹回段外" : (inside ? " 落在本段" : " 未落本段")));
            return ok && inside && !bounced;
        }

        private static string NormParaText(string text)
        {
            if (string.IsNullOrEmpty(text))
            {
                return "";
            }

            return text.TrimEnd('\r', '\a', '\n');
        }

        private static StaleWalk Walk(Word.Document doc, string needle)
        {
            var walk = new StaleWalk();
            int contentEnd = doc.Content.End;
            walk.ContentEnd = contentEnd;
            if (doc.Tables.Count > 0)
            {
                Word.Table table = doc.Tables[1];
                Word.Range tr = table.Range;
                walk.TableStart = tr.Start;
                walk.TableEnd = tr.End;
                walk.CellLog = DumpCellRanges(table);
            }

            int cursor = 0;
            int guard = 0;
            while (cursor < contentEnd && guard < 40)
            {
                guard++;
                if (!TryFind(doc, cursor, contentEnd, needle, out int hitStart, out int hitEnd))
                {
                    break;
                }

                bool stale = hitStart < cursor;
                string zone = Zone(cursor, walk.TableStart, walk.TableEnd);
                walk.AddNext(cursor, hitStart, hitEnd, stale, zone);
                if (hitEnd <= cursor)
                {
                    walk.NextStuck = true;
                    break;
                }

                cursor = hitEnd;
            }

            bool oldUpdate = true;
            try
            {
                oldUpdate = doc.Application.ScreenUpdating;
                doc.Application.ScreenUpdating = false;
                for (int c = 0; c < contentEnd; c++)
                {
                    if (!TryFind(doc, c, contentEnd, needle, out int hitStart, out _))
                    {
                        continue;
                    }

                    if (hitStart >= c)
                    {
                        continue;
                    }

                    string zone = Zone(c, walk.TableStart, walk.TableEnd);
                    walk.AddCursorStale(c, hitStart, zone);
                }
            }
            finally
            {
                try
                {
                    doc.Application.ScreenUpdating = oldUpdate;
                }
                catch
                {
                }
            }

            return walk;
        }

        private static bool TryFind(
            Word.Document doc,
            int searchStart,
            int contentEnd,
            string needle,
            out int hitStart,
            out int hitEnd)
        {
            hitStart = -1;
            hitEnd = -1;
            Word.Range search = doc.Range(searchStart, contentEnd);
            Word.Find find = search.Find;
            find.ClearFormatting();
            find.Text = needle;
            find.MatchCase = true;
            find.MatchWholeWord = false;
            find.Wrap = Word.WdFindWrap.wdFindStop;
            if (!find.Execute())
            {
                return false;
            }

            hitStart = search.Start;
            hitEnd = search.End;
            return true;
        }

        private static string DumpCellRanges(Word.Table table)
        {
            var sb = new StringBuilder();
            int rows = table.Rows.Count;
            int cols = table.Columns.Count;
            sb.AppendLine("    单元格 Range（" + rows + "×" + cols + "）");
            for (int r = 1; r <= rows; r++)
            {
                for (int c = 1; c <= cols; c++)
                {
                    try
                    {
                        Word.Range cr = table.Cell(r, c).Range;
                        string text = (cr.Text ?? "").Replace("\r", "\\r").Replace("\a", "\\a");
                        if (text.Length > 20)
                        {
                            text = text.Substring(0, 20) + "…";
                        }

                        sb.Append("    [").Append(r).Append(",").Append(c).Append("] ")
                            .Append(cr.Start).Append("-").Append(cr.End)
                            .Append("  text=\"").Append(text).Append("\"")
                            .AppendLine();
                    }
                    catch (Exception ex)
                    {
                        sb.Append("    [").Append(r).Append(",").Append(c).Append("] 读不到：")
                            .Append(ex.Message)
                            .AppendLine();
                    }
                }
            }

            return sb.ToString();
        }

        private static string Zone(int cursor, int tableStart, int tableEnd)
        {
            if (tableStart < 0)
            {
                return "plain";
            }

            if (cursor < tableStart)
            {
                return "before";
            }

            if (cursor < tableEnd)
            {
                return "in";
            }

            return "after";
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

        private static void Ok(TestRun run, string name, bool pass, string reason)
        {
            if (pass)
            {
                run.CaseOk(name);
            }
            else
            {
                run.CaseFail(name, reason ?? "");
            }
        }

        private sealed class StaleWalk
        {
            public int ContentEnd;
            public int TableStart = -1;
            public int TableEnd = -1;
            public int NextHits;
            public bool NextStuck;
            public int StaleBefore;
            public int StaleIn;
            public int StaleAfter;
            public int StalePlain;
            public string CellLog = "";
            public readonly StringBuilder NextLog = new StringBuilder();
            public readonly StringBuilder StaleSamples = new StringBuilder();
            private int _sample;

            public int StaleTotal
            {
                get { return StaleBefore + StaleIn + StaleAfter + StalePlain; }
            }

            public void AddNext(int cursor, int hitStart, int hitEnd, bool stale, string zone)
            {
                NextHits++;
                NextLog.Append("    #").Append(NextHits)
                    .Append(" cursor=").Append(cursor)
                    .Append(" hit=").Append(hitStart).Append("-").Append(hitEnd)
                    .Append(" zone=").Append(zone)
                    .Append(stale ? " STALE" : " ok")
                    .AppendLine();
            }

            public void AddCursorStale(int cursor, int hitStart, string zone)
            {
                if (zone == "before")
                {
                    StaleBefore++;
                }
                else if (zone == "in")
                {
                    StaleIn++;
                }
                else if (zone == "after")
                {
                    StaleAfter++;
                }
                else
                {
                    StalePlain++;
                }

                if (_sample >= 8)
                {
                    return;
                }

                _sample++;
                StaleSamples.Append("    stale#").Append(_sample)
                    .Append(" cursor=").Append(cursor)
                    .Append(" hitStart=").Append(hitStart)
                    .Append(" zone=").Append(zone)
                    .AppendLine();
            }

            public string FormatLog(string title)
            {
                var sb = new StringBuilder();
                sb.AppendLine("  --- Find stale  " + title + " ---");
                sb.Append("    Content.End=").Append(ContentEnd);
                if (TableStart >= 0)
                {
                    sb.Append("  table=").Append(TableStart).Append("-").Append(TableEnd);
                }

                sb.AppendLine();
                if (!string.IsNullOrEmpty(CellLog))
                {
                    sb.Append(CellLog);
                }

                sb.AppendLine("    Find Next（cursor=hit.End）命中 " + NextHits
                    + (NextStuck ? "  中途卡住" : ""));
                sb.Append(NextLog);
                sb.AppendLine("    游标+1 stale 合计 " + StaleTotal
                    + "  before=" + StaleBefore
                    + "  in=" + StaleIn
                    + "  after=" + StaleAfter
                    + "  plain=" + StalePlain);
                sb.Append(StaleSamples);
                return sb.ToString();
            }
        }
    }
}
