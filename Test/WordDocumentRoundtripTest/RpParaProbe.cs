using System;
using System.Runtime.InteropServices;
using Word = Microsoft.Office.Interop.Word;

namespace WordDocumentRoundtripTest
{
    /// <summary>
    /// 看 Range.Paragraphs 是否只返回圈定范围内的段落。只新增临时稿，不 Quit 宿主。
    /// </summary>
    internal static class RpParaProbe
    {
        public static int Run(string host)
        {
            Word.Application app = AttachApp(host, out bool created);
            if (app == null)
            {
                Console.WriteLine("连不上 Word/WPS。");
                return 1;
            }

            bool screen = true;
            try
            {
                screen = app.ScreenUpdating;
            }
            catch
            {
            }

            try
            {
                try
                {
                    app.ScreenUpdating = false;
                }
                catch
                {
                }

                RunPlain(app);
                RunTable(app);
                RunNumbered(app);
                RunFind(app);
                RunTableFind(app);
                return 0;
            }
            finally
            {
                try
                {
                    app.ScreenUpdating = screen;
                }
                catch
                {
                }

                if (created)
                {
                    try
                    {
                        app.Quit(false);
                    }
                    catch
                    {
                    }
                }
            }
        }

        private static void RunPlain(Word.Application app)
        {
            Word.Document doc = app.Documents.Add();
            try
            {
                doc.Content.Text = "AAA前段。\rBBB目标一。\rCCC目标二。\rDDD后段。\r";
                Word.Paragraph p1 = doc.Paragraphs[1];
                Word.Paragraph p2 = doc.Paragraphs[2];
                Word.Paragraph p3 = doc.Paragraphs[3];
                DumpDoc(doc, "正文全文");
                Dump(doc, "正文：整段目标一到目标二", doc.Range(p2.Range.Start, p3.Range.End));
                Dump(doc, "正文：从目标一中间到目标二中间", doc.Range(p2.Range.Start + 2, p3.Range.End - 2));
                Dump(doc, "正文：只含目标一段落符", doc.Range(p2.Range.Start, p2.Range.End));
                Dump(doc, "正文：目标一不含段落符", doc.Range(p2.Range.Start, p2.Range.End - 1));
                Dump(doc, "正文：卡在目标一和目标二的段落符上", doc.Range(p2.Range.End - 1, p3.Range.Start + 1));
                Dump(doc, "正文：零宽点在目标一起点", doc.Range(p2.Range.Start, p2.Range.Start));
                Dump(doc, "正文：零宽点在目标一段落符上", doc.Range(p2.Range.End - 1, p2.Range.End - 1));
                Console.WriteLine("plain-p1-start=" + p1.Range.Start);
            }
            finally
            {
                Close(doc);
            }
        }

        private static void RunTable(Word.Application app)
        {
            Word.Document doc = app.Documents.Add();
            try
            {
                doc.Content.Text = "表前。\r";
                Word.Table table = doc.Tables.Add(doc.Paragraphs[doc.Paragraphs.Count].Range, 2, 2);
                table.Cell(1, 1).Range.Text = "甲一\r甲二";
                table.Cell(1, 2).Range.Text = "乙";
                table.Cell(2, 2).Range.Text = "丁";
                Word.Range tail = doc.Range(doc.Content.End - 1, doc.Content.End - 1);
                tail.InsertAfter("表后。");
                DumpDoc(doc, "表格全文");

                Word.Range cell11 = table.Cell(1, 1).Range;
                Word.Range cell12 = table.Cell(1, 2).Range;
                Word.Range cell21 = table.Cell(2, 1).Range;
                Word.Range cell22 = table.Cell(2, 2).Range;
                Dump(doc, "表：单元格(1,1)整格", cell11);
                Dump(doc, "表：单元格(1,1)去掉格末标记", doc.Range(cell11.Start, cell11.End - 1));
                Dump(doc, "表：空单元格(2,1)整格", cell21);
                Dump(doc, "表：同一行两格 (1,1) 到 (1,2)", doc.Range(cell11.Start, cell12.End));
                Dump(doc, "表：整张表 Range", table.Range);
                Dump(doc, "表：表前段落起点到(1,1)格末", doc.Range(doc.Paragraphs[1].Range.Start, cell11.End));
                Dump(doc, "表：(2,2)格起点到文末", doc.Range(cell22.Start, doc.Content.End));
            }
            finally
            {
                Close(doc);
            }
        }

        private static void RunNumbered(Word.Application app)
        {
            Word.Document doc = app.Documents.Add();
            try
            {
                doc.Content.Text = "一条\r二条\r三条\r";
                Word.Range all = doc.Range(doc.Paragraphs[1].Range.Start, doc.Paragraphs[3].Range.End);
                all.ListFormat.ApplyNumberDefault();
                DumpDoc(doc, "编号全文");
                Word.Paragraph mid = doc.Paragraphs[2];
                Dump(doc, "编号：只圈第二段", doc.Range(mid.Range.Start, mid.Range.End));
                Dump(doc, "编号：第二段中间到第三段中间", doc.Range(mid.Range.Start + 1, doc.Paragraphs[3].Range.End - 2));
            }
            finally
            {
                Close(doc);
            }
        }

        private static void RunTableFind(Word.Application app)
        {
            Word.Document doc = app.Documents.Add();
            try
            {
                doc.Content.Text = "针前。\r";
                Word.Table table = doc.Tables.Add(doc.Paragraphs[doc.Paragraphs.Count].Range, 2, 2);
                table.Cell(1, 1).Range.Text = "针甲";
                table.Cell(1, 2).Range.Text = "目标";
                table.Cell(2, 1).Range.Text = "甲针\r乙针";
                FindIn(doc, "表：在(1,2)「目标」里找「针」，前面正文和(1,1)都有", table.Cell(1, 2).Range, "针");
                Word.Range cell21 = table.Cell(2, 1).Range;
                Word.Paragraph second = null;
                foreach (Word.Paragraph p in cell21.Paragraphs)
                {
                    second = p;
                }

                if (second != null)
                {
                    FindIn(doc, "表：在(2,1)第二段里找「针」，同格第一段也有", second.Range, "针");
                }
            }
            finally
            {
                Close(doc);
            }
        }

        private static void FindIn(Word.Document doc, string title, Word.Range scope, string needle)
        {
            int scopeStart = scope.Start;
            int scopeEnd = scope.End;
            Word.Find find = scope.Find;
            find.ClearFormatting();
            find.Replacement.ClearFormatting();
            object findText = needle;
            object falseArg = false;
            object trueArg = true;
            object wrap = Word.WdFindWrap.wdFindStop;
            object missing = Type.Missing;
            bool ok = find.Execute(
                findText,
                falseArg,
                falseArg,
                falseArg,
                falseArg,
                falseArg,
                trueArg,
                wrap,
                falseArg,
                missing,
                missing);
            Console.WriteLine("== Find " + title);
            Console.WriteLine(
                "  scope " + scopeStart + ".." + scopeEnd
                + " text=" + Esc(SafeText(doc.Range(scopeStart, scopeEnd)))
                + " execute=" + ok
                + " hit " + scope.Start + ".." + scope.End
                + " beforeScope=" + (scope.Start < scopeStart)
                + " hitText=" + Esc(SafeText(scope)));
        }

        private static void RunFind(Word.Application app)
        {
            Word.Document doc = app.Documents.Add();
            try
            {
                doc.Content.Text = "针前。\r目标针后。\r";
                Word.Paragraph target = doc.Paragraphs[2];
                Word.Range scope = doc.Range(target.Range.Start, target.Range.End);
                int scopeStart = scope.Start;
                int scopeEnd = scope.End;
                Word.Find find = scope.Find;
                find.ClearFormatting();
                find.Replacement.ClearFormatting();
                object findText = "针";
                object falseArg = false;
                object trueArg = true;
                object wrap = Word.WdFindWrap.wdFindStop;
                object missing = Type.Missing;
                bool ok = find.Execute(
                    findText,
                    falseArg,
                    falseArg,
                    falseArg,
                    falseArg,
                    falseArg,
                    trueArg,
                    wrap,
                    falseArg,
                    missing,
                    missing);
                Console.WriteLine("== Find：第二段内找「针」，第一段也有「针」");
                Console.WriteLine(
                    "  scope " + scopeStart + ".." + scopeEnd
                    + " execute=" + ok
                    + " hit " + scope.Start + ".." + scope.End
                    + " beforeScope=" + (scope.Start < scopeStart)
                    + " text=" + Esc(SafeText(scope)));
            }
            finally
            {
                Close(doc);
            }
        }

        private static void DumpDoc(Word.Document doc, string title)
        {
            Console.WriteLine("== " + title + "  doc.Paragraphs");
            for (int i = 1; i <= doc.Paragraphs.Count; i++)
            {
                Word.Range pr = doc.Paragraphs[i].Range;
                Console.WriteLine(
                    "  i=" + i
                    + " " + pr.Start + ".." + pr.End
                    + " inTable=" + InTable(pr)
                    + " ls=\"" + ListString(doc.Paragraphs[i]) + "\""
                    + " text=" + Esc(SafeText(pr)));
            }
        }

        private static void Dump(Word.Document doc, string title, Word.Range range)
        {
            int start = range.Start;
            int end = range.End;
            Console.WriteLine("== " + title);
            Console.WriteLine("  range " + start + ".." + end + " text=" + Esc(SafeText(range)));
            Word.Paragraphs paras = range.Paragraphs;
            int count = paras.Count;
            int startBefore = 0;
            int startAtOrAfterEnd = 0;
            for (int i = 1; i <= count; i++)
            {
                Word.Paragraph p = paras[i];
                Word.Range pr = p.Range;
                string flag = "INSIDE";
                if (pr.Start < start)
                {
                    flag = "START_BEFORE";
                    startBefore++;
                }
                else if (pr.Start >= end)
                {
                    flag = "START_AT_OR_AFTER_END";
                    startAtOrAfterEnd++;
                }
                else if (pr.End > end)
                {
                    flag = "ENDS_PAST";
                }

                Console.WriteLine(
                    "  rp[" + i + "] " + flag
                    + " docPara=" + DocIndex(doc, pr.Start)
                    + " " + pr.Start + ".." + pr.End
                    + " inTable=" + InTable(pr)
                    + " ls=\"" + ListString(p) + "\""
                    + " text=" + Esc(SafeText(pr)));
            }

            Console.WriteLine(
                "  range.Paragraphs.Count=" + count
                + " startBefore=" + startBefore
                + " startAtOrAfterEnd=" + startAtOrAfterEnd
                + " rangeAfterEnum " + range.Start + ".." + range.End);
        }

        private static int DocIndex(Word.Document doc, int start)
        {
            for (int i = 1; i <= doc.Paragraphs.Count; i++)
            {
                if (doc.Paragraphs[i].Range.Start == start)
                {
                    return i;
                }
            }

            return -1;
        }

        private static string ListString(Word.Paragraph paragraph)
        {
            try
            {
                return (paragraph.Range.ListFormat.ListString ?? "").Trim();
            }
            catch (Exception ex)
            {
                return "ERR " + ex.Message;
            }
        }

        private static bool InTable(Word.Range range)
        {
            try
            {
                object value = range.Information[Word.WdInformation.wdWithInTable];
                if (value is bool flag)
                {
                    return flag;
                }

                return Convert.ToInt32(value) != 0;
            }
            catch
            {
                return false;
            }
        }

        private static string SafeText(Word.Range range)
        {
            try
            {
                return range.Text ?? "";
            }
            catch (Exception ex)
            {
                return "ERR " + ex.Message;
            }
        }

        private static string Esc(string text)
        {
            if (text == null)
            {
                return "(null)";
            }

            return text
                .Replace("\\", "\\\\")
                .Replace("\r", "\\r")
                .Replace("\n", "\\n")
                .Replace("\a", "\\a")
                .Replace("\f", "\\f")
                .Replace("\v", "\\v");
        }

        private static void Close(Word.Document doc)
        {
            try
            {
                doc.Close(false);
            }
            catch
            {
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

                return null;
            }

            try
            {
                return (Word.Application)Marshal.GetActiveObject("Word.Application");
            }
            catch
            {
                Word.Application createdApp = new Word.Application();
                createdApp.Visible = false;
                created = true;
                return createdApp;
            }
        }
    }
}
