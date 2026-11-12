using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using WordAddIn1;
using Word = Microsoft.Office.Interop.Word;

namespace WordDocumentRoundtripTest
{
    /// <summary>
    /// 会话 1156 的分片写法：正文片不写上标，中间一片 [1] 写上标。
    /// 这几片没有换行，写路径就是补默认后 ApplySnapshot。旧路径只套传入键。
    /// </summary>
    internal static class SuperBleedProbe
    {
        private const string Head = "误报较多";
        private const string Cite = "[1]";
        private const string Tail = "。张敏和陈涛采用模型。";

        public static int Show()
        {
            bool created = false;
            Word.Application app = Acquire(out created);
            if (app == null)
            {
                Console.WriteLine("没有 Word，文档没打开。");
                return 1;
            }

            try
            {
                app.Visible = true;
                Word.Document doc = app.Documents.Add();
                doc.Content.Text = "旧路径（修复前）\r占位旧\r新路径（修复后）\r占位新\r";
                ReplacePlaceholder(doc, "占位旧", useFix: false);
                ReplacePlaceholder(doc, "占位新", useFix: true);
                doc.Activate();
                try
                {
                    app.Activate();
                }
                catch
                {
                }

                Console.WriteLine("已打开对照稿，未关闭。");
                Console.WriteLine("旧路径 " + DumpParagraphAt(doc, 2));
                Console.WriteLine("新路径 " + DumpParagraphAt(doc, 4));
                return 0;
            }
            catch (Exception ex)
            {
                Console.WriteLine("打开对照稿失败: " + ex.Message);
                if (created)
                {
                    try
                    {
                        object save = false;
                        app.Quit(ref save);
                    }
                    catch
                    {
                    }
                }

                return 1;
            }
        }

        private static void ReplacePlaceholder(Word.Document doc, string placeholder, bool useFix)
        {
            Word.Range hit = doc.Range(doc.Content.Start, doc.Content.End);
            object findText = placeholder;
            object missing = Type.Missing;
            bool found = hit.Find.Execute(
                ref findText,
                ref missing, ref missing, ref missing, ref missing,
                ref missing, ref missing, ref missing, ref missing,
                ref missing, ref missing, ref missing, ref missing,
                ref missing, ref missing);
            if (!found)
            {
                throw new InvalidOperationException("找不到 " + placeholder);
            }

            Word.Range cursor = WriteRun(doc, hit, Head, BodyFormat(), useFix, replace: true);
            cursor = WriteRun(doc, cursor, Cite, CiteFormat(), useFix, replace: false);
            WriteRun(doc, cursor, Tail, BodyFormat(), useFix, replace: false);
        }

        private static string DumpParagraphAt(Word.Document doc, int index)
        {
            Word.Range para = doc.Paragraphs[index].Range;
            int end = para.End;
            var parts = new List<string>();
            int i = para.Start;
            while (i < end)
            {
                Word.Range ch = doc.Range(i, i + 1);
                string text = ch.Text ?? "";
                if (text == "\r" || text.Length == 0)
                {
                    i++;
                    continue;
                }

                bool on = IsOn(ch.Font.Superscript);
                if (parts.Count > 0 && parts[parts.Count - 1].EndsWith(on ? "=上标" : "=正文", StringComparison.Ordinal))
                {
                    int eq = parts[parts.Count - 1].LastIndexOf('=');
                    parts[parts.Count - 1] = parts[parts.Count - 1].Substring(0, eq) + text + parts[parts.Count - 1].Substring(eq);
                }
                else
                {
                    parts.Add(text + (on ? "=上标" : "=正文"));
                }

                i++;
            }

            return string.Join(" | ", parts);
        }

        public static int Run()
        {
            bool created = false;
            Word.Application app = Acquire(out created);
            if (app == null)
            {
                Console.WriteLine("没有 Word，探针没跑。");
                return 1;
            }

            bool oldScreen = true;
            try
            {
                try
                {
                    oldScreen = app.ScreenUpdating;
                    app.ScreenUpdating = false;
                }
                catch
                {
                }

                bool oldLeaks = WriteAndRead(app, useFix: false);
                bool newLeaks = WriteAndRead(app, useFix: true);
                Console.WriteLine(oldLeaks
                    ? "旧路径：后面的字变成了上标"
                    : "旧路径：后面的字不是上标");
                Console.WriteLine(newLeaks
                    ? "新路径：后面的字变成了上标"
                    : "新路径：后面的字不是上标");
                return newLeaks ? 1 : 0;
            }
            finally
            {
                try
                {
                    app.ScreenUpdating = oldScreen;
                }
                catch
                {
                }

                if (created)
                {
                    try
                    {
                        object save = false;
                        app.Quit(ref save);
                    }
                    catch
                    {
                    }
                }

                try
                {
                    Marshal.ReleaseComObject(app);
                }
                catch
                {
                }
            }
        }

        private static bool WriteAndRead(Word.Application app, bool useFix)
        {
            Word.Document doc = null;
            try
            {
                doc = app.Documents.Add();
                Word.Range para = doc.Paragraphs[1].Range;
                int end = para.End;
                if (end > para.Start && doc.Range(end - 1, end).Text == "\r")
                {
                    end--;
                }

                Word.Range cursor = doc.Range(para.Start, end);
                cursor = WriteRun(doc, cursor, Head, BodyFormat(), useFix, replace: true);
                cursor = WriteRun(doc, cursor, Cite, CiteFormat(), useFix, replace: false);
                WriteRun(doc, cursor, Tail, BodyFormat(), useFix, replace: false);

                string story = DumpParagraph(doc);
                Console.WriteLine((useFix ? "新路径" : "旧路径") + " " + story);
                return TailIsSuperscript(doc);
            }
            finally
            {
                if (doc != null)
                {
                    try
                    {
                        object save = false;
                        doc.Close(ref save);
                    }
                    catch
                    {
                    }
                }
            }
        }

        private static Word.Range WriteRun(
            Word.Document doc,
            Word.Range cursor,
            string text,
            Dictionary<string, object> format,
            bool useFix,
            bool replace)
        {
            Word.Range writeAt = cursor;
            if (!replace)
            {
                writeAt = doc.Range(cursor.End, cursor.End);
            }

            writeAt.Text = text;
            Word.Range written = doc.Range(writeAt.Start, writeAt.End);
            var snap = new Dictionary<string, object>(format);
            if (useFix)
            {
                FormatInheritHelper.FillUnspecifiedWriteDefaults(snap);
            }

            FormatInheritHelper.EnsureWrittenBackground(snap);
            FormatInheritHelper.ApplySnapshot(written, snap, true);

            return written;
        }

        private static bool TailIsSuperscript(Word.Document doc)
        {
            Word.Range para = doc.Paragraphs[1].Range;
            string all = para.Text ?? "";
            int tailAt = all.IndexOf(Tail, StringComparison.Ordinal);
            if (tailAt < 0)
            {
                Console.WriteLine("段里找不到后文");
                return true;
            }

            int start = para.Start + tailAt;
            for (int i = 0; i < Tail.Length; i++)
            {
                Word.Range ch = doc.Range(start + i, start + i + 1);
                if (IsOn(ch.Font.Superscript))
                {
                    return true;
                }
            }

            return false;
        }

        private static string DumpParagraph(Word.Document doc)
        {
            Word.Range para = doc.Paragraphs[1].Range;
            int end = para.End;
            var parts = new List<string>();
            int i = para.Start;
            while (i < end)
            {
                Word.Range ch = doc.Range(i, i + 1);
                string text = ch.Text ?? "";
                if (text == "\r" || text.Length == 0)
                {
                    i++;
                    continue;
                }

                bool on = IsOn(ch.Font.Superscript);
                if (parts.Count > 0 && parts[parts.Count - 1].EndsWith(on ? "=上标" : "=正文", StringComparison.Ordinal))
                {
                    int eq = parts[parts.Count - 1].LastIndexOf('=');
                    parts[parts.Count - 1] = parts[parts.Count - 1].Substring(0, eq) + text + parts[parts.Count - 1].Substring(eq);
                }
                else
                {
                    parts.Add(text + (on ? "=上标" : "=正文"));
                }

                i++;
            }

            return string.Join(" | ", parts);
        }

        private static bool IsOn(int superscript)
        {
            return superscript != 0 && superscript != 9999999;
        }

        private static Dictionary<string, object> BodyFormat()
        {
            return new Dictionary<string, object>
            {
                ["font_name"] = "宋体",
                ["font_size"] = 12,
                ["is_bold"] = false,
                ["name_bi"] = "Times New Roman"
            };
        }

        private static Dictionary<string, object> CiteFormat()
        {
            Dictionary<string, object> format = BodyFormat();
            format["is_superscript"] = true;
            return format;
        }

        private static Word.Application Acquire(out bool created)
        {
            created = false;
            try
            {
                return (Word.Application)Marshal.GetActiveObject("Word.Application");
            }
            catch
            {
                Word.Application app = new Word.Application();
                app.Visible = false;
                created = true;
                return app;
            }
        }
    }
}
