using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using WordAddIn1;

namespace WordDocumentRoundtripTest
{
    /// <summary>
    /// 只测分页 \f / 分节 ^b 的读入、写入、回读 Find、OOXML。
    /// </summary>
    internal static class LbCases
    {
        public const int Target = 35;

        private const string Cover = "封面日期甲。";
        private const string Overview = "总概述乙。";
        private const string Chapter = "第一章丙。";
        private const string WriteA = "分页写入甲。";
        private const string WriteB = "分页写入乙。";
        private const string SectA = "分节写入甲。";
        private const string SectB = "分节写入乙。";
        private const string MixA = "混写封面。";
        private const string MixB = "混写概述。";
        private const string MixC = "混写正文。";

        public static void Run(TestRun run, string host)
        {
            var s = new WdSession(host);
            RunReadComFixture(run, s);
            RunWriteInsertPage(run, s);
            RunWriteInsertSection(run, s);
            RunWriteMixRoundtrip(run, s);
            RunWriteReplaceDropPage(run, s);
            RunInsertPageBreakTool(run, s);
        }

        private static void RunReadComFixture(TestRun run, WdSession s)
        {
            string path = s.NewDocxPath();
            ToolResult opened = s.Open(path, true, s.App);
            if (!opened.Success)
            {
                FailMany(run, opened.Error ?? "新建失败",
                    "lb-read-open", "lb-read-com-write", "lb-read-text",
                    "lb-read-has-f", "lb-read-has-b", "lb-read-f-count", "lb-read-b-count",
                    "lb-read-glue-f", "lb-read-glue-b",
                    "lb-read-s-find", "lb-read-p-find",
                    "lb-read-ooxml-page", "lb-read-ooxml-sect");
                return;
            }

            Ok(run, "lb-read-open", true, "");
            string id = WdSession.Str(opened, "channel_id");
            if (!s.WriteLayoutBreakFixture(path))
            {
                FailMany(run, "COM 写分页/分节夹具失败",
                    "lb-read-com-write", "lb-read-text",
                    "lb-read-has-f", "lb-read-has-b", "lb-read-f-count", "lb-read-b-count",
                    "lb-read-glue-f", "lb-read-glue-b",
                    "lb-read-s-find", "lb-read-p-find",
                    "lb-read-ooxml-page", "lb-read-ooxml-sect");
                s.CloseDocByPath(path);
                return;
            }

            Ok(run, "lb-read-com-write", true, "");

            ToolResult got = s.GetContent(id);
            if (!got.Success)
            {
                FailMany(run, got.Error ?? "读文失败",
                    "lb-read-text",
                    "lb-read-has-f", "lb-read-has-b", "lb-read-f-count", "lb-read-b-count",
                    "lb-read-glue-f", "lb-read-glue-b",
                    "lb-read-s-find", "lb-read-p-find",
                    "lb-read-ooxml-page", "lb-read-ooxml-sect");
                s.CloseDocByPath(path);
                return;
            }

            string text = WdSession.Str(got, "document_content");
            Ok(run, "lb-read-text",
                text.IndexOf(Cover, StringComparison.Ordinal) >= 0
                && text.IndexOf(Overview, StringComparison.Ordinal) >= 0
                && text.IndexOf(Chapter, StringComparison.Ordinal) >= 0,
                Preview(text));
            Ok(run, "lb-read-has-f", text.IndexOf('\f') >= 0, "display 不见 \\f");
            Ok(run, "lb-read-has-b", text.IndexOf("^b", StringComparison.Ordinal) >= 0, "display 不见 ^b");
            Ok(run, "lb-read-f-count", CountChar(text, '\f') == 1,
                "\\f 次数=" + CountChar(text, '\f') + "（lastRendered 不应进流）");
            Ok(run, "lb-read-b-count", CountToken(text, "^b") == 1,
                "^b 次数=" + CountToken(text, "^b") + "（body sectPr 不应多出）");

            List<string> sentences = ParseStoredSpans(text, "S");
            Ok(run, "lb-read-glue-f", AnyContains(sentences, Cover) && AnyContainsChar(sentences, '\f'),
                "封面句未粘到 \\f");
            Ok(run, "lb-read-glue-b", AnyContains(sentences, Overview) && AnyContainsToken(sentences, "^b"),
                "概述句未粘到 ^b");

            string sMiss = null;
            bool sFound = sentences.Count > 0 && s.FindAllStoredTexts(path, sentences, out sMiss);
            Ok(run, "lb-read-s-find", sFound, sentences.Count == 0 ? "未解析到 S_" : (sMiss ?? "Find 失败"));

            ToolResult para = s.GetContent(id, "paragraph");
            if (!para.Success)
            {
                run.CaseFail("lb-read-p-find", para.Error ?? "段落读失败");
            }
            else
            {
                List<string> paras = ParseStoredSpans(WdSession.Str(para, "document_content"), "P");
                string pMiss = null;
                bool pFound = paras.Count > 0 && s.FindAllStoredTexts(path, paras, out pMiss);
                Ok(run, "lb-read-p-find", pFound, paras.Count == 0 ? "未解析到 P_" : (pMiss ?? "Find 失败"));
            }

            string xml = s.GetLiveWordOpenXml(path);
            Ok(run, "lb-read-ooxml-page", WdSession.XmlHasPageBreak(xml), "OOXML 无 w:br type=page");
            Ok(run, "lb-read-ooxml-sect", WdSession.XmlHasParagraphSectPr(xml), "OOXML 段上无 sectPr");
            s.CloseDocByPath(path);
        }

        private static void RunWriteInsertPage(TestRun run, WdSession s)
        {
            string path;
            string id;
            if (!TryOpen(run, s, "lb-wr-f-open", out path, out id,
                "lb-wr-f-insert", "lb-wr-f-read", "lb-wr-f-ooxml", "lb-wr-f-find"))
            {
                return;
            }

            string payload = WriteA + "\f" + WriteB;
            ToolResult wr = s.ProcessActions(id, InsertStart(payload));
            Ok(run, "lb-wr-f-insert", wr.Success, wr.Error ?? "insert \\f 失败");
            if (!wr.Success)
            {
                FailMany(run, wr.Error ?? "insert 失败", "lb-wr-f-read", "lb-wr-f-ooxml", "lb-wr-f-find");
                s.CloseDocByPath(path);
                return;
            }

            AssertReadFindXml(run, s, path, id,
                expectF: true, expectB: false,
                "lb-wr-f-read", "lb-wr-f-ooxml", "lb-wr-f-find");
            s.CloseDocByPath(path);
        }

        private static void RunWriteInsertSection(TestRun run, WdSession s)
        {
            string path;
            string id;
            if (!TryOpen(run, s, "lb-wr-b-open", out path, out id,
                "lb-wr-b-insert", "lb-wr-b-read", "lb-wr-b-ooxml", "lb-wr-b-find"))
            {
                return;
            }

            string payload = SectA + "^b" + SectB;
            ToolResult wr = s.ProcessActions(id, InsertStart(payload));
            Ok(run, "lb-wr-b-insert", wr.Success, wr.Error ?? "insert ^b 失败");
            if (!wr.Success)
            {
                FailMany(run, wr.Error ?? "insert 失败", "lb-wr-b-read", "lb-wr-b-ooxml", "lb-wr-b-find");
                s.CloseDocByPath(path);
                return;
            }

            ToolResult got = s.GetContent(id);
            if (!got.Success)
            {
                FailMany(run, got.Error ?? "读文失败", "lb-wr-b-read", "lb-wr-b-ooxml", "lb-wr-b-find");
                s.CloseDocByPath(path);
                return;
            }

            string text = WdSession.Str(got, "document_content");
            Ok(run, "lb-wr-b-read",
                text.IndexOf(SectA, StringComparison.Ordinal) >= 0
                && text.IndexOf(SectB, StringComparison.Ordinal) >= 0
                && text.IndexOf("^b", StringComparison.Ordinal) >= 0,
                Preview(text));
            Ok(run, "lb-wr-b-ooxml", WdSession.XmlHasParagraphSectPr(s.GetLiveWordOpenXml(path)),
                "OOXML 段上无 sectPr");
            string miss = null;
            List<string> stored = ParseStoredSpans(text, "S");
            Ok(run, "lb-wr-b-find",
                stored.Count > 0 && s.FindAllStoredTexts(path, stored, out miss),
                stored.Count == 0 ? "未解析到 S_" : (miss ?? "Find 失败"));
            s.CloseDocByPath(path);
        }

        private static void RunWriteMixRoundtrip(TestRun run, WdSession s)
        {
            string path;
            string id;
            if (!TryOpen(run, s, "lb-wr-mix-open", out path, out id,
                "lb-wr-mix-insert", "lb-wr-mix-read", "lb-wr-mix-find"))
            {
                return;
            }

            string payload = MixA + "\f" + MixB + "^b" + MixC;
            ToolResult wr = s.ProcessActions(id, InsertStart(payload));
            Ok(run, "lb-wr-mix-insert", wr.Success, wr.Error ?? "混写失败");
            if (!wr.Success)
            {
                FailMany(run, wr.Error ?? "混写失败", "lb-wr-mix-read", "lb-wr-mix-find");
                s.CloseDocByPath(path);
                return;
            }

            ToolResult got = s.GetContent(id);
            if (!got.Success)
            {
                FailMany(run, got.Error ?? "读文失败", "lb-wr-mix-read", "lb-wr-mix-find");
                s.CloseDocByPath(path);
                return;
            }

            string text = WdSession.Str(got, "document_content");
            Ok(run, "lb-wr-mix-read",
                text.IndexOf(MixA, StringComparison.Ordinal) >= 0
                && text.IndexOf(MixB, StringComparison.Ordinal) >= 0
                && text.IndexOf(MixC, StringComparison.Ordinal) >= 0
                && text.IndexOf('\f') >= 0
                && text.IndexOf("^b", StringComparison.Ordinal) >= 0,
                Preview(text));
            string miss = null;
            List<string> stored = ParseStoredSpans(text, "S");
            Ok(run, "lb-wr-mix-find",
                stored.Count > 0 && s.FindAllStoredTexts(path, stored, out miss),
                stored.Count == 0 ? "未解析到 S_" : (miss ?? "Find 失败"));
            s.CloseDocByPath(path);
        }

        private static void RunWriteReplaceDropPage(TestRun run, WdSession s)
        {
            string path;
            string id;
            if (!TryOpen(run, s, "lb-wr-rep-open", out path, out id,
                "lb-wr-rep-seed", "lb-wr-rep-drop", "lb-wr-rep-gone"))
            {
                return;
            }

            ToolResult seed = s.ProcessActions(id, InsertStart(WriteA + "\f" + WriteB));
            Ok(run, "lb-wr-rep-seed", seed.Success, seed.Error ?? "铺分页失败");
            if (!seed.Success)
            {
                FailMany(run, seed.Error ?? "铺分页失败", "lb-wr-rep-drop", "lb-wr-rep-gone");
                s.CloseDocByPath(path);
                return;
            }

            ToolResult got = s.GetContent(id);
            string code = FirstCode(WdSession.Str(got, "document_content"), "S");
            if (string.IsNullOrEmpty(code))
            {
                FailMany(run, "铺完后无 S_", "lb-wr-rep-drop", "lb-wr-rep-gone");
                s.CloseDocByPath(path);
                return;
            }

            ToolResult rep = s.ProcessActions(id, ReplaceCode(code, WriteA));
            Ok(run, "lb-wr-rep-drop", rep.Success, rep.Error ?? "replace 去掉 \\f 失败");
            if (!rep.Success)
            {
                run.CaseFail("lb-wr-rep-gone", rep.Error ?? "replace 失败");
                s.CloseDocByPath(path);
                return;
            }

            ToolResult after = s.GetContent(id);
            string text = WdSession.Str(after, "document_content");
            bool gone = after.Success
                && text.IndexOf('\f') < 0
                && !WdSession.XmlHasPageBreak(s.GetLiveWordOpenXml(path));
            Ok(run, "lb-wr-rep-gone", gone,
                after.Success ? "display 或 OOXML 仍有分页" : (after.Error ?? "再读失败"));
            s.CloseDocByPath(path);
        }

        private static void RunInsertPageBreakTool(TestRun run, WdSession s)
        {
            string path;
            string id;
            if (!TryOpen(run, s, "lb-tool-open", out path, out id,
                "lb-tool-insert", "lb-tool-read", "lb-tool-ooxml"))
            {
                return;
            }

            ToolResult seed = s.ProcessActions(id, InsertStart(WriteA + WriteB));
            if (!seed.Success)
            {
                FailMany(run, seed.Error ?? "铺字失败", "lb-tool-insert", "lb-tool-read", "lb-tool-ooxml");
                s.CloseDocByPath(path);
                return;
            }

            ToolResult got = s.GetContent(id);
            string code = FirstCode(WdSession.Str(got, "document_content"), "S");
            var target = new Dictionary<string, object>();
            if (!string.IsNullOrEmpty(code))
            {
                target["codes"] = code;
            }
            else
            {
                target["position"] = "start";
            }

            ToolResult ins = s.InsertPageBreakTool(id, target);
            Ok(run, "lb-tool-insert", ins.Success, ins.Error ?? "F_insert_page_break 失败");
            if (!ins.Success)
            {
                FailMany(run, ins.Error ?? "插分页失败", "lb-tool-read", "lb-tool-ooxml");
                s.CloseDocByPath(path);
                return;
            }

            ToolResult after = s.GetContent(id);
            Ok(run, "lb-tool-read",
                after.Success && WdSession.Str(after, "document_content").IndexOf('\f') >= 0,
                after.Success ? "display 不见 \\f" : (after.Error ?? "读文失败"));
            Ok(run, "lb-tool-ooxml",
                WdSession.XmlHasPageBreak(s.GetLiveWordOpenXml(path)),
                "OOXML 无 w:br type=page");
            s.CloseDocByPath(path);
        }

        private static bool TryOpen(
            TestRun run,
            WdSession s,
            string openId,
            out string path,
            out string id,
            params string[] rest)
        {
            path = s.NewDocxPath();
            id = null;
            ToolResult opened = s.Open(path, true, s.App);
            if (!opened.Success)
            {
                FailMany(run, opened.Error ?? "新建失败", Prepend(openId, rest));
                return false;
            }

            Ok(run, openId, true, "");
            id = WdSession.Str(opened, "channel_id");
            return true;
        }

        private static void AssertReadFindXml(
            TestRun run,
            WdSession s,
            string path,
            string id,
            bool expectF,
            bool expectB,
            string readId,
            string xmlId,
            string findId)
        {
            ToolResult got = s.GetContent(id);
            if (!got.Success)
            {
                FailMany(run, got.Error ?? "读文失败", readId, xmlId, findId);
                return;
            }

            string text = WdSession.Str(got, "document_content");
            bool readOk = text.IndexOf(WriteA, StringComparison.Ordinal) >= 0
                && text.IndexOf(WriteB, StringComparison.Ordinal) >= 0
                && (!expectF || text.IndexOf('\f') >= 0)
                && (!expectB || text.IndexOf("^b", StringComparison.Ordinal) >= 0);
            Ok(run, readId, readOk, Preview(text));

            string xml = s.GetLiveWordOpenXml(path);
            bool xmlOk = (!expectF || WdSession.XmlHasPageBreak(xml))
                && (!expectB || WdSession.XmlHasParagraphSectPr(xml));
            Ok(run, xmlId, xmlOk, expectF ? "OOXML 无分页符" : "OOXML 无分节符");

            string miss = null;
            List<string> stored = ParseStoredSpans(text, "S");
            Ok(run, findId,
                stored.Count > 0 && s.FindAllStoredTexts(path, stored, out miss),
                stored.Count == 0 ? "未解析到 S_" : (miss ?? "Find 失败"));
        }

        private static List<Dictionary<string, object>> InsertStart(string insert)
        {
            return new List<Dictionary<string, object>>
            {
                new Dictionary<string, object>
                {
                    ["type"] = "insert",
                    ["detail"] = new Dictionary<string, object>
                    {
                        ["target"] = new Dictionary<string, object> { ["position"] = "start" },
                        ["insert"] = insert,
                        ["format"] = DefaultFormat()
                    }
                }
            };
        }

        private static List<Dictionary<string, object>> ReplaceCode(string code, string neu)
        {
            return new List<Dictionary<string, object>>
            {
                new Dictionary<string, object>
                {
                    ["type"] = "replace",
                    ["detail"] = new Dictionary<string, object>
                    {
                        ["original"] = new Dictionary<string, object> { ["codes"] = code },
                        ["new"] = neu,
                        ["format"] = DefaultFormat()
                    }
                }
            };
        }

        private static Dictionary<string, object> DefaultFormat()
        {
            return new Dictionary<string, object>
            {
                ["char_format"] = new Dictionary<string, object>
                {
                    ["font_name"] = "宋体",
                    ["font_size"] = 12,
                    ["is_bold"] = false
                }
            };
        }

        private static List<string> ParseStoredSpans(string display, string prefix)
        {
            var list = new List<string>();
            if (string.IsNullOrEmpty(display))
            {
                return list;
            }

            foreach (Match m in Regex.Matches(
                display,
                @"\[" + prefix + @"_([0-9a-zA-Z]{5})-start\](.*?)(\[" + prefix + @"_\1-end\])",
                RegexOptions.Singleline))
            {
                string stored = m.Groups[2].Value;
                if (!string.IsNullOrEmpty(stored))
                {
                    list.Add(stored);
                }
            }

            return list;
        }

        private static string FirstCode(string display, string prefix)
        {
            if (string.IsNullOrEmpty(display))
            {
                return null;
            }

            Match m = Regex.Match(display, @"\[" + prefix + @"_([0-9a-zA-Z]{5})-start\]");
            return m.Success ? prefix + "_" + m.Groups[1].Value : null;
        }

        private static bool AnyContains(List<string> spans, string needle)
        {
            if (spans == null)
            {
                return false;
            }

            for (int i = 0; i < spans.Count; i++)
            {
                if (spans[i] != null && spans[i].IndexOf(needle, StringComparison.Ordinal) >= 0)
                {
                    return true;
                }
            }

            return false;
        }

        private static bool AnyContainsChar(List<string> spans, char ch)
        {
            if (spans == null)
            {
                return false;
            }

            for (int i = 0; i < spans.Count; i++)
            {
                if (spans[i] != null && spans[i].IndexOf(ch) >= 0)
                {
                    return true;
                }
            }

            return false;
        }

        private static bool AnyContainsToken(List<string> spans, string token)
        {
            return AnyContains(spans, token);
        }

        private static int CountChar(string text, char ch)
        {
            if (string.IsNullOrEmpty(text))
            {
                return 0;
            }

            int n = 0;
            for (int i = 0; i < text.Length; i++)
            {
                if (text[i] == ch)
                {
                    n++;
                }
            }

            return n;
        }

        private static int CountToken(string text, string token)
        {
            if (string.IsNullOrEmpty(text) || string.IsNullOrEmpty(token))
            {
                return 0;
            }

            int n = 0;
            int i = 0;
            while (i >= 0 && i < text.Length)
            {
                int at = text.IndexOf(token, i, StringComparison.Ordinal);
                if (at < 0)
                {
                    break;
                }

                n++;
                i = at + token.Length;
            }

            return n;
        }

        private static string Preview(string text)
        {
            if (string.IsNullOrEmpty(text))
            {
                return "";
            }

            string t = TextUtils.EscapeNewlinesForDb(text);
            return t.Length <= 60 ? t : t.Substring(0, 60);
        }

        private static string[] Prepend(string first, string[] rest)
        {
            var list = new List<string> { first };
            if (rest != null)
            {
                list.AddRange(rest);
            }

            return list.ToArray();
        }

        private static void Ok(TestRun run, string id, bool cond, string reason)
        {
            if (cond)
            {
                run.CaseOk(id);
            }
            else
            {
                run.CaseFail(id, reason ?? "");
            }
        }

        private static void FailMany(TestRun run, string reason, params string[] ids)
        {
            foreach (string id in ids)
            {
                run.CaseFail(id, reason);
            }
        }
    }
}
