using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using WordAddIn1;

namespace WordDocumentRoundtripTest
{
    /// <summary>
    /// ClearExtra 改用写入范围取段落。断言 process_document_actions 会话里的 [autonum] 行。
    /// </summary>
    internal static class CeCases
    {
        public const int Target = 8;

        public static string IdFilter;

        public static void Run(
            TestRun run,
            string host,
            int wordBefore,
            int wpsBefore,
            bool wordWinBefore,
            bool wpsWinBefore)
        {
            var s = new WdSession(host);
            if (Want("ce-rep-tail-clear"))
            {
                RunReplace(
                    run, s, "ce-rep-tail-clear",
                    "第一条。\r", true, 0,
                    "纯故事",
                    "numbered=True",
                    "TryFindReplace before=",
                    "ClearExtra rp=1",
                    "wantPrefix=False",
                    "remove=True",
                    "ClearExtra done checked=1");
            }

            if (Want("ce-rep-tail-keep"))
            {
                RunReplace(
                    run, s, "ce-rep-tail-keep",
                    "第一条。\r", true, 0,
                    "2. 新条",
                    "numbered=True",
                    "TryFindReplace before=",
                    "ClearExtra rp=1",
                    "wantPrefix=True",
                    "remove=False",
                    "ClearExtra done checked=1");
            }

            if (Want("ce-rep-plain"))
            {
                RunReplace(
                    run, s, "ce-rep-plain",
                    "普通句。\r", false, 0,
                    "1. 手打",
                    "numbered=False",
                    "fallback WriteTextWithBreaks",
                    "ls=\"\"",
                    "wantPrefix=False",
                    "remove=False",
                    "ClearExtra done checked=1");
            }

            if (Want("ce-rep-two"))
            {
                RunReplace(
                    run, s, "ce-rep-two",
                    "第一条。\r", true, 0,
                    "2. 甲\r乙",
                    "numbered=True",
                    "ClearExtra rp=1",
                    "wantPrefix=True",
                    "ClearExtra rp=2",
                    "wantPrefix=False",
                    "remove=True",
                    "ClearExtra done checked=2");
            }

            if (Want("ce-rep-mid"))
            {
                RunReplace(
                    run, s, "ce-rep-mid",
                    "前句。后句。\r", true, 1,
                    "新后句",
                    "numbered=True",
                    "TryFindReplace before=",
                    "ClearExtra done checked=0");
            }

            if (Want("ce-ins-nocr"))
            {
                RunInsert(
                    run, s, "ce-ins-nocr",
                    "前句。后句。\r", false, 0,
                    "插入字",
                    "WriteInsert ",
                    "ClearExtra");
            }

            if (Want("ce-ins-mid-cr"))
            {
                RunInsert(
                    run, s, "ce-ins-mid-cr",
                    "前句。后句。\r", true, 0,
                    "甲\r乙",
                    "WriteInsert ",
                    "ClearExtra rp=2",
                    "seg=甲",
                    "wantPrefix=False",
                    "remove=True",
                    "ClearExtra done checked=1");
            }

            if (Want("ce-no-quit"))
            {
                bool wordOk = OfficeAlive.NotQuit("WINWORD", wordBefore, wordWinBefore, out string wordDetail);
                bool wpsOk = OfficeAlive.NotQuit("wps", wpsBefore, wpsWinBefore, out string wpsDetail);
                Ok(run, "ce-no-quit", wordOk && wpsOk, wordDetail + " " + wpsDetail);
            }
        }

        private static void RunReplace(
            TestRun run,
            WdSession s,
            string id,
            string plain,
            bool numbered,
            int codeIndex,
            string neu,
            params string[] needles)
        {
            if (!TryPlant(run, s, plain, numbered, out string path, out string idChan, out string text, id))
            {
                return;
            }

            string code = CodeAt(text, codeIndex);
            if (string.IsNullOrEmpty(code))
            {
                run.CaseFail(id, "没有第 " + codeIndex + " 个 S_。正文=" + Snip(text));
                s.CloseDocByPath(path);
                return;
            }

            DateTime since = DateTime.Now.AddSeconds(-2);
            ToolResult wr = s.ProcessActions(idChan, Replace(code, neu));
            if (!wr.Success)
            {
                run.CaseFail(id, wr.Error ?? "替换失败");
                s.CloseDocByPath(path);
                return;
            }

            string log = ReadActionLog(since, neu);
            Expect(run, id, log, needles, forbidRp: id == "ce-rep-mid");
            s.CloseDocByPath(path);
        }

        private static void RunInsert(
            TestRun run,
            WdSession s,
            string id,
            string plain,
            bool numbered,
            int codeIndex,
            string insert,
            params string[] rules)
        {
            if (!TryPlant(run, s, plain, numbered, out string path, out string idChan, out string text, id))
            {
                return;
            }

            string code = CodeAt(text, codeIndex);
            if (string.IsNullOrEmpty(code))
            {
                run.CaseFail(id, "没有第 " + codeIndex + " 个 S_。正文=" + Snip(text));
                s.CloseDocByPath(path);
                return;
            }

            DateTime since = DateTime.Now.AddSeconds(-2);
            ToolResult wr = s.ProcessActions(idChan, InsertAfter(code, insert));
            if (!wr.Success)
            {
                run.CaseFail(id, wr.Error ?? "插入失败");
                s.CloseDocByPath(path);
                return;
            }

            string log = ReadActionLog(since, insert);
            if (id == "ce-ins-nocr")
            {
                ExpectAbsent(run, id, log, rules);
            }
            else
            {
                Expect(run, id, log, rules, forbidRp: false);
            }

            s.CloseDocByPath(path);
        }

        private static void Expect(TestRun run, string id, string log, string[] needles, bool forbidRp)
        {
            if (string.IsNullOrEmpty(log))
            {
                run.CaseFail(id, "没有读到 process_document_actions 日志");
                return;
            }

            if (log.IndexOf("lists=", StringComparison.Ordinal) >= 0)
            {
                run.CaseFail(id, "热路径仍在打 lists= 全文扫");
                return;
            }

            if (forbidRp && log.IndexOf("ClearExtra rp=", StringComparison.Ordinal) >= 0)
            {
                run.CaseFail(id, "段中替换不该检查段落，却出现了 ClearExtra rp=");
                return;
            }

            foreach (string n in needles)
            {
                if (log.IndexOf(n, StringComparison.Ordinal) < 0)
                {
                    run.CaseFail(id, "日志缺「" + n + "」 " + AutonumSlice(log));
                    return;
                }
            }

            if (id == "ce-rep-plain" && log.IndexOf("TryFindReplace before=", StringComparison.Ordinal) >= 0)
            {
                run.CaseFail(id, "无编号却走了 TryFindReplace");
                return;
            }

            if (id == "ce-ins-mid-cr" && Count(log, "ClearExtra rp=") != 1)
            {
                run.CaseFail(id, "段中插入应只检查 1 个新段，rp 行数=" + Count(log, "ClearExtra rp="));
                return;
            }

            run.CaseOk(id);
        }

        private static void ExpectAbsent(TestRun run, string id, string log, string[] rules)
        {
            if (string.IsNullOrEmpty(log))
            {
                run.CaseFail(id, "没有读到 process_document_actions 日志");
                return;
            }

            if (log.IndexOf(rules[0], StringComparison.Ordinal) < 0)
            {
                run.CaseFail(id, "日志缺「" + rules[0] + "」 " + AutonumSlice(log));
                return;
            }

            if (log.IndexOf(rules[1], StringComparison.Ordinal) >= 0)
            {
                run.CaseFail(id, "无换行插入不该出现 " + rules[1]);
                return;
            }

            if (log.IndexOf("lists=", StringComparison.Ordinal) >= 0)
            {
                run.CaseFail(id, "热路径仍在打 lists=");
                return;
            }

            run.CaseOk(id);
        }

        private static bool TryPlant(
            TestRun run,
            WdSession s,
            string plain,
            bool numbered,
            out string path,
            out string channelId,
            out string text,
            string id)
        {
            path = s.NewDocxPath();
            channelId = null;
            text = "";
            ToolResult opened = s.Open(path, true, s.App);
            if (!opened.Success)
            {
                run.CaseFail(id, opened.Error ?? "新建失败");
                return false;
            }

            channelId = WdSession.Str(opened, "channel_id");
            if (!s.TypePlain(path, plain))
            {
                run.CaseFail(id, "COM 写句失败");
                s.CloseDocByPath(path);
                return false;
            }

            if (numbered && !s.ApplyNumberDefault(path, 1, 1))
            {
                run.CaseFail(id, "COM 套自动编号失败");
                s.CloseDocByPath(path);
                return false;
            }

            ToolResult got = s.GetContent(channelId);
            if (!got.Success)
            {
                run.CaseFail(id, got.Error ?? "读失败");
                s.CloseDocByPath(path);
                return false;
            }

            text = WdSession.Str(got, "document_content");
            if (string.IsNullOrEmpty(CodeAt(text, 0)))
            {
                run.CaseFail(id, "无可抄 S_");
                s.CloseDocByPath(path);
                return false;
            }

            return true;
        }

        private static string ReadActionLog(DateTime since, string token)
        {
            string escaped = (token ?? "").Replace("\r", "\\r");
            for (int attempt = 0; attempt < 8; attempt++)
            {
                string found = FindActionLog(since, token, escaped);
                if (!string.IsNullOrEmpty(found))
                {
                    return found;
                }

                System.Threading.Thread.Sleep(250);
            }

            return "";
        }

        private static string FindActionLog(DateTime since, string token, string escaped)
        {
            string found = "";
            foreach (string dir in LogDirs())
            {
                if (!Directory.Exists(dir))
                {
                    continue;
                }

                foreach (string file in Directory.GetFiles(dir, "process_document_actions_*.log"))
                {
                    DateTime stamp;
                    try
                    {
                        stamp = File.GetLastWriteTime(file);
                    }
                    catch
                    {
                        continue;
                    }

                    if (stamp < since)
                    {
                        continue;
                    }

                    string body;
                    try
                    {
                        using (var stream = new FileStream(file, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                        using (var reader = new StreamReader(stream))
                        {
                            body = reader.ReadToEnd();
                        }
                    }
                    catch
                    {
                        continue;
                    }

                    if (body.IndexOf(token, StringComparison.Ordinal) < 0
                        && body.IndexOf(escaped, StringComparison.Ordinal) < 0)
                    {
                        continue;
                    }

                    if (body.Length >= found.Length)
                    {
                        found = body;
                    }
                }
            }

            return found;
        }

        private static IEnumerable<string> LogDirs()
        {
            var dirs = new List<string>();
            string dir = AppDomain.CurrentDomain.BaseDirectory;
            for (int i = 0; i < 8 && !string.IsNullOrEmpty(dir); i++)
            {
                string desktop = Path.Combine(dir, "EasyWriteClient", "Desktop");
                if (Directory.Exists(desktop))
                {
                    dirs.Add(Path.Combine(desktop, "bin", "logs"));
                    dirs.Add(Path.Combine(desktop, "bin", "Debug", "logs"));
                    dirs.Add(Path.Combine(desktop, "bin", "Debug", "net472", "logs"));
                    break;
                }

                dir = Path.GetDirectoryName(dir);
            }

            return dirs;
        }

        private static string CodeAt(string text, int index)
        {
            if (string.IsNullOrEmpty(text) || index < 0)
            {
                return "";
            }

            MatchCollection matches = Regex.Matches(text, @"\[S_([0-9a-zA-Z]{5})-start\]");
            if (index >= matches.Count)
            {
                return "";
            }

            return "S_" + matches[index].Groups[1].Value;
        }

        private static int Count(string log, string token)
        {
            int n = 0;
            int i = 0;
            while (i < log.Length)
            {
                int at = log.IndexOf(token, i, StringComparison.Ordinal);
                if (at < 0)
                {
                    return n;
                }

                n++;
                i = at + token.Length;
            }

            return n;
        }

        private static List<Dictionary<string, object>> Replace(string codes, string neu)
        {
            return One("replace", new Dictionary<string, object>
            {
                ["original"] = new Dictionary<string, object> { ["codes"] = codes },
                ["new"] = neu,
                ["format"] = Fmt()
            });
        }

        private static List<Dictionary<string, object>> InsertAfter(string codes, string insert)
        {
            return One("insert", new Dictionary<string, object>
            {
                ["target"] = new Dictionary<string, object> { ["codes"] = codes },
                ["insert"] = insert,
                ["format"] = Fmt()
            });
        }

        private static List<Dictionary<string, object>> One(string type, Dictionary<string, object> detail)
        {
            return new List<Dictionary<string, object>>
            {
                new Dictionary<string, object>
                {
                    ["type"] = type,
                    ["detail"] = detail
                }
            };
        }

        private static Dictionary<string, object> Fmt()
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

        private static bool Want(string id)
        {
            string f = IdFilter;
            if (string.IsNullOrEmpty(f)
                || string.Equals(f, "ce", StringComparison.OrdinalIgnoreCase)
                || string.Equals(f, "ce-", StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            return string.Equals(f, id, StringComparison.OrdinalIgnoreCase)
                || id.StartsWith(f, StringComparison.OrdinalIgnoreCase);
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

        private static string AutonumSlice(string log)
        {
            if (string.IsNullOrEmpty(log))
            {
                return "";
            }

            int at = log.IndexOf("[autonum]", StringComparison.Ordinal);
            string slice = at < 0 ? log : log.Substring(at);
            return Snip(slice);
        }

        private static string Snip(string text)
        {
            if (string.IsNullOrEmpty(text))
            {
                return "";
            }

            string s = text.Replace("\r", "\\r").Replace("\n", "\\n");
            return s.Length <= 240 ? s : s.Substring(0, 240);
        }
    }
}
