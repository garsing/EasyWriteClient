using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using Newtonsoft.Json;
using WordAddIn1;

namespace WordDocumentRoundtripTest
{
    /// <summary>
    /// 写入套格式切块：对 process_actions 会话日志行 <c>[写入套格式]</c> 断言前/中/后缀套法。
    /// </summary>
    internal static class WlCases
    {
        public const int Target = 8;

        public static void Run(
            TestRun run,
            string host,
            int wordBefore,
            int wpsBefore,
            bool wordWinBefore,
            bool wpsWinBefore)
        {
            var s = new WdSession(host);
            string warmPath = WarmWpsChannel(s);
            try
            {
                RunSafeMid(run, s);
                RunGlue(run, s);
                RunSuffix(run, s);
                RunNoCr(run, s);
                RunLayout(run, s);
                RunRuns(run, s);
                RunEcho(run, s);
                RunNoQuit(run, wordBefore, wpsBefore, wordWinBefore, wpsWinBefore);
            }
            finally
            {
                if (!string.IsNullOrEmpty(warmPath))
                {
                    s.CloseDocByPath(warmPath);
                }
            }
        }

        private static string WarmWpsChannel(WdSession s)
        {
            if (!s.IsWps)
            {
                return null;
            }

            string path = s.NewDocxPath();
            ToolResult opened = s.Open(path, true, s.App);
            if (!opened.Success)
            {
                return null;
            }

            string id = WdSession.Str(opened, "channel_id");
            s.ActivateDoc(path);
            s.GetContent(id);
            return path;
        }

        private static void RunSafeMid(TestRun run, WdSession s)
        {
            if (!TryPlant(run, s, "锚句。\r", out string path, out string id, out string text, "wl-safe-mid"))
            {
                return;
            }

            ToolResult wr = s.ProcessActions(id, InsertAfter(FirstCode(text), "世界。\r第二段。\r"));
            ExpectPlan(run, "wl-safe-mid", wr,
                "prefixSafe=true", "mapped=true", "prefix=join", "mid=once", "suffix=skip", "reason=map");
            s.CloseDocByPath(path);
        }

        private static void RunGlue(TestRun run, WdSession s)
        {
            if (!TryPlant(run, s, "你好。还好。", out string path, out string id, out string text, "wl-glue"))
            {
                return;
            }

            ToolResult wr = s.ProcessActions(id, InsertAfter(FirstCode(text), "世界。\r第二段。\r"));
            ExpectPlan(run, "wl-glue", wr,
                "prefixSafe=false", "mapped=true", "prefix=sentences", "mid=once", "suffix=skip", "reason=map");
            s.CloseDocByPath(path);
        }

        private static void RunSuffix(TestRun run, WdSession s)
        {
            if (!TryPlant(run, s, "锚句。\r", out string path, out string id, out string text, "wl-suffix"))
            {
                return;
            }

            ToolResult wr = s.ProcessActions(id, InsertAfter(FirstCode(text), "A\rB\rC"));
            ExpectPlan(run, "wl-suffix", wr,
                "mapped=true", "prefix=join", "mid=once", "suffix=sentences", "reason=map");
            s.CloseDocByPath(path);
        }

        private static void RunNoCr(TestRun run, WdSession s)
        {
            if (!TryPlant(run, s, "锚句。\r", out string path, out string id, out string text, "wl-nocr"))
            {
                return;
            }

            ToolResult wr = s.ProcessActions(id, InsertAfter(FirstCode(text), "单独一句。"));
            ExpectPlan(run, "wl-nocr", wr,
                "mapped=false", "prefix=skip", "mid=skip", "suffix=sentences", "reason=no-cr");
            s.CloseDocByPath(path);
        }

        private static void RunLayout(TestRun run, WdSession s)
        {
            if (!TryPlant(run, s, "锚句。\r", out string path, out string id, out string text, "wl-layout"))
            {
                return;
            }

            ToolResult wr = s.ProcessActions(id, InsertAfter(FirstCode(text), "甲。\f乙。"));
            ExpectPlan(run, "wl-layout", wr,
                "mapped=false", "prefix=skip", "mid=skip", "suffix=sentences", "reason=layout");
            s.CloseDocByPath(path);
        }

        private static void RunRuns(TestRun run, WdSession s)
        {
            if (!TryPlant(run, s, "你好。还好。", out string path, out string id, out string text, "wl-runs"))
            {
                return;
            }

            var runs = new List<object>
            {
                Run("标题", false),
                Run("下段。\r", false)
            };
            var detail = new Dictionary<string, object>
            {
                ["target"] = Locator(FirstCode(text)),
                ["insert"] = runs
            };
            ToolResult wr = s.ProcessActions(id, One("insert", detail));
            string blob = Blob(wr);
            int tags = CountTag(blob);
            Ok(run, "wl-runs",
                wr != null && wr.Success && tags >= 2 && blob.IndexOf("prefixSafe=false", StringComparison.Ordinal) >= 0,
                wr == null || !wr.Success
                    ? (wr == null ? "null" : wr.Error)
                    : ("tag=" + tags + " blob=" + TrimBlob(blob)));
            s.CloseDocByPath(path);
        }

        private static void RunEcho(TestRun run, WdSession s)
        {
            if (!TryPlant(run, s, "锚句。\r", out string path, out string id, out string text, "wl-echo"))
            {
                return;
            }

            ToolResult wr = s.ProcessActions(id, InsertAfter(FirstCode(text), "单独一句。"));
            bool hasKey = WdSession.HasKey(wr, "format_apply_log");
            bool inBlob = Blob(wr).IndexOf(WrittenFormatGlueHelper.LogTag, StringComparison.Ordinal) >= 0;
            Ok(run, "wl-echo", wr != null && wr.Success && hasKey && inBlob,
                wr == null || !wr.Success
                    ? (wr == null ? "null" : wr.Error)
                    : ("hasKey=" + hasKey + " inBlob=" + inBlob));
            s.CloseDocByPath(path);
        }

        private static void RunNoQuit(
            TestRun run,
            int wordBefore,
            int wpsBefore,
            bool wordWinBefore,
            bool wpsWinBefore)
        {
            string wordDetail;
            string wpsDetail;
            bool wordOk = OfficeAlive.NotQuit("WINWORD", wordBefore, wordWinBefore, out wordDetail);
            bool wpsOk = OfficeAlive.NotQuit("wps", wpsBefore, wpsWinBefore, out wpsDetail);
            Ok(run, "wl-no-quit", wordOk && wpsOk, wordDetail + " " + wpsDetail);
        }

        private static void ExpectPlan(TestRun run, string id, ToolResult wr, params string[] needles)
        {
            if (wr == null || !wr.Success)
            {
                run.CaseFail(id, wr == null ? "null" : (wr.Error ?? "失败"));
                return;
            }

            string blob = Blob(wr);
            if (blob.IndexOf(WrittenFormatGlueHelper.LogTag, StringComparison.Ordinal) < 0)
            {
                run.CaseFail(id, "回包无 " + WrittenFormatGlueHelper.LogTag + " blob=" + TrimBlob(blob));
                return;
            }

            foreach (string n in needles)
            {
                if (blob.IndexOf(n, StringComparison.Ordinal) < 0)
                {
                    run.CaseFail(id, "日志缺「" + n + "」 blob=" + TrimBlob(blob));
                    return;
                }
            }

            run.CaseOk(id);
        }

        private static string Blob(ToolResult r)
        {
            if (r == null || r.Data == null)
            {
                return r == null ? "" : (r.Error ?? "");
            }

            try
            {
                return JsonConvert.SerializeObject(r.Data) + " " + (r.Error ?? "");
            }
            catch
            {
                return Convert.ToString(r.Data) ?? "";
            }
        }

        private static int CountTag(string blob)
        {
            if (string.IsNullOrEmpty(blob))
            {
                return 0;
            }

            int n = 0;
            int i = 0;
            while (true)
            {
                int at = blob.IndexOf(WrittenFormatGlueHelper.LogTag, i, StringComparison.Ordinal);
                if (at < 0)
                {
                    return n;
                }

                n++;
                i = at + WrittenFormatGlueHelper.LogTag.Length;
            }
        }

        private static string TrimBlob(string blob)
        {
            if (string.IsNullOrEmpty(blob))
            {
                return "";
            }

            return blob.Length <= 400 ? blob : blob.Substring(0, 400);
        }

        private static bool TryPlant(
            TestRun run,
            WdSession s,
            string plain,
            out string path,
            out string id,
            out string text,
            params string[] failIds)
        {
            path = s.NewDocxPath();
            id = null;
            text = "";
            ToolResult opened = s.Open(path, true, s.App);
            if (!opened.Success)
            {
                FailMany(run, opened.Error ?? "新建失败", failIds);
                return false;
            }

            id = WdSession.Str(opened, "channel_id");
            if (!s.TypePlain(path, plain))
            {
                FailMany(run, "COM 写句失败", failIds);
                s.CloseDocByPath(path);
                return false;
            }

            ToolResult got = s.GetContent(id);
            if (!got.Success)
            {
                FailMany(run, got.Error ?? "读失败", failIds);
                s.CloseDocByPath(path);
                return false;
            }

            text = WdSession.Str(got, "document_content");
            if (string.IsNullOrEmpty(FirstCode(text)))
            {
                FailMany(run, "无可抄 S_", failIds);
                s.CloseDocByPath(path);
                return false;
            }

            return true;
        }

        private static List<Dictionary<string, object>> InsertAfter(string codes, string insert)
        {
            return One("insert", new Dictionary<string, object>
            {
                ["target"] = Locator(codes),
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

        private static Dictionary<string, object> Locator(string codes)
        {
            return new Dictionary<string, object> { ["codes"] = codes };
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

        private static Dictionary<string, object> Run(string text, bool bold)
        {
            return new Dictionary<string, object>
            {
                ["text"] = text,
                ["format"] = new Dictionary<string, object>
                {
                    ["char_format"] = new Dictionary<string, object>
                    {
                        ["font_name"] = "宋体",
                        ["font_size"] = 12,
                        ["is_bold"] = bold
                    }
                }
            };
        }

        private static string FirstCode(string text)
        {
            if (string.IsNullOrEmpty(text))
            {
                return "";
            }

            Match m = Regex.Match(text, @"\[S_([0-9a-zA-Z]{5})-start\]");
            return m.Success ? "S_" + m.Groups[1].Value : "";
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
