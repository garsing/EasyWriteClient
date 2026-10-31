using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using Newtonsoft.Json;
using WordAddIn1;

namespace WordDocumentRoundtripTest
{
    /// <summary>
    /// 写后带回句子 display，并允许部分成功。前缀 wy-。
    /// </summary>
    internal static class WyCases
    {
        public const int Target = 6;

        private const string SentA = "易写带回甲句。";
        private const string SentB = "易写带回乙句。";
        private const string SentC = "易写带回丙句。";
        private const string NewA = "易写带回甲已改。";
        private const string NewB = "易写带回乙已改。";

        public static void Run(TestRun run, string host)
        {
            var s = new WdSession(host);
            RunAllOk(run, s);
            RunPartialFormat(run, s);
            RunPartialMiss(run, s);
            RunCross(run, s);
            RunAllFail(run, s);
            RunIntra(run, s);
        }

        private static void RunAllOk(TestRun run, WdSession s)
        {
            string path, id, text;
            if (!TryPlant(run, s, SentA + SentB, out path, out id, out text, "wy-all-ok"))
            {
                return;
            }

            try
            {
                List<string> codes = ListCodes(text, "S");
                if (codes.Count < 2)
                {
                    run.CaseFail("wy-all-ok", "需要两句 S_，实际 " + codes.Count);
                    return;
                }

                var actions = new List<Dictionary<string, object>>();
                actions.AddRange(Replace(codes[0], NewA));
                actions.AddRange(Replace(codes[1], NewB));
                ToolResult wr = s.ProcessActions(id, actions);
                string display = WdSession.Str(wr, "document_content");
                string again = FirstCode(display, "S");
                bool shape = wr.Success
                    && WdSession.Str(wr, "batch_outcome") == "all_succeeded"
                    && Contains(WdSession.Str(wr, "message"), "全部成功")
                    && !string.IsNullOrEmpty(display)
                    && display.IndexOf(NewA, StringComparison.Ordinal) >= 0
                    && display.IndexOf(NewB, StringComparison.Ordinal) >= 0
                    && !string.IsNullOrEmpty(again);
                if (!shape)
                {
                    run.CaseFail("wy-all-ok", wr.Error ?? WdSession.Str(wr, "message") ?? "回包不是全部成功或没有 display");
                    return;
                }

                ToolResult second = s.ProcessActions(id, Replace(again, "易写带回再用回包码。"));
                Ok(run, "wy-all-ok",
                    second.Success && StillHas(s, id, "易写带回再用回包码。"),
                    second.Success ? "回包里的 S_ 没能再改上" : (second.Error ?? "用回包码再改失败"));
            }
            finally
            {
                s.CloseDocByPath(path);
            }
        }

        private static void RunPartialFormat(TestRun run, WdSession s)
        {
            string path, id, text;
            if (!TryPlant(run, s, SentA + SentB, out path, out id, out text, "wy-partial-fmt"))
            {
                return;
            }

            try
            {
                List<string> codes = ListCodes(text, "S");
                if (codes.Count < 2)
                {
                    run.CaseFail("wy-partial-fmt", "需要两句 S_");
                    return;
                }

                var actions = new List<Dictionary<string, object>>();
                actions.AddRange(Replace(codes[0], NewA));
                actions.AddRange(Replace(codes[1], NewB, format: null));
                ToolResult wr = s.ProcessActions(id, actions);
                Ok(run, "wy-partial-fmt",
                    !wr.Success
                    && WdSession.Str(wr, "batch_outcome") == "partial_succeeded"
                    && Contains(WdSession.Str(wr, "message"), "部分成功")
                    && Contains(WdSession.Str(wr, "message"), "不要重做")
                    && BlobHas(wr, "char_format")
                    && !string.IsNullOrEmpty(WdSession.Str(wr, "document_content"))
                    && StillHas(s, id, NewA)
                    && StillHas(s, id, SentB)
                    && !StillHas(s, id, NewB),
                    wr.Error ?? "部分成功未留稿或没指出缺格式");
            }
            finally
            {
                s.CloseDocByPath(path);
            }
        }

        private static void RunPartialMiss(TestRun run, WdSession s)
        {
            string path, id, text;
            if (!TryPlant(run, s, SentA + SentB, out path, out id, out text, "wy-partial-miss"))
            {
                return;
            }

            try
            {
                string code = FirstCode(text, "S");
                var actions = new List<Dictionary<string, object>>();
                actions.AddRange(Replace(code, NewA));
                actions.AddRange(Replace("S_000zz", NewB));
                ToolResult wr = s.ProcessActions(id, actions);
                string message = WdSession.Str(wr, "message");
                if (string.IsNullOrEmpty(message))
                {
                    message = wr.Error ?? "";
                }

                Ok(run, "wy-partial-miss",
                    !wr.Success
                    && WdSession.Str(wr, "batch_outcome") == "partial_succeeded"
                    && Contains(message, "当前定位列表里没有")
                    && Contains(message, "S_000zz")
                    && !Contains(message, "跨段")
                    && !Contains(message, "跨 seg")
                    && StillHas(s, id, NewA)
                    && !string.IsNullOrEmpty(WdSession.Str(wr, "document_content")),
                    message ?? "单颗坏码被写成跨段，或合法条没留下");
            }
            finally
            {
                s.CloseDocByPath(path);
            }
        }

        private static void RunCross(TestRun run, WdSession s)
        {
            string path, id, text;
            if (!TryPlant(run, s, SentA + SentB + SentC, out path, out id, out text, "wy-cross"))
            {
                return;
            }

            try
            {
                List<string> codes = ListCodes(text, "S");
                if (codes.Count < 3)
                {
                    run.CaseFail("wy-cross", "需要三句 S_，实际 " + codes.Count);
                    return;
                }

                var actions = new List<Dictionary<string, object>>();
                actions.AddRange(Replace(codes[0] + "," + codes[2], NewA));
                actions.AddRange(Replace(codes[1], NewB));
                ToolResult wr = s.ProcessActions(id, actions);
                string message = string.IsNullOrEmpty(WdSession.Str(wr, "message"))
                    ? (wr.Error ?? "")
                    : WdSession.Str(wr, "message");
                Ok(run, "wy-cross",
                    !wr.Success
                    && WdSession.Str(wr, "batch_outcome") == "partial_succeeded"
                    && (Contains(message, "不连续") || Contains(message, "跨段"))
                    && StillHas(s, id, NewB)
                    && StillHas(s, id, SentA)
                    && StillHas(s, id, SentC)
                    && !StillHas(s, id, NewA),
                    message ?? "不相邻两码没有按该条失败，或中间句没留下");
            }
            finally
            {
                s.CloseDocByPath(path);
            }
        }

        private static void RunAllFail(TestRun run, WdSession s)
        {
            string path, id, text;
            if (!TryPlant(run, s, SentA + SentB, out path, out id, out text, "wy-all-fail"))
            {
                return;
            }

            try
            {
                var actions = new List<Dictionary<string, object>>();
                actions.AddRange(Replace("S_000za", NewA));
                actions.AddRange(Replace("S_000zb", NewB));
                ToolResult wr = s.ProcessActions(id, actions);
                string message = string.IsNullOrEmpty(WdSession.Str(wr, "message"))
                    ? (wr.Error ?? "")
                    : WdSession.Str(wr, "message");
                Dictionary<string, object> data = WdSession.Data(wr);
                bool noDisplay = data == null || !data.ContainsKey("document_content");
                Ok(run, "wy-all-fail",
                    !wr.Success
                    && WdSession.Str(wr, "batch_outcome") == "all_failed"
                    && Contains(message, "全部失败")
                    && Contains(message, "S_000za")
                    && Contains(message, "S_000zb")
                    && noDisplay
                    && StillHas(s, id, SentA)
                    && StillHas(s, id, SentB),
                    message ?? "全部失败仍改了稿或只报了一颗码");
            }
            finally
            {
                s.CloseDocByPath(path);
            }
        }

        private static void RunIntra(TestRun run, WdSession s)
        {
            string path, id, text;
            if (!TryPlant(run, s, SentA + SentB, out path, out id, out text, "wy-intra"))
            {
                return;
            }

            try
            {
                List<string> codes = ListCodes(text, "S");
                if (codes.Count < 2)
                {
                    run.CaseFail("wy-intra", "需要两句 S_");
                    return;
                }

                var actions = new List<Dictionary<string, object>>();
                actions.AddRange(Replace(codes[0], SentA));
                actions.AddRange(Replace(codes[1], SentA));
                ToolResult wr = s.ProcessActions(id, actions);
                string message = string.IsNullOrEmpty(WdSession.Str(wr, "message"))
                    ? (wr.Error ?? "")
                    : WdSession.Str(wr, "message");
                Ok(run, "wy-intra",
                    !wr.Success
                    && WdSession.Str(wr, "batch_outcome") == "partial_succeeded"
                    && Contains(message, "不要重做")
                    && Contains(message, "当前定位列表里没有")
                    && !Contains(message, "跨段")
                    && !Contains(message, "跨 seg")
                    && StillHas(s, id, SentA)
                    && !string.IsNullOrEmpty(WdSession.Str(wr, "document_content")),
                    message ?? "同批后写的旧码没有留下先写上的那条");
            }
            finally
            {
                s.CloseDocByPath(path);
            }
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
            if (string.IsNullOrEmpty(FirstCode(text, "S")))
            {
                FailMany(run, "无可抄 S_", failIds);
                s.CloseDocByPath(path);
                return false;
            }

            return true;
        }

        private static List<Dictionary<string, object>> Replace(
            string codes,
            string neu,
            Dictionary<string, object> format)
        {
            var detail = new Dictionary<string, object>
            {
                ["original"] = new Dictionary<string, object> { ["codes"] = codes },
                ["new"] = neu
            };
            if (format != null)
            {
                detail["format"] = format;
            }

            return new List<Dictionary<string, object>>
            {
                new Dictionary<string, object>
                {
                    ["type"] = "replace",
                    ["detail"] = detail
                }
            };
        }

        private static List<Dictionary<string, object>> Replace(string codes, string neu)
        {
            return Replace(codes, neu, DefaultFormat());
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

        private static List<string> ListCodes(string text, string prefix)
        {
            var list = new List<string>();
            if (string.IsNullOrEmpty(text))
            {
                return list;
            }

            foreach (Match m in Regex.Matches(text, @"\[" + prefix + @"_([0-9a-zA-Z]{5})-start\]"))
            {
                list.Add(prefix + "_" + m.Groups[1].Value);
            }

            return list;
        }

        private static string FirstCode(string text, string prefix)
        {
            List<string> list = ListCodes(text, prefix);
            return list.Count == 0 ? "" : list[0];
        }

        private static string ReadBody(WdSession s, string id)
        {
            ToolResult got = s.GetContent(id);
            return got.Success ? WdSession.Str(got, "document_content") : "";
        }

        private static bool StillHas(WdSession s, string id, string needle)
        {
            return ReadBody(s, id).IndexOf(needle, StringComparison.Ordinal) >= 0;
        }

        private static bool BlobHas(ToolResult r, string needle)
        {
            if (Contains(r == null ? null : r.Error, needle))
            {
                return true;
            }

            if (r == null || r.Data == null)
            {
                return false;
            }

            try
            {
                return Contains(JsonConvert.SerializeObject(r.Data), needle);
            }
            catch
            {
                return Contains(Convert.ToString(r.Data), needle);
            }
        }

        private static bool Contains(string text, string needle)
        {
            return !string.IsNullOrEmpty(text)
                && text.IndexOf(needle, StringComparison.OrdinalIgnoreCase) >= 0;
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
