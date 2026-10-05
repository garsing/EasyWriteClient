using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using WordAddIn1;

namespace WordDocumentRoundtripTest
{
    internal static class WcCases
    {
        public const int Target = 42;

        private const string SentA = "易写读编码甲句。";
        private const string SentB = "易写读编码乙句。";
        private const string ParaA = "段甲夹具。";
        private const string ParaB = "段乙夹具。";
        private const string Dup = "重复句夹具。";

        public static void Run(
            TestRun run,
            string host,
            int wordBefore,
            int wpsBefore,
            bool wordWinBefore,
            bool wpsWinBefore)
        {
            var s = new WdSession(host);
            RunNegEarly(run, s);
            RunEmpty(run, s);
            string textPath = RunSentence(run, s);
            RunParagraph(run, s);
            RunDuplicate(run, s);
            RunTable(run, s);
            RunImage(run, s);
            RunChart(run, s);
            RunContract(run, s, textPath);
            RunNegLate(run, s);
            RunNoQuit(run, wordBefore, wpsBefore, wordWinBefore, wpsWinBefore);
        }

        private static void RunNegEarly(TestRun run, WdSession s)
        {
            ToolResult r = s.GetContent(null);
            if (!r.Success && Contains(r.Error, "打开"))
            {
                run.CaseOk("wc-neg-no-channel");
            }
            else if (r.Success)
            {
                run.CaseSkip("wc-neg-no-channel", "现场已有默认文字渠，不关用户稿");
            }
            else
            {
                run.CaseFail("wc-neg-no-channel", r.Error ?? "应失败或跳过");
            }

            ExpectFail(run, "wc-neg-unknown-id", s.GetContent("w999"), "未知");
        }

        private static void RunEmpty(TestRun run, WdSession s)
        {
            string path = s.NewDocxPath();
            ToolResult opened = s.Open(path, true, s.App);
            if (!opened.Success)
            {
                FailMany(run, opened.Error ?? "新建失败",
                    "wc-empty-ok", "wc-empty-no-body", "wc-empty-note",
                    "wc-empty-level", "wc-empty-host-id", "wc-neg-bad-level");
                return;
            }

            string id = WdSession.Str(opened, "channel_id");
            ExpectFail(run, "wc-neg-bad-level", s.GetContent(id, "word"), "sentence");

            ToolResult got = s.GetContent(id);
            if (!got.Success)
            {
                FailMany(run, got.Error ?? "空稿读失败",
                    "wc-empty-ok", "wc-empty-no-body", "wc-empty-note",
                    "wc-empty-level", "wc-empty-host-id");
                s.CloseDocByPath(path);
                return;
            }

            string body = StripCodes(WdSession.Str(got, "document_content"));
            Ok(run, "wc-empty-ok", true, "");
            Ok(run, "wc-empty-no-body",
                body.IndexOf("易写", StringComparison.Ordinal) < 0
                && body.IndexOf("夹具", StringComparison.Ordinal) < 0
                && body.Trim().Length < 16,
                "空稿仍有正文「" + TrimPreview(body) + "」");
            Ok(run, "wc-empty-note",
                Contains(WdSession.Str(got, "note"), "S_00000")
                && (Contains(WdSession.Str(got, "note"), "勿锚")
                    || Contains(WdSession.Str(got, "note"), "几乎为空")),
                WdSession.Str(got, "note"));
            Ok(run, "wc-empty-level",
                string.Equals(WdSession.Str(got, "code_level"), "sentence", StringComparison.Ordinal),
                "code_level=" + WdSession.Str(got, "code_level"));
            Ok(run, "wc-empty-host-id",
                string.Equals(WdSession.Str(got, "host"), s.ExpectedKind, StringComparison.OrdinalIgnoreCase)
                && WdSession.IsShortId(WdSession.Str(got, "channel_id")),
                "host/id " + WdSession.Str(got, "host") + " " + WdSession.Str(got, "channel_id"));
            s.CloseDocByPath(path);
        }

        private static string RunSentence(TestRun run, WdSession s)
        {
            string path = s.NewDocxPath();
            ToolResult opened = s.Open(path, true, s.App);
            if (!opened.Success)
            {
                FailMany(run, opened.Error ?? "新建失败",
                    "wc-s-default", "wc-s-explicit", "wc-s-alias", "wc-s-markers",
                    "wc-s-text", "wc-s-codes", "wc-s-default-ch", "wc-s-explicit-id");
                return null;
            }

            string id = WdSession.Str(opened, "channel_id");
            if (!s.TypePlain(path, SentA + SentB))
            {
                FailMany(run, "COM 写句失败",
                    "wc-s-default", "wc-s-explicit", "wc-s-alias", "wc-s-markers",
                    "wc-s-text", "wc-s-codes", "wc-s-default-ch", "wc-s-explicit-id");
                return path;
            }

            ToolResult def = s.GetContent(id);
            if (!def.Success)
            {
                FailMany(run, def.Error ?? "句级读失败",
                    "wc-s-default", "wc-s-explicit", "wc-s-alias", "wc-s-markers",
                    "wc-s-text", "wc-s-codes", "wc-s-default-ch", "wc-s-explicit-id");
                return path;
            }

            string text = WdSession.Str(def, "document_content");
            Ok(run, "wc-s-default",
                string.Equals(WdSession.Str(def, "code_level"), "sentence", StringComparison.Ordinal),
                "code_level=" + WdSession.Str(def, "code_level"));

            ToolResult expl = s.GetContent(id, "sentence");
            Ok(run, "wc-s-explicit",
                expl.Success
                && string.Equals(WdSession.Str(expl, "code_level"), "sentence", StringComparison.Ordinal),
                expl.Success ? "" : (expl.Error ?? "explicit 失败"));

            ToolResult alias = s.GetContent(id, "s");
            Ok(run, "wc-s-alias",
                alias.Success
                && string.Equals(WdSession.Str(alias, "code_level"), "sentence", StringComparison.Ordinal),
                alias.Success ? "" : (alias.Error ?? "alias 失败"));

            Ok(run, "wc-s-markers",
                text.IndexOf("[S_", StringComparison.Ordinal) >= 0
                && text.IndexOf("-start]", StringComparison.Ordinal) >= 0
                && text.IndexOf("-end]", StringComparison.Ordinal) >= 0,
                "无 S_ 起止标记");
            Ok(run, "wc-s-text", text.IndexOf("易写读编码甲句", StringComparison.Ordinal) >= 0, "夹具未出现");
            Ok(run, "wc-s-codes", CountCodes(text, "S") >= 1, "无可抄 S_");

            ToolResult noId = s.GetContent(null);
            Ok(run, "wc-s-default-ch",
                noId.Success && string.Equals(WdSession.Str(noId, "channel_id"), id, StringComparison.Ordinal),
                noId.Success ? "默认渠 " + WdSession.Str(noId, "channel_id") : (noId.Error ?? "默认失败"));
            Ok(run, "wc-s-explicit-id",
                string.Equals(WdSession.Str(def, "channel_id"), id, StringComparison.Ordinal),
                "回包 id=" + WdSession.Str(def, "channel_id"));
            return path;
        }

        private static void RunParagraph(TestRun run, WdSession s)
        {
            string path = s.NewDocxPath();
            ToolResult opened = s.Open(path, true, s.App);
            if (!opened.Success)
            {
                FailMany(run, opened.Error ?? "新建失败",
                    "wc-p-level", "wc-p-alias", "wc-p-markers", "wc-p-two", "wc-p-note");
                return;
            }

            string id = WdSession.Str(opened, "channel_id");
            if (!s.TypePlain(path, ParaA + "\r" + ParaB))
            {
                FailMany(run, "COM 写段失败",
                    "wc-p-level", "wc-p-alias", "wc-p-markers", "wc-p-two", "wc-p-note");
                s.CloseDocByPath(path);
                return;
            }

            ToolResult got = s.GetContent(id, "paragraph");
            if (!got.Success)
            {
                FailMany(run, got.Error ?? "段级读失败",
                    "wc-p-level", "wc-p-alias", "wc-p-markers", "wc-p-two", "wc-p-note");
                s.CloseDocByPath(path);
                return;
            }

            string text = WdSession.Str(got, "document_content");
            Ok(run, "wc-p-level",
                string.Equals(WdSession.Str(got, "code_level"), "paragraph", StringComparison.Ordinal),
                "code_level=" + WdSession.Str(got, "code_level"));

            ToolResult alias = s.GetContent(id, "p_");
            Ok(run, "wc-p-alias",
                alias.Success
                && string.Equals(WdSession.Str(alias, "code_level"), "paragraph", StringComparison.Ordinal),
                alias.Success ? "" : (alias.Error ?? "alias 失败"));
            Ok(run, "wc-p-markers",
                text.IndexOf("[P_", StringComparison.Ordinal) >= 0
                && text.IndexOf("-start]", StringComparison.Ordinal) >= 0,
                "无 P_ 标记");
            Ok(run, "wc-p-two", CountCodes(text, "P") >= 2, "P_ 数=" + CountCodes(text, "P"));
            Ok(run, "wc-p-note",
                Contains(WdSession.Str(got, "note"), "P_")
                || Contains(WdSession.Str(got, "note"), "段落"),
                WdSession.Str(got, "note"));
            s.CloseDocByPath(path);
        }

        private static void RunDuplicate(TestRun run, WdSession s)
        {
            string path = s.NewDocxPath();
            ToolResult opened = s.Open(path, true, s.App);
            if (!opened.Success)
            {
                FailMany(run, opened.Error ?? "新建失败", "wc-dup-s", "wc-dup-p");
                return;
            }

            string id = WdSession.Str(opened, "channel_id");
            if (!s.TypePlain(path, Dup + "\r" + Dup))
            {
                FailMany(run, "COM 写重复句失败", "wc-dup-s", "wc-dup-p");
                s.CloseDocByPath(path);
                return;
            }

            ToolResult sent = s.GetContent(id);
            Ok(run, "wc-dup-s",
                sent.Success && Contains(WdSession.Str(sent, "note"), "出现多次"),
                sent.Success ? WdSession.Str(sent, "note") : (sent.Error ?? "读失败"));

            ToolResult para = s.GetContent(id, "paragraph");
            Ok(run, "wc-dup-p",
                para.Success && Contains(WdSession.Str(para, "note"), "出现多次"),
                para.Success ? WdSession.Str(para, "note") : (para.Error ?? "读失败"));
            s.CloseDocByPath(path);
        }

        private static void RunTable(TestRun run, WdSession s)
        {
            string path = s.NewDocxPath();
            ToolResult opened = s.Open(path, true, s.App);
            if (!opened.Success)
            {
                FailMany(run, opened.Error ?? "新建失败",
                    "wc-t-tag", "wc-t-tr-td", "wc-t-id", "wc-t-no-slots");
                return;
            }

            string id = WdSession.Str(opened, "channel_id");
            if (!s.InsertTable(path, 2, 2, new[] { "甲", "乙", "丙", "丁" }))
            {
                FailMany(run, "COM 建表失败",
                    "wc-t-tag", "wc-t-tr-td", "wc-t-id", "wc-t-no-slots");
                s.CloseDocByPath(path);
                return;
            }

            ToolResult got = s.GetContent(id);
            if (!got.Success)
            {
                FailMany(run, got.Error ?? "读表失败",
                    "wc-t-tag", "wc-t-tr-td", "wc-t-id", "wc-t-no-slots");
                s.CloseDocByPath(path);
                return;
            }

            string text = WdSession.Str(got, "document_content");
            Ok(run, "wc-t-tag",
                text.IndexOf("<table id=\"T_", StringComparison.Ordinal) >= 0, "无 table/T_");
            Ok(run, "wc-t-tr-td",
                Regex.IsMatch(text, @"<tr\b", RegexOptions.IgnoreCase)
                && Regex.IsMatch(text, @"<td\b", RegexOptions.IgnoreCase),
                "无 tr/td");
            Ok(run, "wc-t-id",
                Regex.IsMatch(text, @"table id=""T_[^""]+"""),
                "T_ 不可抄");
            Ok(run, "wc-t-no-slots",
                !WdSession.HasKey(got, "empty_count") && !WdSession.HasKey(got, "slots"),
                "Data 不该带填空槽");
            s.CloseDocByPath(path);
        }

        private static void RunImage(TestRun run, WdSession s)
        {
            string path = s.NewDocxPath();
            ToolResult opened = s.Open(path, true, s.App);
            if (!opened.Success)
            {
                FailMany(run, opened.Error ?? "新建失败",
                    "wc-i-tag", "wc-i-selfclose", "wc-i-id");
                return;
            }

            string id = WdSession.Str(opened, "channel_id");
            if (!s.InsertPicture(path))
            {
                FailMany(run, "COM 插图失败",
                    "wc-i-tag", "wc-i-selfclose", "wc-i-id");
                s.CloseDocByPath(path);
                return;
            }

            ToolResult got = s.GetContent(id);
            if (!got.Success)
            {
                FailMany(run, got.Error ?? "读图失败",
                    "wc-i-tag", "wc-i-selfclose", "wc-i-id");
                s.CloseDocByPath(path);
                return;
            }

            string text = WdSession.Str(got, "document_content");
            Ok(run, "wc-i-tag",
                text.IndexOf("<image", StringComparison.OrdinalIgnoreCase) >= 0
                && text.IndexOf("id=\"I_", StringComparison.Ordinal) >= 0,
                "无 image/I_");
            Ok(run, "wc-i-selfclose",
                Regex.IsMatch(text, @"<image\b[^>]*/>"),
                "image 未自闭合");
            Ok(run, "wc-i-id",
                Regex.IsMatch(text, @"id=""I_[^""]+"""),
                "I_ 不可抄");
            s.CloseDocByPath(path);
        }

        private static void RunChart(TestRun run, WdSession s)
        {
            string path;
            ToolResult opened;
            if (s.IsWps)
            {
                path = s.CreateChartFixtureDocx();
                if (string.IsNullOrEmpty(path))
                {
                    FailMany(run, "无法用 Word 做出 Chart 夹具（未 SKIP）",
                        "wc-c-tag", "wc-c-id", "wc-c-not-only-image");
                    return;
                }

                opened = s.Open(path, false, s.App);
            }
            else
            {
                path = s.NewDocxPath();
                opened = s.Open(path, true, s.App);
            }

            if (!opened.Success)
            {
                FailMany(run, opened.Error ?? "打开 Chart 稿失败",
                    "wc-c-tag", "wc-c-id", "wc-c-not-only-image");
                return;
            }

            string id = WdSession.Str(opened, "channel_id");
            if (!s.IsWps && !s.InsertChart(path))
            {
                FailMany(run, "COM 插 Chart 失败（未 SKIP）",
                    "wc-c-tag", "wc-c-id", "wc-c-not-only-image");
                s.CloseDocByPath(path);
                return;
            }

            ToolResult got = s.GetContent(id);
            if (!got.Success)
            {
                FailMany(run, got.Error ?? "读 Chart 失败",
                    "wc-c-tag", "wc-c-id", "wc-c-not-only-image");
                s.CloseDocByPath(path);
                return;
            }

            string text = WdSession.Str(got, "document_content");
            string preview = TrimPreview(text);
            Ok(run, "wc-c-tag",
                Regex.IsMatch(text, @"<Chart\s+id=""C_[^""]+""", RegexOptions.IgnoreCase),
                "无 Chart/C_ 「" + preview + "」");
            Ok(run, "wc-c-id",
                Regex.IsMatch(text, @"id=""C_[^""]+"""),
                "C_ 不可抄 「" + preview + "」");
            Ok(run, "wc-c-not-only-image",
                Regex.IsMatch(text, @"<Chart\b", RegexOptions.IgnoreCase),
                "只有图没有 Chart 标签 「" + preview + "」");
            s.CloseDocByPath(path);
        }

        private static void RunContract(TestRun run, WdSession s, string textPath)
        {
            if (string.IsNullOrEmpty(textPath))
            {
                FailMany(run, "依赖句级打开",
                    "wc-fields", "wc-local", "wc-no-snapshot-key", "wc-id-short", "wc-host");
                return;
            }

            ToolResult opened = s.Open(textPath, false, s.App);
            string id = WdSession.Str(opened, "channel_id");
            ToolResult got = opened.Success ? s.GetContent(id) : opened;
            if (!got.Success)
            {
                FailMany(run, got.Error ?? "契约读失败",
                    "wc-fields", "wc-local", "wc-no-snapshot-key", "wc-id-short", "wc-host");
                return;
            }

            Ok(run, "wc-fields",
                WdSession.HasKey(got, "document_content")
                && WdSession.HasKey(got, "code_level")
                && WdSession.HasKey(got, "processing_method")
                && WdSession.HasKey(got, "note")
                && WdSession.HasKey(got, "channel_id")
                && WdSession.HasKey(got, "host"),
                "缺字段");
            Ok(run, "wc-local",
                string.Equals(WdSession.Str(got, "processing_method"), "Local", StringComparison.OrdinalIgnoreCase),
                "processing_method=" + WdSession.Str(got, "processing_method"));
            Ok(run, "wc-no-snapshot-key",
                !WdSession.HasKey(got, "snapshot")
                && !WdSession.HasKey(got, "display")
                && !WdSession.HasKey(got, "display_contents"),
                "Data 带了 snapshot 旁路");
            Ok(run, "wc-id-short",
                WdSession.IsShortId(WdSession.Str(got, "channel_id"))
                && WdSession.Str(got, "channel_id").IndexOf(":", StringComparison.Ordinal) < 0,
                "id=" + WdSession.Str(got, "channel_id"));
            Ok(run, "wc-host",
                string.Equals(WdSession.Str(got, "host"), s.ExpectedKind, StringComparison.OrdinalIgnoreCase),
                "host=" + WdSession.Str(got, "host"));
            s.CloseDocByPath(textPath);
        }

        private static void RunNegLate(TestRun run, WdSession s)
        {
            string xlsx = Path.Combine(Path.GetTempPath(), "ew-wc-" + Guid.NewGuid().ToString("N") + ".xlsx");
            ToolResult xl = s.Open(xlsx, true, "excel");
            if (!xl.Success)
            {
                run.CaseFail("wc-neg-wrong-kind-excel", xl.Error ?? "建 xlsx 失败");
            }
            else
            {
                ExpectFail(run, "wc-neg-wrong-kind-excel", s.GetContent(WdSession.Str(xl, "channel_id")), "excel");
                WdSession.CloseExcelByPath(WdSession.Str(xl, "path"));
            }

            string pptx = Path.Combine(Path.GetTempPath(), "ew-wc-" + Guid.NewGuid().ToString("N") + ".pptx");
            ToolResult ppt = s.Open(pptx, true, "powerpoint");
            if (!ppt.Success)
            {
                run.CaseFail("wc-neg-wrong-kind-ppt", ppt.Error ?? "建 pptx 失败");
            }
            else
            {
                ExpectFail(run, "wc-neg-wrong-kind-ppt", s.GetContent(WdSession.Str(ppt, "channel_id")), "ppt");
                WdSession.ClosePptByPath(WdSession.Str(ppt, "path"));
            }

            string closePath = s.NewDocxPath();
            ToolResult opened = s.Open(closePath, true, s.App);
            if (!opened.Success)
            {
                run.CaseFail("wc-neg-closed", opened.Error ?? "再建稿失败");
                return;
            }

            string cid = WdSession.Str(opened, "channel_id");
            if (!s.CloseDocByPath(closePath))
            {
                run.CaseFail("wc-neg-closed", "COM 关稿失败");
                return;
            }

            ToolResult closed = s.GetContent(cid);
            if (closed != null && !closed.Success
                && (Contains(closed.Error, "关闭")
                    || Contains(closed.Error, "找不到")
                    || Contains(closed.Error, "未知")
                    || Contains(closed.Error, "unsupported")))
            {
                run.CaseOk("wc-neg-closed");
            }
            else if (closed != null && !closed.Success)
            {
                run.CaseFail("wc-neg-closed", closed.Error ?? "");
            }
            else
            {
                run.CaseFail("wc-neg-closed", "关稿后仍成功");
            }
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
            Ok(run, "wc-no-quit", wordOk && wpsOk, wordDetail + " " + wpsDetail);
        }

        private static int CountCodes(string text, string prefix)
        {
            if (string.IsNullOrEmpty(text))
            {
                return 0;
            }

            var seen = new HashSet<string>(StringComparer.Ordinal);
            foreach (Match m in Regex.Matches(text, @"\[" + prefix + @"_([0-9a-zA-Z]{5})-start\]"))
            {
                seen.Add(m.Groups[1].Value);
            }

            return seen.Count;
        }

        private static string StripCodes(string text)
        {
            if (string.IsNullOrEmpty(text))
            {
                return "";
            }

            string t = Regex.Replace(text, @"\[(?:S|P)_[0-9a-zA-Z]{5}-(?:start|end)\]", "");
            t = Regex.Replace(t, @"</?(?:body|p)[^>]*>", "", RegexOptions.IgnoreCase);
            return t;
        }

        private static string TrimPreview(string text)
        {
            if (string.IsNullOrEmpty(text))
            {
                return "";
            }

            string t = text.Replace("\r", " ").Replace("\n", " ").Trim();
            return t.Length <= 40 ? t : t.Substring(0, 40);
        }

        private static void ExpectFail(TestRun run, string id, ToolResult r, string needle)
        {
            if (r != null && !r.Success && Contains(r.Error, needle))
            {
                run.CaseOk(id);
            }
            else if (r != null && !r.Success)
            {
                run.CaseFail(id, "失败文案不含「" + needle + "」: " + (r.Error ?? ""));
            }
            else
            {
                run.CaseFail(id, "应失败");
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
