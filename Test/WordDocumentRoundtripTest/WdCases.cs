using System;
using System.Collections.Generic;
using System.IO;
using WordAddIn1;

namespace WordDocumentRoundtripTest
{
    internal static class WdCases
    {
        public static void Run(
            TestRun run,
            string host,
            int wordBefore,
            int wpsBefore,
            bool wordWinBefore,
            bool wpsWinBefore)
        {
            var s = new WdSession(host);
            RunOpenNeg(run, s);
            RunPageNegEarly(run, s);
            string livePath = RunOpenOk(run, s);
            RunPageOk(run, s, livePath);
            RunPageNegLate(run, s);
            RunNoQuit(run, wordBefore, wpsBefore, wordWinBefore, wpsWinBefore);
        }

        private static void RunOpenNeg(TestRun run, WdSession s)
        {
            ExpectFail(run, "wd-neg-open-nopath", s.Open(null), "path");
            ExpectFail(run, "wd-neg-open-empty-path", s.Open("   "), "path");
            string nosuffix = Path.Combine(Path.GetTempPath(), "ew-wd-nosuffix-" + Guid.NewGuid().ToString("N"));
            File.WriteAllBytes(nosuffix, new byte[] { 0 });
            ExpectFail(run, "wd-neg-open-nosuffix", s.Open(nosuffix), "后缀");
            ExpectFail(run, "wd-neg-open-badsuffix", s.Open(WdSession.DummyPath(".txt")), "后缀");
            ExpectFail(
                run,
                "wd-neg-open-missing",
                s.Open(s.NewDocxPath(), false, s.App),
                "不存在");
            ExpectFail(
                run,
                "wd-neg-open-conflict-xlsx",
                s.Open(WdSession.DummyPath(".xlsx"), false, s.App),
                "表格");
            ExpectFail(
                run,
                "wd-neg-open-conflict-excel",
                s.Open(WdSession.DummyPath(".docx"), false, "excel"),
                "文字");
            ExpectFail(
                run,
                "wd-neg-open-conflict-ppt",
                s.Open(WdSession.DummyPath(".pptx"), false, s.App),
                "演示");
            string exists = WdSession.DummyPath(".docx");
            ExpectFail(run, "wd-neg-open-blank-exists", s.Open(exists, true, s.App), "已存在");
            ExpectFail(run, "wd-neg-open-bad-app", s.Open(exists, false, "360"), "app");
        }

        private static void RunPageNegEarly(TestRun run, WdSession s)
        {
            ToolResult r = s.PageInfo(null);
            if (!r.Success && Contains(r.Error, "打开"))
            {
                run.CaseOk("wd-pg-neg-no-channel");
            }
            else if (r.Success)
            {
                run.CaseSkip("wd-pg-neg-no-channel", "现场已有默认文字渠，不关用户稿");
            }
            else
            {
                run.CaseFail("wd-pg-neg-no-channel", r.Error ?? "应失败或跳过");
            }

            ExpectFail(run, "wd-pg-neg-unknown-id", s.PageInfo("w999"), "未知");
        }

        private static string RunOpenOk(TestRun run, WdSession s)
        {
            string path = s.NewDocxPath();
            ToolResult blank = s.Open(path, true, s.App);
            if (!blank.Success)
            {
                FailMany(run, blank.Error ?? "新建失败",
                    "wd-open-blank", "wd-open-created", "wd-open-kind", "wd-id-short",
                    "wd-open-name-path", "wd-open-docx", "wd-open-reuse");
                return null;
            }

            string id = WdSession.Str(blank, "channel_id");
            Ok(run, "wd-open-blank",
                string.Equals(WdSession.Str(blank, "kind"), s.ExpectedKind, StringComparison.Ordinal)
                && WdSession.IsShortId(id)
                && WdSession.Bool(blank, "created")
                && !WdSession.Bool(blank, "reused"),
                "kind/id/created " + WdSession.Str(blank, "kind") + " " + id);
            Ok(run, "wd-open-created", WdSession.Bool(blank, "created")
                && PathEquals(WdSession.Str(blank, "path"), path),
                "created/path");
            Ok(run, "wd-open-kind",
                string.Equals(WdSession.Str(blank, "kind"), s.ExpectedKind, StringComparison.Ordinal),
                "kind=" + WdSession.Str(blank, "kind"));
            Ok(run, "wd-id-short", WdSession.IsShortId(id), "id=" + id);
            string name = WdSession.Str(blank, "name");
            Ok(run, "wd-open-name-path",
                !string.IsNullOrEmpty(name) && !string.IsNullOrEmpty(WdSession.Str(blank, "path")),
                "name/path 空");

            if (!s.CloseDocByPath(path))
            {
                run.CaseFail("wd-open-docx", "关新建稿失败，无法再开已有文件");
                run.CaseFail("wd-open-reuse", "依赖再开");
                return path;
            }

            ToolResult opened = s.Open(path, false, s.App);
            if (!opened.Success)
            {
                run.CaseFail("wd-open-docx", opened.Error ?? "再开失败");
                run.CaseFail("wd-open-reuse", "再开失败");
                return path;
            }

            string id2 = WdSession.Str(opened, "channel_id");
            Ok(run, "wd-open-docx",
                string.Equals(WdSession.Str(opened, "kind"), s.ExpectedKind, StringComparison.Ordinal)
                && WdSession.IsShortId(id2),
                "kind/id " + WdSession.Str(opened, "kind") + " " + id2);

            ToolResult again = s.Open(path, false, s.App);
            Ok(run, "wd-open-reuse",
                again.Success
                && string.Equals(WdSession.Str(again, "channel_id"), id2, StringComparison.Ordinal)
                && (WdSession.Bool(again, "reused")
                    || string.Equals(WdSession.Str(again, "channel_id"), id2, StringComparison.Ordinal)),
                again.Success ? "短号/reused" : (again.Error ?? "复用失败"));
            return path;
        }

        private static void RunPageOk(TestRun run, WdSession s, string livePath)
        {
            if (string.IsNullOrEmpty(livePath))
            {
                FailMany(run, "依赖打开",
                    "wd-pg-blank", "wd-pg-default", "wd-pg-explicit-id",
                    "wd-pg-multipage", "wd-pg-no-snapshot", "wd-pg-cursor-range", "wd-pg-title");
                return;
            }

            ToolResult opened = s.Open(livePath, false, s.App);
            string id = WdSession.Str(opened, "channel_id");
            ToolResult byId = s.PageInfo(id);
            if (!byId.Success)
            {
                FailMany(run, byId.Error ?? "查页失败",
                    "wd-pg-blank", "wd-pg-default", "wd-pg-explicit-id",
                    "wd-pg-multipage", "wd-pg-no-snapshot", "wd-pg-cursor-range", "wd-pg-title");
                return;
            }

            int pages = WdSession.IntOf(byId, "total_pages");
            int cursor = WdSession.IntOf(byId, "cursor_page_number");
            Ok(run, "wd-pg-blank",
                pages >= 1
                && string.Equals(WdSession.Str(byId, "host"), s.ExpectedKind, StringComparison.OrdinalIgnoreCase),
                "pages=" + pages + " host=" + WdSession.Str(byId, "host"));

            ToolResult def = s.PageInfo(null);
            Ok(run, "wd-pg-default",
                def.Success && string.Equals(WdSession.Str(def, "channel_id"), id, StringComparison.Ordinal),
                def.Success ? "默认渠 " + WdSession.Str(def, "channel_id") : (def.Error ?? "默认失败"));
            Ok(run, "wd-pg-explicit-id",
                string.Equals(WdSession.Str(byId, "channel_id"), id, StringComparison.Ordinal),
                "回包 id=" + WdSession.Str(byId, "channel_id"));

            bool dirty = DataLooksLikeSnapshot(byId);
            Ok(run, "wd-pg-no-snapshot", !dirty, dirty ? "回包装了正文/S_" : "");
            Ok(run, "wd-pg-cursor-range", cursor >= 1 && cursor <= pages,
                "cursor=" + cursor + " pages=" + pages);
            Ok(run, "wd-pg-title", WdSession.HasKey(byId, "doc_title"), "无 doc_title");

            if (!s.InsertPageBreak(livePath))
            {
                run.CaseFail("wd-pg-multipage", "COM 插分页符失败（未找到已开稿）");
                return;
            }

            ToolResult multi = s.PageInfo(id);
            Ok(run, "wd-pg-multipage",
                multi.Success && WdSession.IntOf(multi, "total_pages") >= 2,
                multi.Success ? "pages=" + WdSession.IntOf(multi, "total_pages") : (multi.Error ?? "再查失败"));
        }

        private static void RunPageNegLate(TestRun run, WdSession s)
        {
            string xlsx = Path.Combine(Path.GetTempPath(), "ew-wd-" + Guid.NewGuid().ToString("N") + ".xlsx");
            ToolResult xl = s.Open(xlsx, true, "excel");
            if (!xl.Success)
            {
                run.CaseFail("wd-pg-neg-wrong-kind-excel", xl.Error ?? "建 xlsx 失败");
            }
            else
            {
                string xid = WdSession.Str(xl, "channel_id");
                ExpectFail(run, "wd-pg-neg-wrong-kind-excel", s.PageInfo(xid), "excel");
                WdSession.CloseExcelByPath(WdSession.Str(xl, "path"));
            }

            string pptx = Path.Combine(Path.GetTempPath(), "ew-wd-" + Guid.NewGuid().ToString("N") + ".pptx");
            ToolResult ppt = s.Open(pptx, true, "powerpoint");
            if (!ppt.Success)
            {
                run.CaseFail("wd-pg-neg-wrong-kind-ppt", ppt.Error ?? "建 pptx 失败");
            }
            else
            {
                string pid = WdSession.Str(ppt, "channel_id");
                ExpectFail(run, "wd-pg-neg-wrong-kind-ppt", s.PageInfo(pid), "ppt");
                WdSession.ClosePptByPath(WdSession.Str(ppt, "path"));
            }

            string closePath = s.NewDocxPath();
            ToolResult opened = s.Open(closePath, true, s.App);
            if (!opened.Success)
            {
                run.CaseFail("wd-pg-neg-closed", opened.Error ?? "再建稿失败");
                return;
            }

            string cid = WdSession.Str(opened, "channel_id");
            if (!s.CloseDocByPath(closePath))
            {
                run.CaseFail("wd-pg-neg-closed", "COM 关稿失败");
                return;
            }

            ToolResult closed = s.PageInfo(cid);
            if (closed != null && !closed.Success
                && (Contains(closed.Error, "关闭")
                    || Contains(closed.Error, "找不到")
                    || Contains(closed.Error, "未知")
                    || Contains(closed.Error, "unsupported")))
            {
                run.CaseOk("wd-pg-neg-closed");
            }
            else if (closed != null && !closed.Success)
            {
                run.CaseFail("wd-pg-neg-closed", closed.Error ?? "");
            }
            else
            {
                run.CaseFail("wd-pg-neg-closed", "关稿后仍成功");
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
            Ok(run, "wd-no-quit", wordOk && wpsOk, wordDetail + " " + wpsDetail);
        }

        private static bool DataLooksLikeSnapshot(ToolResult r)
        {
            Dictionary<string, object> d = WdSession.Data(r);
            if (d == null)
            {
                return false;
            }

            foreach (string key in new[] { "snapshot", "display", "display_contents" })
            {
                if (d.ContainsKey(key))
                {
                    return true;
                }
            }

            string blob = "";
            foreach (object v in d.Values)
            {
                blob += Convert.ToString(v) ?? "";
            }

            return blob.IndexOf("S_", StringComparison.Ordinal) >= 0
                && blob.IndexOf("[S_", StringComparison.Ordinal) >= 0;
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

        private static bool PathEquals(string a, string b)
        {
            if (string.IsNullOrEmpty(a) || string.IsNullOrEmpty(b))
            {
                return false;
            }

            try
            {
                return string.Equals(
                    Path.GetFullPath(a).TrimEnd('\\'),
                    Path.GetFullPath(b).TrimEnd('\\'),
                    StringComparison.OrdinalIgnoreCase);
            }
            catch
            {
                return string.Equals(a, b, StringComparison.OrdinalIgnoreCase);
            }
        }
    }
}
