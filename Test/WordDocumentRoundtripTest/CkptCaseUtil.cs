using System;
using System.Text.RegularExpressions;
using Newtonsoft.Json;
using WordAddIn1;

namespace WordDocumentRoundtripTest
{
    internal static class CkptCaseUtil
    {
        public static void ExpectNoChannel(TestRun run, string id, ToolResult r)
        {
            if (!r.Success
                && (Contains(r.Error, "打开")
                    || Contains(r.Error, "unsupported")
                    || Contains(r.Error, "browser")
                    || Contains(r.Error, "渠道")))
            {
                run.CaseOk(id);
            }
            else if (r.Success)
            {
                run.CaseSkip(id, "现场已有默认渠，不关用户稿");
            }
            else
            {
                run.CaseFail(id, r.Error ?? "应失败或跳过");
            }
        }

        public static void ExpectFail(TestRun run, string id, ToolResult r, string needle)
        {
            if (r != null && !r.Success && (Contains(r.Error, needle) || BlobHas(r, needle)))
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

        public static bool BlobHas(ToolResult r, string needle)
        {
            if (r == null || r.Data == null || string.IsNullOrEmpty(needle))
            {
                return false;
            }

            try
            {
                return JsonConvert.SerializeObject(r.Data)
                    .IndexOf(needle, StringComparison.OrdinalIgnoreCase) >= 0;
            }
            catch
            {
                return false;
            }
        }

        public static string DataText(ToolResult r)
        {
            if (r == null || r.Data == null)
            {
                return "";
            }

            if (r.Data is string s)
            {
                return s;
            }

            try
            {
                return JsonConvert.SerializeObject(r.Data);
            }
            catch
            {
                return r.Data.ToString() ?? "";
            }
        }

        public static string FirstCopyTs(string listText)
        {
            if (string.IsNullOrEmpty(listText))
            {
                return "";
            }

            Match m = Regex.Match(listText, @"副本(\S+)");
            return m.Success ? m.Groups[1].Value : "";
        }

        public static bool Contains(string text, string needle)
        {
            return !string.IsNullOrEmpty(text)
                && text.IndexOf(needle, StringComparison.OrdinalIgnoreCase) >= 0;
        }

        public static string Snip(string body)
        {
            if (string.IsNullOrEmpty(body))
            {
                return "";
            }

            return body.Length <= 80 ? body : body.Substring(0, 80);
        }

        public static void Ok(TestRun run, string id, bool cond, string reason)
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

        public static void FailMany(TestRun run, string reason, params string[] ids)
        {
            foreach (string id in ids)
            {
                run.CaseFail(id, reason);
            }
        }

        public static bool TryOpen(
            TestRun run,
            OfficeCkptSession s,
            out string path,
            out string id,
            params string[] ids)
        {
            path = s.NewPath();
            id = null;
            ToolResult opened = null;
            for (int i = 0; i < 3; i++)
            {
                opened = s.Open(path);
                if (opened.Success)
                {
                    break;
                }

                System.Threading.Thread.Sleep(400);
            }

            if (opened == null || !opened.Success)
            {
                FailMany(run, "开稿失败: " + (opened == null ? "" : (opened.Error ?? "")), ids);
                s.CloseByPath(path);
                return false;
            }

            id = WdSession.Str(opened, "channel_id");
            return true;
        }

        public static void RunNoQuit(TestRun run, string id, string proc, bool winBefore)
        {
            if (winBefore && !OfficeAlive.HasWindow(proc))
            {
                run.CaseFail(id, "用户 " + proc + " 窗没了");
                return;
            }

            run.CaseOk(id);
        }
    }
}
