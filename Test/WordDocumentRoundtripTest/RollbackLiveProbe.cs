using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using WordAddIn1;

namespace WordDocumentRoundtripTest
{
    /// <summary>
    /// 双开专测：本稿先写后失败应回滚，旁边另一宿主稿不得被改。
    /// </summary>
    internal static class RollbackLiveProbe
    {
        private const string BeforeA = "回滚前甲句。";
        private const string BeforeB = "回滚前乙句。";
        private const string Dirty = "脏写不该留下。";
        private const string WordMarker = "Word旁路标记句。";
        private const string WpsMarker = "WPS旁路标记句。";

        public static int Run(string host)
        {
            bool targetWps = string.Equals(host, "wps", StringComparison.OrdinalIgnoreCase);
            string targetName = targetWps ? "WPS" : "Word";
            string otherName = targetWps ? "Word" : "WPS";
            Console.WriteLine(
                "双开回滚探针：" + targetName + " 先写后失败，看本稿是否复原、" + otherName + " 是否没动");
            if (!DesktopToolPipe.TryWaitUntilAvailable(8000))
            {
                Console.WriteLine("FAIL 连不上 Desktop");
                return 1;
            }

            var other = new WdSession(targetWps ? "word" : "wps");
            string otherPath = other.NewDocxPath();
            ToolResult otherOpened = other.Open(otherPath, true, other.App);
            if (!otherOpened.Success)
            {
                Console.WriteLine("FAIL 新建 " + otherName + " 旁路稿: " + otherOpened.Error);
                return 1;
            }

            string otherId = WdSession.Str(otherOpened, "channel_id");
            string otherMarker = targetWps ? WordMarker : WpsMarker;
            if (!other.TypePlain(otherPath, otherMarker))
            {
                Console.WriteLine("FAIL COM 写 " + otherName + " 旁路句");
                other.CloseDocByPath(otherPath);
                return 1;
            }

            string otherBefore = ReadBody(other, otherId);
            if (string.IsNullOrEmpty(otherBefore) || otherBefore.IndexOf(otherMarker, StringComparison.Ordinal) < 0)
            {
                Console.WriteLine("FAIL 读不到 " + otherName + " 旁路稿 " + Describe(otherBefore));
                other.CloseDocByPath(otherPath);
                return 1;
            }

            Console.WriteLine(otherName + " 旁路稿前 " + Describe(otherBefore));

            var s = new WdSession(targetWps ? "wps" : "word");
            string path = s.NewDocxPath();
            ToolResult opened = s.Open(path, true, s.App);
            if (!opened.Success)
            {
                Console.WriteLine("FAIL 新建 " + targetName + " 稿: " + opened.Error);
                other.CloseDocByPath(otherPath);
                return 1;
            }

            string id = WdSession.Str(opened, "channel_id");
            if (!s.TypePlain(path, BeforeA + BeforeB))
            {
                Console.WriteLine("FAIL COM 写句");
                s.CloseDocByPath(path);
                other.CloseDocByPath(otherPath);
                return 1;
            }

            ToolResult got = s.GetContent(id);
            string body = got.Success ? WdSession.Str(got, "document_content") : "";
            List<string> codes = ListCodes(body);
            if (codes.Count < 2)
            {
                Console.WriteLine("FAIL 未抄到两个 S_ body=" + Describe(body));
                s.CloseDocByPath(path);
                other.CloseDocByPath(otherPath);
                return 1;
            }

            var batch = new List<Dictionary<string, object>>();
            batch.Add(Action("replace", new Dictionary<string, object>
            {
                ["original"] = new Dictionary<string, object> { ["codes"] = codes[1] },
                ["new"] = Dirty,
                ["format"] = DefaultFormat()
            }));
            batch.Add(Action("replace", new Dictionary<string, object>
            {
                ["original"] = new Dictionary<string, object> { ["codes"] = codes[0] },
                ["new"] = "不该写出。",
                ["inherit_from"] = codes[0]
            }));

            ToolResult wr = s.ProcessActions(id, batch);
            Console.WriteLine(
                "process_actions success=" + wr.Success
                + " executed=" + WdSession.IntOf(wr, "executed_actions")
                + " err=" + (wr.Error ?? ""));

            string after = ReadBody(s, id);
            string otherAfter = ReadBody(other, otherId);
            s.CloseDocByPath(path);
            other.CloseDocByPath(otherPath);

            int executed = WdSession.IntOf(wr, "executed_actions");
            bool wroteThenFailed = !wr.Success && executed >= 1;
            bool targetRestored = after.IndexOf(BeforeA, StringComparison.Ordinal) >= 0
                && after.IndexOf(BeforeB, StringComparison.Ordinal) >= 0
                && after.IndexOf(Dirty, StringComparison.Ordinal) < 0
                && after.IndexOf("不该写出", StringComparison.Ordinal) < 0;
            bool otherUntouched = otherAfter.IndexOf(otherMarker, StringComparison.Ordinal) >= 0
                && otherAfter.IndexOf(Dirty, StringComparison.Ordinal) < 0
                && otherAfter.IndexOf(BeforeA, StringComparison.Ordinal) < 0
                && string.Equals(otherBefore, otherAfter, StringComparison.Ordinal);

            Console.WriteLine(targetName + " 回读 " + Describe(after));
            Console.WriteLine(otherName + " 旁路稿后 " + Describe(otherAfter));
            Console.WriteLine(targetRestored
                ? targetName + " 回滚: OK"
                : targetName + " 回滚: FAIL（脏写留下或旧句没回来）");
            Console.WriteLine(otherUntouched
                ? otherName + " 未改: OK"
                : otherName + " 未改: FAIL（旁路稿变了或读不到）");
            Console.WriteLine(wroteThenFailed
                ? "先写后败: OK"
                : "先写后败: FAIL（应 executed>=1 且整批失败）");

            if (!wroteThenFailed)
            {
                return 1;
            }

            return targetRestored && otherUntouched ? 0 : 1;
        }

        private static string ReadBody(WdSession s, string channelId)
        {
            ToolResult got = s.GetContent(channelId);
            return got.Success ? (WdSession.Str(got, "document_content") ?? "") : "";
        }

        private static Dictionary<string, object> Action(string type, Dictionary<string, object> detail)
        {
            return new Dictionary<string, object>
            {
                ["type"] = type,
                ["detail"] = detail
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

        private static List<string> ListCodes(string text)
        {
            var list = new List<string>();
            if (string.IsNullOrEmpty(text))
            {
                return list;
            }

            foreach (Match m in Regex.Matches(text, @"\[S_([0-9a-zA-Z]{5})-start\]"))
            {
                list.Add("S_" + m.Groups[1].Value);
            }

            return list;
        }

        private static string Describe(string text)
        {
            if (text == null)
            {
                return "(null)";
            }

            string one = text.Replace("\r", "\\r").Replace("\n", "\\n");
            if (one.Length > 80)
            {
                one = one.Substring(0, 80) + "...";
            }

            return "len=" + text.Length + " " + one;
        }
    }
}
