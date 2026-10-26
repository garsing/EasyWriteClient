using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using WordAddIn1;

namespace WordDocumentRoundtripTest
{
    /// <summary>
    /// 表格编号：读文 &lt;table id&gt;、建表回包、COM Title 都按出现顺序，嵌套不得把外表号发给内表。
    /// </summary>
    internal static class WtIdCases
    {
        public const int Target = 18;

        private const string TopA = "易写号顶甲。";
        private const string TopB = "易写号顶乙。";
        private const string OuterP = "易写号外表。";
        private const string InnerP = "易写号内表。";
        private const string InnerQ = "易写号内表贰。";
        private const string InnerL = "易写号内左。";
        private const string InnerR = "易写号内右。";
        private const string OuterL = "易写号外左。";
        private const string OuterR = "易写号外右。";
        private const string Grand = "易写号孙表。";
        private const string AfterInner = "易写号内表后句。";
        private const string OuterM = "易写号并外表。";
        private const string InnerM = "易写号并内表。";

        public static void Run(
            TestRun run,
            string host,
            int wordBefore,
            int wpsBefore,
            bool wordWinBefore,
            bool wpsWinBefore)
        {
            var s = new WdSession(host);
            RunTwoTop(run, s);
            RunNested(run, s, false);
            RunNested(run, s, true);
            RunTwoInnerOneCell(run, s);
            RunTwoInnerTwoCells(run, s);
            RunDeep3(run, s);
            RunStableAndDel(run, s);
            RunOuterInnerThenTop(run, s);
            RunNoQuit(run, wordBefore, wpsBefore, wordWinBefore, wpsWinBefore);
        }

        private static void RunTwoTop(TestRun run, WdSession s)
        {
            string[] ids = { "wt-id-two-top-order", "wt-id-two-top-return" };
            string path, ch;
            if (!TryOpen(run, s, out path, out ch, ids))
            {
                return;
            }

            try
            {
                ToolResult a = s.CreateTableFromXml(CreateArgs(ch, WriteXml(OneCell(TopA))));
                string idA = WdSession.Str(a, "table_id");
                var after = CreateArgs(ch, WriteXml(OneCell(TopB)));
                after["target"] = new Dictionary<string, object> { ["table"] = idA };
                ToolResult b = s.CreateTableFromXml(after);
                string idB = WdSession.Str(b, "table_id");
                string body = ReadBody(s, ch);
                List<string> order = CollectTableIds(body);
                string mapA = InnermostId(body, TopA);
                string mapB = InnermostId(body, TopB);
                Ok(run, "wt-id-two-top-order",
                    a.Success && b.Success
                        && order.Count >= 2
                        && order[0] == mapA
                        && order[1] == mapB
                        && mapA != mapB,
                    FailHint(a, b, "出现序不是顶甲、顶乙: " + string.Join(",", order.ToArray())));
                Ok(run, "wt-id-two-top-return",
                    a.Success && b.Success && idA == mapA && idB == mapB && idA != idB,
                    FailHint(a, b, "回包 " + idA + "," + idB + " 读文 " + mapA + "," + mapB));
            }
            finally
            {
                s.CloseDocByPath(path);
            }
        }

        private static void RunNested(TestRun run, WdSession s, bool merge)
        {
            string orderId = merge ? "wt-id-nested-merge-order" : "wt-id-nested-plain-order";
            string returnId = merge ? "wt-id-nested-merge-return" : "wt-id-nested-plain-return";
            string outerTxt = merge ? OuterM : OuterP;
            string innerTxt = merge ? InnerM : InnerP;
            string[] ids = { orderId, returnId };
            string path, ch;
            if (!TryOpen(run, s, out path, out ch, ids))
            {
                return;
            }

            try
            {
                string outerXml = merge ? MergeOneCell(outerTxt) : OneCell(outerTxt);
                ToolResult outer = s.CreateTableFromXml(CreateArgs(ch, WriteXml(outerXml)));
                string outerId = WdSession.Str(outer, "table_id");
                string sent = CodeContaining(ReadBody(s, ch), outerTxt, "S");
                if (!outer.Success || string.IsNullOrEmpty(sent))
                {
                    FailMany(run, outer.Error ?? "外表无 S_", ids);
                    return;
                }

                var innerArgs = CreateArgs(ch, WriteXml(OneCell(innerTxt)));
                innerArgs["target"] = new Dictionary<string, object> { ["codes"] = sent };
                ToolResult inner = s.CreateTableFromXml(innerArgs);
                string returned = WdSession.Str(inner, "table_id");
                string body = ReadBody(s, ch);
                List<string> order = CollectTableIds(body);
                string mapOuter = InnermostId(body, outerTxt);
                string mapInner = InnermostId(body, innerTxt);
                Ok(run, orderId,
                    inner.Success
                        && order.Count >= 2
                        && order[0] == mapOuter
                        && order[1] == mapInner
                        && mapOuter != mapInner
                        && mapOuter == outerId,
                    FailHint(outer, inner,
                        "序 " + string.Join(",", order.ToArray())
                        + " 外=" + mapOuter + " 内=" + mapInner));
                Ok(run, returnId,
                    inner.Success && returned == mapInner && returned != mapOuter,
                    FailHint(outer, inner, "回包 " + returned + " 读文内 " + mapInner + " 外 " + mapOuter));
            }
            finally
            {
                s.CloseDocByPath(path);
            }
        }

        private static void RunTwoInnerOneCell(TestRun run, WdSession s)
        {
            string[] ids = { "wt-id-two-inner-cell-order", "wt-id-two-inner-cell-distinct" };
            string path, ch;
            if (!TryOpen(run, s, out path, out ch, ids))
            {
                return;
            }

            try
            {
                ToolResult outer = s.CreateTableFromXml(CreateArgs(ch, WriteXml(MergeOneCell(OuterP))));
                string sent = CodeContaining(ReadBody(s, ch), OuterP, "S");
                if (!outer.Success || string.IsNullOrEmpty(sent))
                {
                    FailMany(run, outer.Error ?? "外表无 S_", ids);
                    return;
                }

                var inner1Args = CreateArgs(ch, WriteXml(OneCell(InnerP)));
                inner1Args["target"] = new Dictionary<string, object> { ["codes"] = sent };
                ToolResult inner1 = s.CreateTableFromXml(inner1Args);
                string id1 = WdSession.Str(inner1, "table_id");
                var inner2Args = CreateArgs(ch, WriteXml(OneCell(InnerQ)));
                inner2Args["target"] = new Dictionary<string, object> { ["table"] = id1 };
                ToolResult inner2 = s.CreateTableFromXml(inner2Args);
                string body = ReadBody(s, ch);
                List<string> order = CollectTableIds(body);
                string mapO = InnermostId(body, OuterP);
                string map1 = InnermostId(body, InnerP);
                string map2 = InnermostId(body, InnerQ);
                Ok(run, "wt-id-two-inner-cell-order",
                    inner1.Success && inner2.Success
                        && order.Count >= 3
                        && order[0] == mapO
                        && order[1] == map1
                        && order[2] == map2,
                    FailHint(inner1, inner2, "序 " + string.Join(",", order.ToArray())));
                Ok(run, "wt-id-two-inner-cell-distinct",
                    inner1.Success && inner2.Success
                        && Distinct3(mapO, map1, map2)
                        && WdSession.Str(inner2, "table_id") == map2,
                    FailHint(inner1, inner2, "外=" + mapO + " 内1=" + map1 + " 内2=" + map2));
            }
            finally
            {
                s.CloseDocByPath(path);
            }
        }

        private static void RunTwoInnerTwoCells(TestRun run, WdSession s)
        {
            string[] ids = { "wt-id-two-inner-cols-order", "wt-id-two-inner-cols-map" };
            string path, ch;
            if (!TryOpen(run, s, out path, out ch, ids))
            {
                return;
            }

            try
            {
                ToolResult outer = s.CreateTableFromXml(CreateArgs(ch, WriteXml(TwoCell(OuterL, OuterR))));
                string body0 = ReadBody(s, ch);
                string sentL = CodeContaining(body0, OuterL, "S");
                string sentR = CodeContaining(body0, OuterR, "S");
                if (!outer.Success || string.IsNullOrEmpty(sentL) || string.IsNullOrEmpty(sentR))
                {
                    FailMany(run, outer.Error ?? "左右格无 S_", ids);
                    return;
                }

                var leftArgs = CreateArgs(ch, WriteXml(OneCell(InnerL)));
                leftArgs["target"] = new Dictionary<string, object> { ["codes"] = sentL };
                ToolResult left = s.CreateTableFromXml(leftArgs);
                var rightArgs = CreateArgs(ch, WriteXml(OneCell(InnerR)));
                rightArgs["target"] = new Dictionary<string, object> { ["codes"] = sentR };
                ToolResult right = s.CreateTableFromXml(rightArgs);
                string body = ReadBody(s, ch);
                List<string> order = CollectTableIds(body);
                string mapO = InnermostId(body, OuterL);
                string mapL = InnermostId(body, InnerL);
                string mapR = InnermostId(body, InnerR);
                Ok(run, "wt-id-two-inner-cols-order",
                    left.Success && right.Success
                        && order.Count >= 3
                        && order[0] == mapO
                        && order[1] == mapL
                        && order[2] == mapR,
                    FailHint(left, right, "序 " + string.Join(",", order.ToArray())));
                Ok(run, "wt-id-two-inner-cols-map",
                    left.Success && right.Success
                        && Distinct3(mapO, mapL, mapR)
                        && WdSession.Str(left, "table_id") == mapL
                        && WdSession.Str(right, "table_id") == mapR,
                    FailHint(left, right, "外=" + mapO + " 左内=" + mapL + " 右内=" + mapR));
            }
            finally
            {
                s.CloseDocByPath(path);
            }
        }

        private static void RunDeep3(TestRun run, WdSession s)
        {
            string[] ids = { "wt-id-deep3-order", "wt-id-deep3-map" };
            string path, ch;
            if (!TryOpen(run, s, out path, out ch, ids))
            {
                return;
            }

            try
            {
                ToolResult outer = s.CreateTableFromXml(CreateArgs(ch, WriteXml(MergeOneCell(OuterP))));
                string sentO = CodeContaining(ReadBody(s, ch), OuterP, "S");
                if (!outer.Success || string.IsNullOrEmpty(sentO))
                {
                    FailMany(run, outer.Error ?? "外表无 S_", ids);
                    return;
                }

                var innerArgs = CreateArgs(ch, WriteXml(OneCell(InnerP)));
                innerArgs["target"] = new Dictionary<string, object> { ["codes"] = sentO };
                ToolResult inner = s.CreateTableFromXml(innerArgs);
                string sentI = CodeContaining(ReadBody(s, ch), InnerP, "S");
                if (!inner.Success || string.IsNullOrEmpty(sentI))
                {
                    FailMany(run, inner.Error ?? "内表无 S_", ids);
                    return;
                }

                var grandArgs = CreateArgs(ch, WriteXml(OneCell(Grand)));
                grandArgs["target"] = new Dictionary<string, object> { ["codes"] = sentI };
                ToolResult grand = s.CreateTableFromXml(grandArgs);
                string body = ReadBody(s, ch);
                List<string> order = CollectTableIds(body);
                string mapO = InnermostId(body, OuterP);
                string mapI = InnermostId(body, InnerP);
                string mapG = InnermostId(body, Grand);
                Ok(run, "wt-id-deep3-order",
                    grand.Success
                        && order.Count >= 3
                        && order[0] == mapO
                        && order[1] == mapI
                        && order[2] == mapG,
                    FailHint(inner, grand, "序 " + string.Join(",", order.ToArray())));
                Ok(run, "wt-id-deep3-map",
                    grand.Success
                        && Distinct3(mapO, mapI, mapG)
                        && WdSession.Str(inner, "table_id") == mapI
                        && WdSession.Str(grand, "table_id") == mapG,
                    FailHint(inner, grand, "外=" + mapO + " 内=" + mapI + " 孙=" + mapG));
            }
            finally
            {
                s.CloseDocByPath(path);
            }
        }

        private static void RunStableAndDel(TestRun run, WdSession s)
        {
            string[] ids =
            {
                "wt-id-reread", "wt-id-after-ins",
                "wt-id-del-inner-outer", "wt-id-del-one-of-two-inner"
            };
            string path, ch;
            if (!TryOpen(run, s, out path, out ch, ids))
            {
                return;
            }

            try
            {
                ToolResult outer = s.CreateTableFromXml(CreateArgs(ch, WriteXml(MergeOneCell(OuterM))));
                string sent = CodeContaining(ReadBody(s, ch), OuterM, "S");
                if (!outer.Success || string.IsNullOrEmpty(sent))
                {
                    FailMany(run, outer.Error ?? "并外表无 S_", ids);
                    return;
                }

                var innerArgs = CreateArgs(ch, WriteXml(OneCell(InnerM)));
                innerArgs["target"] = new Dictionary<string, object> { ["codes"] = sent };
                ToolResult inner = s.CreateTableFromXml(innerArgs);
                string body1 = ReadBody(s, ch);
                string body2 = ReadBody(s, ch);
                Ok(run, "wt-id-reread",
                    inner.Success
                        && InnermostId(body1, OuterM) == InnermostId(body2, OuterM)
                        && InnermostId(body1, InnerM) == InnermostId(body2, InnerM)
                        && JoinIds(body1) == JoinIds(body2),
                    inner.Success ? "再读编号变了 " + JoinIds(body1) + " -> " + JoinIds(body2) : (inner.Error ?? ""));

                string outerKeep = InnermostId(body2, OuterM);
                string innerKeep = InnermostId(body2, InnerM);
                ToolResult ins = s.ProcessActions(ch, new List<Dictionary<string, object>>
                {
                    new Dictionary<string, object>
                    {
                        ["type"] = "insert",
                        ["detail"] = new Dictionary<string, object>
                        {
                            ["target"] = new Dictionary<string, object> { ["table"] = innerKeep },
                            ["insert"] = AfterInner,
                            ["format"] = DefaultFormat()
                        }
                    }
                });
                string body3 = ReadBody(s, ch);
                Ok(run, "wt-id-after-ins",
                    ins.Success
                        && body3.IndexOf(AfterInner, StringComparison.Ordinal) >= 0
                        && InnermostId(body3, OuterM) == outerKeep
                        && InnermostId(body3, InnerM) == innerKeep,
                    ins.Success ? "表后插入后编号变了 外" + InnermostId(body3, OuterM) + " 内" + InnermostId(body3, InnerM)
                        : (ins.Error ?? ""));

                ToolResult delInner = s.DeleteTable(new Dictionary<string, object>
                {
                    ["channel_id"] = ch,
                    ["table_id"] = innerKeep
                });
                string afterDel = ReadBody(s, ch);
                Ok(run, "wt-id-del-inner-outer",
                    delInner.Success
                        && afterDel.IndexOf(InnerM, StringComparison.Ordinal) < 0
                        && afterDel.IndexOf(OuterM, StringComparison.Ordinal) >= 0
                        && InnermostId(afterDel, OuterM) == outerKeep,
                    delInner.Success
                        ? "删内表后外表号 " + InnermostId(afterDel, OuterM) + " 原 " + outerKeep
                        : (delInner.Error ?? ""));

                s.CloseDocByPath(path);
                if (!TryOpen(run, s, out path, out ch, "wt-id-del-one-of-two-inner"))
                {
                    return;
                }

                ToolResult o2 = s.CreateTableFromXml(CreateArgs(ch, WriteXml(MergeOneCell(OuterP))));
                string s2 = CodeContaining(ReadBody(s, ch), OuterP, "S");
                var i1a = CreateArgs(ch, WriteXml(OneCell(InnerP)));
                i1a["target"] = new Dictionary<string, object> { ["codes"] = s2 };
                ToolResult i1 = s.CreateTableFromXml(i1a);
                string firstInner = WdSession.Str(i1, "table_id");
                var i2a = CreateArgs(ch, WriteXml(OneCell(InnerQ)));
                i2a["target"] = new Dictionary<string, object> { ["table"] = firstInner };
                ToolResult i2 = s.CreateTableFromXml(i2a);
                string bodyTwo = ReadBody(s, ch);
                string keepOuter = InnermostId(bodyTwo, OuterP);
                string keepSecond = InnermostId(bodyTwo, InnerQ);
                ToolResult delFirst = s.DeleteTable(new Dictionary<string, object>
                {
                    ["channel_id"] = ch,
                    ["table_id"] = InnermostId(bodyTwo, InnerP)
                });
                string afterFirst = ReadBody(s, ch);
                Ok(run, "wt-id-del-one-of-two-inner",
                    o2.Success && i1.Success && i2.Success && delFirst.Success
                        && afterFirst.IndexOf(InnerP, StringComparison.Ordinal) < 0
                        && InnermostId(afterFirst, OuterP) == keepOuter
                        && InnermostId(afterFirst, InnerQ) == keepSecond,
                    delFirst.Success
                        ? "删其一后 外" + InnermostId(afterFirst, OuterP) + " 内贰" + InnermostId(afterFirst, InnerQ)
                        : (delFirst.Error ?? ""));
            }
            finally
            {
                s.CloseDocByPath(path);
            }
        }

        private static void RunOuterInnerThenTop(TestRun run, WdSession s)
        {
            string[] ids = { "wt-id-outer-inner-top" };
            string path, ch;
            if (!TryOpen(run, s, out path, out ch, ids))
            {
                return;
            }

            try
            {
                ToolResult outer = s.CreateTableFromXml(CreateArgs(ch, WriteXml(OneCell(OuterP))));
                string outerId = WdSession.Str(outer, "table_id");
                string sent = CodeContaining(ReadBody(s, ch), OuterP, "S");
                var innerArgs = CreateArgs(ch, WriteXml(OneCell(InnerP)));
                innerArgs["target"] = new Dictionary<string, object> { ["codes"] = sent };
                ToolResult inner = s.CreateTableFromXml(innerArgs);
                var topArgs = CreateArgs(ch, WriteXml(OneCell(TopB)));
                topArgs["target"] = new Dictionary<string, object> { ["table"] = outerId };
                ToolResult top = s.CreateTableFromXml(topArgs);
                string body = ReadBody(s, ch);
                List<string> order = CollectTableIds(body);
                string mapO = InnermostId(body, OuterP);
                string mapI = InnermostId(body, InnerP);
                string mapB = InnermostId(body, TopB);
                Ok(run, "wt-id-outer-inner-top",
                    outer.Success && inner.Success && top.Success
                        && order.Count >= 3
                        && order[0] == mapO
                        && order[1] == mapI
                        && order[2] == mapB
                        && Distinct3(mapO, mapI, mapB),
                    FailHint(inner, top, "序 " + string.Join(",", order.ToArray())));
            }
            finally
            {
                s.CloseDocByPath(path);
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
            Ok(run, "wt-id-no-quit", wordOk && wpsOk, wordDetail + " " + wpsDetail);
        }

        private static bool TryOpen(
            TestRun run,
            WdSession s,
            out string path,
            out string id,
            params string[] failIds)
        {
            path = s.NewDocxPath();
            id = null;
            ToolResult opened = s.Open(path, true, s.App);
            if (!opened.Success)
            {
                FailMany(run, opened.Error ?? "新建失败", failIds);
                return false;
            }

            id = WdSession.Str(opened, "channel_id");
            return true;
        }

        private static Dictionary<string, object> CreateArgs(string channelId, string path)
        {
            var args = new Dictionary<string, object> { ["path"] = path ?? "" };
            if (channelId != null)
            {
                args["channel_id"] = channelId;
            }

            return args;
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

        private static string OneCell(string text)
        {
            return
                "<TableConfig><General><TableFontColor>000000</TableFontColor></General><Table><Properties><ColWidths><float>120</float></ColWidths></Properties><Data>"
                + "<tr><td>" + text + "</td></tr>"
                + "</Data></Table></TableConfig>";
        }

        private static string MergeOneCell(string text)
        {
            return
                "<TableConfig><General><TableFontColor>000000</TableFontColor></General><Table><Properties><ColWidths><float>120</float><float>180</float></ColWidths></Properties><Data>"
                + "<tr><td colspan=\"2\">" + text + "</td></tr>"
                + "</Data></Table></TableConfig>";
        }

        private static string TwoCell(string left, string right)
        {
            return
                "<TableConfig><General><TableFontColor>000000</TableFontColor></General><Table><Properties><ColWidths><float>120</float><float>180</float></ColWidths></Properties><Data>"
                + "<tr><td>" + left + "</td><td>" + right + "</td></tr>"
                + "</Data></Table></TableConfig>";
        }

        private static string WriteXml(string xml)
        {
            string path = Path.Combine(Path.GetTempPath(), "ew-wt-id-" + Guid.NewGuid().ToString("N") + ".xml");
            File.WriteAllText(path, xml ?? "");
            return path;
        }

        private static string ReadBody(WdSession s, string id)
        {
            ToolResult got = s.GetContent(id);
            return got.Success ? WdSession.Str(got, "document_content") : "";
        }

        private static List<string> CollectTableIds(string body)
        {
            var ids = new List<string>();
            if (string.IsNullOrEmpty(body))
            {
                return ids;
            }

            MatchCollection matches = Regex.Matches(body, "<table id=\"(T_[^\"]+)\"");
            for (int i = 0; i < matches.Count; i++)
            {
                ids.Add(matches[i].Groups[1].Value);
            }

            return ids;
        }

        private static string JoinIds(string body)
        {
            return string.Join(",", CollectTableIds(body).ToArray());
        }

        private static string InnermostId(string body, string needle)
        {
            if (string.IsNullOrEmpty(body) || string.IsNullOrEmpty(needle))
            {
                return null;
            }

            int at = body.IndexOf(needle, StringComparison.Ordinal);
            if (at < 0)
            {
                return null;
            }

            MatchCollection matches = Regex.Matches(body.Substring(0, at), "<table id=\"(T_[^\"]+)\"");
            if (matches.Count == 0)
            {
                return null;
            }

            return matches[matches.Count - 1].Groups[1].Value;
        }

        private static string CodeContaining(string text, string needle, string prefix)
        {
            if (string.IsNullOrEmpty(text) || string.IsNullOrEmpty(needle))
            {
                return "";
            }

            MatchCollection matches = Regex.Matches(
                text,
                @"\[" + prefix + @"_([0-9a-zA-Z]{5})-start\](.*?)\[\s*" + prefix + @"_\1-end\]",
                RegexOptions.Singleline);
            foreach (Match m in matches)
            {
                if (m.Groups[2].Value.IndexOf(needle, StringComparison.Ordinal) >= 0)
                {
                    return prefix + "_" + m.Groups[1].Value;
                }
            }

            return "";
        }

        private static bool Distinct3(string a, string b, string c)
        {
            return !string.IsNullOrEmpty(a)
                && !string.IsNullOrEmpty(b)
                && !string.IsNullOrEmpty(c)
                && a != b && a != c && b != c;
        }

        private static string FailHint(ToolResult a, ToolResult b, string extra)
        {
            if (a != null && !a.Success)
            {
                return a.Error ?? extra;
            }

            if (b != null && !b.Success)
            {
                return b.Error ?? extra;
            }

            return extra;
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
