using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using WordAddIn1;

namespace WordDocumentRoundtripTest
{
    /// <summary>
    /// 表格盖戳与定位：一份文档只走 Title 或只走书签；再读不换号；按 T_ 删/插/抽落到同一张表。
    /// </summary>
    internal static class WtStCases
    {
        public const int Target = 16;

        private const string TopA = "易写戳顶甲。";
        private const string TopB = "易写戳顶乙。";
        private const string OuterP = "易写戳外表。";
        private const string InnerP = "易写戳内表。";
        private const string AfterA = "易写戳表后句。";
        private const string AfterInner = "易写戳内表后句。";
        private const string GeoFill = "易写戳填空。";
        private const string CompatSrc =
            @"D:\documents\易写\易写产品试用\2. 体验素材\1-公司类信贷业务调查模板（长安跨越2） (1).docx";

        public static void Run(
            TestRun run,
            string host,
            int wordBefore,
            int wpsBefore,
            bool wordWinBefore,
            bool wpsWinBefore)
        {
            var s = new WdSession(host);
            RunNewDocChannel(run, s);
            RunReread(run, s);
            RunOps(run, s);
            RunNested(run, s);
            RunCompat(run, s);
            RunNoQuit(run, wordBefore, wpsBefore, wordWinBefore, wpsWinBefore);
        }

        private static void RunNewDocChannel(TestRun run, WdSession s)
        {
            string[] ids =
            {
                "wt-st-title-or-bookmark",
                "wt-st-two-same-channel",
                "wt-st-bookmark-covers-table"
            };
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
                string body = ReadBody(s, ch);
                List<string> order = CollectTableIds(body);
                string mapA = InnermostId(body, TopA);
                string mapB = InnermostId(body, TopB);
                bool idsOk = a.Success && b.Success
                    && order.Count >= 2
                    && IsAlnumTableId(mapA)
                    && IsAlnumTableId(mapB)
                    && mapA != mapB
                    && !HasHashChars(mapA)
                    && !HasHashChars(mapB);
                var delB = s.DeleteTable(new Dictionary<string, object>
                {
                    ["channel_id"] = ch,
                    ["table_id"] = mapB
                });
                string afterDel = ReadBody(s, ch);
                bool sameChannelOps = delB.Success
                    && afterDel.IndexOf(TopB, StringComparison.Ordinal) < 0
                    && afterDel.IndexOf(TopA, StringComparison.Ordinal) >= 0
                    && InnermostId(afterDel, TopA) == mapA;
                Ok(run, "wt-st-title-or-bookmark",
                    idsOk,
                    FailHint(a, b, "表号非法 " + mapA + "," + mapB));
                Ok(run, "wt-st-two-same-channel",
                    idsOk && mapA == idA && order[0] == mapA && order[1] == mapB,
                    FailHint(a, b, "两号 " + mapA + "," + mapB + " 序 " + JoinIds(body)));
                Ok(run, "wt-st-bookmark-covers-table",
                    idsOk && sameChannelOps,
                    FailHint(delB, null, "按乙号删除未只删乙"));
            }
            finally
            {
                if (path != null)
                {
                    s.CloseDocByPath(path);
                }
            }
        }

        private static void RunReread(TestRun run, WdSession s)
        {
            string[] ids = { "wt-st-reread-same-ids", "wt-st-reread-nested" };
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
                s.CreateTableFromXml(after);
                string first = JoinIds(ReadBody(s, ch));
                string second = JoinIds(ReadBody(s, ch));
                Ok(run, "wt-st-reread-same-ids",
                    a.Success && first.Length > 0 && first == second,
                    "先 " + first + " 再 " + second);

                string sent = CodeContaining(ReadBody(s, ch), TopA, "S");
                var innerArgs = CreateArgs(ch, WriteXml(OneCell(InnerP)));
                innerArgs["target"] = new Dictionary<string, object> { ["codes"] = sent };
                ToolResult inner = s.CreateTableFromXml(innerArgs);
                string n1 = JoinIds(ReadBody(s, ch));
                string n2 = JoinIds(ReadBody(s, ch));
                Ok(run, "wt-st-reread-nested",
                    inner.Success && n1.Length > 0 && n1 == n2 && n1.Split(',').Length >= 3,
                    FailHint(inner, null, "先 " + n1 + " 再 " + n2));
            }
            finally
            {
                s.CloseDocByPath(path);
            }
        }

        private static void RunOps(TestRun run, WdSession s)
        {
            string[] ids =
            {
                "wt-st-delete-by-id",
                "wt-st-insert-after-table",
                "wt-st-extract-format",
                "wt-st-geo-write"
            };
            string path, ch;
            if (!TryOpen(run, s, out path, out ch, ids))
            {
                return;
            }

            try
            {
                ToolResult a = s.CreateTableFromXml(CreateArgs(ch, WriteXml(TwoCell(TopA, ""))));
                string idA = WdSession.Str(a, "table_id");
                var after = CreateArgs(ch, WriteXml(OneCell(TopB)));
                after["target"] = new Dictionary<string, object> { ["table"] = idA };
                ToolResult b = s.CreateTableFromXml(after);
                string idB = WdSession.Str(b, "table_id");
                string body0 = ReadBody(s, ch);
                string mapA = InnermostId(body0, TopA);
                string mapB = InnermostId(body0, TopB);

                string extractPath = Path.Combine(
                    Path.GetTempPath(),
                    "ew-wt-st-fmt-" + Guid.NewGuid().ToString("N") + ".xml");
                var extractArgs = new Dictionary<string, object>
                {
                    ["channel_id"] = ch,
                    ["path"] = extractPath,
                    ["table_selection"] = new Dictionary<string, object>
                    {
                        ["type"] = "by_id",
                        ["table_id"] = mapA
                    }
                };
                ToolResult extracted = s.ExtractTableFormat(extractArgs);
                string extractedId = NestedStr(extracted, "table_info", "table_id");

                Ok(run, "wt-st-extract-format",
                    extracted.Success
                        && extractedId == mapA
                        && extractedId != mapB
                        && (WdSession.Bool(extracted, "format_saved")
                            || !string.IsNullOrEmpty(WdSession.Str(extracted, "format_content"))
                            || File.Exists(extractPath)),
                    FailHint(extracted, null, "抽到 " + extractedId + " 期望 " + mapA));

                string geo = mapA + "_r1_c2";
                var geoAction = new Dictionary<string, object>
                {
                    ["type"] = "replace",
                    ["detail"] = new Dictionary<string, object>
                    {
                        ["original"] = new Dictionary<string, object> { ["codes"] = geo },
                        ["new"] = GeoFill,
                        ["format"] = DefaultFormat()
                    }
                };
                ToolResult geoWrite = s.ProcessActions(
                    ch,
                    new List<Dictionary<string, object>> { geoAction });
                string afterGeo = ReadBody(s, ch);
                Ok(run, "wt-st-geo-write",
                    geoWrite.Success
                        && afterGeo.IndexOf(GeoFill, StringComparison.Ordinal) >= 0
                        && InnermostId(afterGeo, GeoFill) == mapA,
                    FailHint(geoWrite, null, "填空未进甲"));

                var insAction = new Dictionary<string, object>
                {
                    ["type"] = "insert",
                    ["detail"] = new Dictionary<string, object>
                    {
                        ["target"] = new Dictionary<string, object> { ["table"] = mapA },
                        ["insert"] = AfterA,
                        ["format"] = DefaultFormat()
                    }
                };
                ToolResult inserted = s.ProcessActions(
                    ch,
                    new List<Dictionary<string, object>> { insAction });
                string afterIns = ReadBody(s, ch);
                int posAfter = afterIns.IndexOf(AfterA, StringComparison.Ordinal);
                int posA = afterIns.IndexOf(TopA, StringComparison.Ordinal);
                int posB = afterIns.IndexOf(TopB, StringComparison.Ordinal);
                Ok(run, "wt-st-insert-after-table",
                    inserted.Success && posAfter > posA && (posB < 0 || posAfter < posB)
                        && InnermostId(afterIns, TopA) == mapA,
                    FailHint(inserted, null, "表后插入位置不对"));

                var delArgs = new Dictionary<string, object>
                {
                    ["channel_id"] = ch,
                    ["table_id"] = mapB
                };
                ToolResult deleted = s.DeleteTable(delArgs);
                string afterDel = ReadBody(s, ch);
                Ok(run, "wt-st-delete-by-id",
                    deleted.Success
                        && afterDel.IndexOf(TopB, StringComparison.Ordinal) < 0
                        && afterDel.IndexOf(TopA, StringComparison.Ordinal) >= 0
                        && InnermostId(afterDel, TopA) == mapA,
                    FailHint(deleted, null, "删乙伤甲"));
            }
            finally
            {
                s.CloseDocByPath(path);
            }
        }

        private static void RunNested(TestRun run, WdSession s)
        {
            string[] ids =
            {
                "wt-st-nested-delete-inner",
                "wt-st-nested-bookmark-span",
                "wt-st-nested-insert-after-inner"
            };
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
                if (!outer.Success || string.IsNullOrEmpty(sent))
                {
                    FailMany(run, outer.Error ?? "外表无 S_", ids);
                    return;
                }

                var innerArgs = CreateArgs(ch, WriteXml(OneCell(InnerP)));
                innerArgs["target"] = new Dictionary<string, object> { ["codes"] = sent };
                ToolResult inner = s.CreateTableFromXml(innerArgs);
                string innerId = WdSession.Str(inner, "table_id");
                string body = ReadBody(s, ch);
                string mapOuter = InnermostId(body, OuterP);
                string mapInner = InnermostId(body, InnerP);

                var insAction = new Dictionary<string, object>
                {
                    ["type"] = "insert",
                    ["detail"] = new Dictionary<string, object>
                    {
                        ["target"] = new Dictionary<string, object> { ["table"] = mapInner },
                        ["insert"] = AfterInner,
                        ["format"] = DefaultFormat()
                    }
                };
                ToolResult inserted = s.ProcessActions(
                    ch,
                    new List<Dictionary<string, object>> { insAction });
                string afterIns = ReadBody(s, ch);
                int posInnerTxt = afterIns.IndexOf(InnerP, StringComparison.Ordinal);
                int posAfterInner = afterIns.IndexOf(AfterInner, StringComparison.Ordinal);
                int posOuterTxt = afterIns.IndexOf(OuterP, StringComparison.Ordinal);
                Ok(run, "wt-st-nested-insert-after-inner",
                    inserted.Success
                        && posAfterInner >= 0
                        && posInnerTxt >= 0
                        && posAfterInner > posInnerTxt
                        && InnermostId(afterIns, InnerP) == mapInner
                        && InnermostId(afterIns, OuterP) == mapOuter,
                    FailHint(inserted, inner, "内表后插入未跟上内表"));

                var delArgs = new Dictionary<string, object>
                {
                    ["channel_id"] = ch,
                    ["table_id"] = mapInner
                };
                ToolResult deleted = s.DeleteTable(delArgs);
                string afterDel = ReadBody(s, ch);
                Ok(run, "wt-st-nested-delete-inner",
                    deleted.Success
                        && afterDel.IndexOf(InnerP, StringComparison.Ordinal) < 0
                        && afterDel.IndexOf(OuterP, StringComparison.Ordinal) >= 0
                        && InnermostId(afterDel, OuterP) == mapOuter,
                    FailHint(deleted, inner, "删内表伤外表"));

                Ok(run, "wt-st-nested-bookmark-span",
                    inner.Success && mapOuter != mapInner && IsAlnumTableId(mapInner),
                    FailHint(inner, outer, "内=" + mapInner + " 外=" + mapOuter));
            }
            finally
            {
                if (path != null)
                {
                    s.CloseDocByPath(path);
                }
            }
        }

        private static void RunCompat(TestRun run, WdSession s)
        {
            string[] ids =
            {
                "wt-st-compat-bookmark",
                "wt-st-compat-reread",
                "wt-st-compat-delete-one"
            };
            if (!File.Exists(CompatSrc))
            {
                FailMany(run, "找不到调查模板 " + CompatSrc, ids);
                return;
            }

            string path = Path.Combine(
                Path.GetTempPath(),
                "ew-wt-st-compat-" + Guid.NewGuid().ToString("N") + ".docx");
            File.Copy(CompatSrc, path, true);
            ToolResult opened = s.Open(path, false, s.App);
            if (!opened.Success)
            {
                FailMany(run, opened.Error ?? "打开副本失败", ids);
                TryDelete(path);
                return;
            }

            string ch = WdSession.Str(opened, "channel_id");
            try
            {
                string first = ReadBody(s, ch);
                List<string> order = CollectTableIds(first);
                bool allAlnum = order.Count >= 17;
                for (int i = 0; i < order.Count; i++)
                {
                    if (!IsAlnumTableId(order[i]) || HasHashChars(order[i]))
                    {
                        allAlnum = false;
                        break;
                    }
                }

                Ok(run, "wt-st-compat-bookmark",
                    allAlnum,
                    "表数 " + order.Count + " 号 " + string.Join(",", order.ToArray()));

                string second = ReadBody(s, ch);
                string joined1 = string.Join(",", order.ToArray());
                string joined2 = JoinIds(second);
                Ok(run, "wt-st-compat-reread",
                    joined1 == joined2 && order.Count >= 17,
                    "先 " + joined1 + " 再 " + joined2);

                string firstId = order[0];
                var delArgs = new Dictionary<string, object>
                {
                    ["channel_id"] = ch,
                    ["table_id"] = firstId
                };
                ToolResult deleted = s.DeleteTable(delArgs);
                List<string> after = CollectTableIds(ReadBody(s, ch));
                bool restSame = after.Count == order.Count - 1;
                if (restSame)
                {
                    for (int i = 1; i < order.Count; i++)
                    {
                        if (after[i - 1] != order[i])
                        {
                            restSame = false;
                            break;
                        }
                    }
                }

                Ok(run, "wt-st-compat-delete-one",
                    deleted.Success && restSame && after.IndexOf(firstId) < 0,
                    FailHint(deleted, null, "删第一张后序 " + string.Join(",", after.ToArray())));
            }
            finally
            {
                s.CloseDocByPath(path);
                TryDelete(path);
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
            Ok(run, "wt-st-no-quit", wordOk && wpsOk, wordDetail + " " + wpsDetail);
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

        private static string TwoCell(string left, string right)
        {
            return
                "<TableConfig><General><TableFontColor>000000</TableFontColor></General><Table><Properties><ColWidths><float>120</float><float>180</float></ColWidths></Properties><Data>"
                + "<tr><td>" + left + "</td><td>" + right + "</td></tr>"
                + "</Data></Table></TableConfig>";
        }

        private static string WriteXml(string xml)
        {
            string path = Path.Combine(Path.GetTempPath(), "ew-wt-st-" + Guid.NewGuid().ToString("N") + ".xml");
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

        private static bool IsAlnumTableId(string tableId)
        {
            return !string.IsNullOrEmpty(tableId) && Regex.IsMatch(tableId, @"^T_[A-Za-z0-9]+$");
        }

        private static string NestedStr(ToolResult r, string parent, string child)
        {
            Dictionary<string, object> d = WdSession.Data(r);
            if (d == null || !d.ContainsKey(parent) || d[parent] == null)
            {
                return "";
            }

            object raw = d[parent];
            var nested = raw as Dictionary<string, object>;
            if (nested == null && raw is Newtonsoft.Json.Linq.JObject jo)
            {
                nested = jo.ToObject<Dictionary<string, object>>();
            }

            if (nested == null || !nested.ContainsKey(child) || nested[child] == null)
            {
                return "";
            }

            return Convert.ToString(nested[child]) ?? "";
        }

        private static bool HasHashChars(string tableId)
        {
            return !string.IsNullOrEmpty(tableId)
                && (tableId.IndexOf('/') >= 0 || tableId.IndexOf('+') >= 0);
        }

        private static void TryDelete(string path)
        {
            try
            {
                if (!string.IsNullOrEmpty(path) && File.Exists(path))
                {
                    File.Delete(path);
                }
            }
            catch
            {
            }
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
