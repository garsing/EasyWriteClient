using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using WordAddIn1;
using Word = Microsoft.Office.Interop.Word;

namespace WordDocumentRoundtripTest
{
    /// <summary>
    /// 第 6 波：XML 建表、填空槽、抽/套格式、删表重建、新建后第一次 optimize。
    /// </summary>
    internal static class WtCases
    {
        public const int Target = 94;

        private const string LabelA = "易写表指标";
        private const string LabelB = "易写表WACC";
        private const string FillA = "8.46%";
        private const string FillB = "12.00%";
        private const string MergeHead = "易写表合并头";
        private const string MergeLeft = "易写表左";
        private const string MergeRight = "易写表右";
        private const string AfterSent = "易写表前句。";
        private const string AfterSentB = "易写表后句。";
        private const string RebuildA = "易写表重建甲";
        private const string RebuildB = "易写表重建乙";
        private const string ConsA = "易写表巩固甲";
        private const string ConsB = "易写表巩固乙";
        private const string ConsC = "易写表巩固丙";
        private const string ConsD = "易写表巩固丁";
        private const string ConsSent = "易写表巩固句。";
        private const string Cell1 = "易写表单格";
        private const string Row3A = "易写表三行A";
        private const string Row3B = "易写表三行B";
        private const string Row3C = "易写表三行C";
        private const string SpanTxt = "易写表重点";
        private const string RsLeft = "易写表跨行左";
        private const string RsUp = "易写表跨行上";
        private const string RsDown = "易写表跨行下";

        private static readonly string[] ExtraIds =
        {
            "wt-create-filename", "wt-create-root", "wt-create-colw-gen", "wt-create-colw-tbl",
            "wt-create-style-prop", "wt-create-pos-end", "wt-create-unk-code", "wt-create-unk-tbl",
            "wt-create-no-data", "wt-create-1x1", "wt-create-3row", "wt-create-span",
            "wt-create-rowspan", "wt-create-two", "wt-create-two-text", "wt-create-after-tbl",
            "wt-fill-short-len", "wt-fill-index", "wt-fill-bad-row", "wt-fill-row0",
            "wt-fill-writes-empty", "wt-fill-writes-obj", "wt-fill-unk-t", "wt-read-unk-t",
            "wt-read-slots", "wt-read-rows-n", "wt-fill-blank", "wt-fill-row2",
            "wt-fill-row2-text", "wt-read-after-fill", "wt-fill-no-empty", "wt-fill-del",
            "wt-xfmt-no-sel", "wt-xfmt-ctx", "wt-xfmt-unk", "wt-xfmt-no-type",
            "wt-afmt-no-id", "wt-afmt-unk", "wt-afmt-bad-xml", "wt-afmt-path",
            "wt-afmt-applied", "wt-afmt-echo", "wt-del-unk", "wt-del-twice",
            "wt-opt-unk", "wt-opt-after-del", "wt-opt-rebuild", "wt-del-keep-sent",
            "wt-read-gone", "wt-two-del-first"
        };

        public static void RunCreateAfterOnly(
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
                string path, id;
                if (!TryOpen(run, s, out path, out id, "wt-create-after"))
                {
                    return;
                }

                if (!s.TypePlain(path, AfterSent))
                {
                    FailMany(run, "COM 写句失败", "wt-create-after");
                    s.CloseDocByPath(path);
                    return;
                }

                s.ActivateDoc(path);
                ToolResult got = s.GetContent(id);
                string sent = CodeContaining(WdSession.Str(got, "document_content"), AfterSent, "S");
                if (string.IsNullOrEmpty(sent))
                {
                    FailMany(run, "无可抄含前句的 S_", "wt-create-after");
                    s.CloseDocByPath(path);
                    return;
                }

                var afterArgs = CreateArgs(id, WriteXml(GridXml(LabelA, "", LabelB, "")));
                afterArgs["target"] = new Dictionary<string, object> { ["codes"] = sent };
                ToolResult after = s.CreateTableFromXml(afterArgs);
                Ok(run, "wt-create-after",
                    after.Success && StillHas(s, id, AfterSent) && StillHas(s, id, LabelA),
                    after.Success ? "句后建表丢了句或表" : (after.Error ?? ""));
                s.CloseDocByPath(path);
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

        /// <summary>
        /// 两句话：存文 ≠ 全文，必须走原来的 Find，不能误用整篇 Content。
        /// 句后建表应夹在两句之间（第一句后）或落在第二句后。
        /// </summary>
        /// <summary>对照巩固套：首张表刚建完 vs 后面又建几张表之后，读空槽。</summary>
        public static void RunEmptySlotDump(string host)
        {
            var s = new WdSession(host);
            var dummy = new TestRun();
            string path, id;
            if (!TryOpen(dummy, s, out path, out id, "wt-wps-empty-slots"))
            {
                Console.WriteLine("open failed");
                return;
            }

            if (!s.TypePlain(path, ConsSent))
            {
                Console.WriteLine("TypePlain failed");
                s.CloseDocByPath(path);
                return;
            }

            s.ActivateDoc(path);
            ToolResult created = s.CreateTableFromXml(
                CreateArgsFilename(id, WriteXml(GridXml(ConsA, "", ConsB, ""))));
            string tableId = WdSession.Str(created, "table_id");
            Console.WriteLine("first id=" + tableId + " ok=" + created.Success);
            DumpRead(s, id, tableId, "just-created");

            var afterTbl = CreateArgs(id, WriteXml(GridXml(ConsC, "", ConsD, "")));
            afterTbl["target"] = new Dictionary<string, object> { ["table"] = tableId };
            ToolResult second = s.CreateTableFromXml(afterTbl);
            Console.WriteLine("after-tbl ok=" + second.Success + " id=" + WdSession.Str(second, "table_id"));
            DumpRead(s, id, tableId, "after-second-table");

            CreateOkHas(s, id, OneCellXml(Cell1), Cell1);
            CreateOkHas(s, id, ThreeRowXml(), Row3A);
            CreateOkHas(s, id, SpanXml(), SpanTxt);
            CreateOkHas(s, id, RowspanXml(), RsLeft);
            Console.WriteLine("firstId still=" + tableId);
            s.WriteTableFacts(path);
            DumpRead(s, id, tableId, "after-extra-tables");

            s.CloseDocByPath(path);
        }

        private static void DumpRead(WdSession s, string id, string tableId, string tag)
        {
            ToolResult read = s.ReadTableRowValues(new Dictionary<string, object>
            {
                ["channel_id"] = id,
                ["table_id"] = tableId
            });
            Console.WriteLine(tag + " success=" + read.Success + " err=" + (read.Error ?? ""));
            JObject data = DataObj(read);
            JArray rows = data == null ? null : data["rows"] as JArray;
            if (rows == null)
            {
                Console.WriteLine(tag + " rows=null data=" + (data == null ? "null" : data.ToString(Formatting.None)));
                return;
            }

            foreach (JToken token in rows)
            {
                JObject o = token as JObject;
                if (o == null)
                {
                    continue;
                }

                Console.WriteLine(tag + " row=" + o["row"] + " empty_count=" + o["empty_count"]
                    + " values=" + (o["values"] == null ? "null" : o["values"].ToString(Formatting.None))
                    + " slots=" + (o["slots"] == null ? "null" : o["slots"].ToString(Formatting.None)));
            }

            int er, ec;
            Console.WriteLine(tag + " TryFirstEmpty=" + TryFirstEmptyRow(read, out er, out ec)
                + " row=" + er + " n=" + ec);
        }

        /// <summary>表后建表后看 COM 是两张还是并成一张。</summary>
        public static void RunAfterTblJoinProbe(string host)
        {
            var s = new WdSession(host);
            var dummy = new TestRun();
            string path, id;
            if (!TryOpen(dummy, s, out path, out id, "wt-after-tbl-join"))
            {
                Console.WriteLine("open failed");
                return;
            }

            if (!s.TypePlain(path, AfterSent))
            {
                Console.WriteLine("TypePlain failed");
                s.CloseDocByPath(path);
                return;
            }

            s.ActivateDoc(path);
            ToolResult first = s.CreateTableFromXml(CreateArgs(id, WriteXml(GridXml(ConsA, "", ConsB, ""))));
            string firstId = WdSession.Str(first, "table_id");
            Console.WriteLine("first success=" + first.Success + " id=" + firstId + " err=" + (first.Error ?? ""));
            DumpTables(s, path, "after-first");

            var afterTbl = CreateArgs(id, WriteXml(GridXml(ConsC, "", ConsD, "")));
            afterTbl["target"] = new Dictionary<string, object> { ["table"] = firstId };
            ToolResult second = s.CreateTableFromXml(afterTbl);
            string secondId = WdSession.Str(second, "table_id");
            Console.WriteLine("second success=" + second.Success + " id=" + secondId + " err=" + (second.Error ?? ""));
            DumpTables(s, path, "after-second");

            string body = ReadBody(s, id);
            Console.WriteLine("has甲=" + (body.IndexOf(ConsA, StringComparison.Ordinal) >= 0)
                + " has乙=" + (body.IndexOf(ConsB, StringComparison.Ordinal) >= 0)
                + " has丙=" + (body.IndexOf(ConsC, StringComparison.Ordinal) >= 0)
                + " has丁=" + (body.IndexOf(ConsD, StringComparison.Ordinal) >= 0));
            Console.WriteLine("T_ count=" + Regex.Matches(body, @"\[T_[0-9a-zA-Z]{5,}-start\]").Count);
            Console.WriteLine("doc left open: " + path);
        }

        private static void DumpTables(WdSession s, string path, string tag)
        {
            s.OnWordDoc(path, doc =>
            {
                Console.WriteLine(tag + " Tables.Count=" + doc.Tables.Count
                    + " Paras=" + doc.Paragraphs.Count
                    + " Content.End=" + doc.Content.End);
                for (int i = 1; i <= doc.Tables.Count; i++)
                {
                    Word.Table t = doc.Tables[i];
                    string title = "";
                    try { title = t.Title ?? ""; } catch { title = "?"; }
                    string cells = "";
                    try
                    {
                        cells = (t.Cell(1, 1).Range.Text ?? "").Replace("\r", "\\r").Replace("\a", "\\a")
                            + " | "
                            + (t.Cell(1, 2).Range.Text ?? "").Replace("\r", "\\r").Replace("\a", "\\a");
                    }
                    catch (Exception ex)
                    {
                        cells = ex.Message;
                    }

                    Console.WriteLine("  T" + i + " rows=" + t.Rows.Count + " cols=" + t.Columns.Count
                        + " rng=" + t.Range.Start + "-" + t.Range.End
                        + " title=" + title
                        + " r1=" + cells);
                }

                if (doc.Tables.Count >= 2)
                {
                    int gap = doc.Tables[2].Range.Start - doc.Tables[1].Range.End;
                    Console.WriteLine("  gap(T2.Start-T1.End)=" + gap);
                }

                return true;
            });
        }

        public static void RunCreateAfterTwoOnly(
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
                RunAfterTwo(run, s, "wt-create-after-first", AfterSent, AfterSentB, AfterSent);
                RunAfterTwo(run, s, "wt-create-after-last", AfterSent, AfterSentB, AfterSentB);
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

        private static void RunAfterTwo(
            TestRun run,
            WdSession s,
            string caseId,
            string first,
            string second,
            string anchorNeedle)
        {
            string path, id;
            if (!TryOpen(run, s, out path, out id, caseId))
            {
                return;
            }

            if (!s.TypePlain(path, first + "\r" + second))
            {
                FailMany(run, "COM 写两句失败", caseId);
                s.CloseDocByPath(path);
                return;
            }

            s.ActivateDoc(path);
            ToolResult got = s.GetContent(id);
            string body = WdSession.Str(got, "document_content");
            string sent = CodeContaining(body, anchorNeedle, "S");
            if (string.IsNullOrEmpty(sent))
            {
                FailMany(run, "无可抄含锚句的 S_", caseId);
                s.CloseDocByPath(path);
                return;
            }

            var args = CreateArgs(id, WriteXml(GridXml(LabelA, "", LabelB, "")));
            args["target"] = new Dictionary<string, object> { ["codes"] = sent };
            ToolResult after = s.CreateTableFromXml(args);
            string afterBody = ReadBody(s, id);
            int iFirst = IndexOf(afterBody, first);
            int iSecond = IndexOf(afterBody, second);
            int iTbl = IndexOf(afterBody, LabelA);
            bool order = string.Equals(anchorNeedle, first, StringComparison.Ordinal)
                ? iFirst >= 0 && iTbl > iFirst && iSecond > iTbl
                : iFirst >= 0 && iSecond > iFirst && iTbl > iSecond;
            string reason;
            if (!after.Success)
            {
                reason = after.Error ?? "建表失败";
            }
            else if (!order)
            {
                reason = string.Equals(anchorNeedle, first, StringComparison.Ordinal)
                    ? "表没夹在两句之间（可能误用了全文 Range）"
                    : "表没落在第二句之后";
            }
            else
            {
                reason = "";
            }

            Ok(run, caseId, after.Success && order, reason);
            s.CloseDocByPath(path);
        }

        private static int IndexOf(string text, string needle)
        {
            return string.IsNullOrEmpty(text) || string.IsNullOrEmpty(needle)
                ? -1
                : text.IndexOf(needle, StringComparison.Ordinal);
        }

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
            string warmPath = WarmWpsChannel(s);
            try
            {
                RunCreateAndFill(run, s);
                RunFormatDeleteOpt(run, s);
                RunConsolidate(run, s);
                RunNegLate(run, s);
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
            s.CreateTableFromXml(CreateArgs(id, WriteXml(GridXml(LabelA, "", LabelB, ""))));
            return path;
        }

        private static void BurnFirstWpsOpen(WdSession s)
        {
            if (!s.IsWps)
            {
                return;
            }

            string path = s.NewDocxPath();
            ToolResult opened = s.Open(path, true, s.App);
            if (!opened.Success)
            {
                return;
            }

            string id = WdSession.Str(opened, "channel_id");
            s.ActivateDoc(path);
            s.CreateTableFromXml(CreateArgs(id, WriteXml(GridXml(LabelA, "", LabelB, ""))));
            s.CloseDocByPath(path);
        }

        private static void RunNegEarly(TestRun run, WdSession s)
        {
            string xml = WriteXml(GridXml(LabelA, "", LabelB, ""));
            ToolResult noCh = s.CreateTableFromXml(CreateArgs(null, xml));
            if (!noCh.Success
                && (Contains(noCh.Error, "打开")
                    || Contains(noCh.Error, "unsupported")
                    || Contains(noCh.Error, "browser")
                    || Contains(noCh.Error, "渠道")))
            {
                run.CaseOk("wt-neg-no-channel");
            }
            else if (noCh.Success)
            {
                run.CaseSkip("wt-neg-no-channel", "现场已有默认文字渠，不关用户稿");
            }
            else
            {
                run.CaseFail("wt-neg-no-channel", noCh.Error ?? "应失败或跳过");
            }

            ExpectFail(run, "wt-neg-unknown-id", s.CreateTableFromXml(CreateArgs("w999", xml)), "未知");
            ExpectFail(
                run,
                "wt-neg-create-no-path",
                s.CreateTableFromXml(new Dictionary<string, object>()),
                "path");
            ExpectFail(
                run,
                "wt-neg-create-missing",
                s.CreateTableFromXml(CreateArgs(null, Path.Combine(Path.GetTempPath(), "ew-wt-missing-" + Guid.NewGuid().ToString("N") + ".xml"))),
                "不存在");
            ExpectFail(
                run,
                "wt-neg-create-bad-xml",
                s.CreateTableFromXml(CreateArgs(null, WriteXml("<TableConfig><Table><Properties/></Table><Data><tr><td>x</td></tr></Data></TableConfig>"))),
                "Data");
            ExpectFail(
                run,
                "wt-neg-create-pos",
                s.CreateTableFromXml(new Dictionary<string, object>
                {
                    ["path"] = xml,
                    ["position"] = "end"
                }),
                "position");
            ExpectFail(
                run,
                "wt-neg-del-no-id",
                s.DeleteTable(new Dictionary<string, object>()),
                "编号");
            ExpectFail(
                run,
                "wt-neg-read-no-id",
                s.ReadTableRowValues(new Dictionary<string, object>()),
                "table_id");
            ExpectFail(
                run,
                "wt-neg-apply-no-id",
                s.ApplyTableRowValues(new Dictionary<string, object>
                {
                    ["format"] = DefaultFormat(),
                    ["writes"] = new object[]
                    {
                        new Dictionary<string, object> { ["row"] = 1, ["values"] = new[] { FillA } }
                    }
                }),
                "table_id");
            ExpectFail(
                run,
                "wt-neg-apply-no-fmt",
                s.ApplyTableRowValues(new Dictionary<string, object>
                {
                    ["table_id"] = "T_00001",
                    ["writes"] = new object[]
                    {
                        new Dictionary<string, object> { ["row"] = 1, ["values"] = new[] { FillA } }
                    }
                }),
                "char_format");
            ExpectFail(
                run,
                "wt-neg-apply-inherit",
                s.ApplyTableRowValues(new Dictionary<string, object>
                {
                    ["table_id"] = "T_00001",
                    ["format"] = DefaultFormat(),
                    ["inherit_from"] = "S_00001",
                    ["writes"] = new object[]
                    {
                        new Dictionary<string, object> { ["row"] = 1, ["values"] = new[] { FillA } }
                    }
                }),
                "inherit");
            ExpectFail(
                run,
                "wt-neg-apply-no-writes",
                s.ApplyTableRowValues(new Dictionary<string, object>
                {
                    ["table_id"] = "T_00001",
                    ["format"] = DefaultFormat()
                }),
                "writes");
            ExpectFail(
                run,
                "wt-neg-apply-anchor",
                s.ApplyTableRowValues(new Dictionary<string, object>
                {
                    ["table_id"] = "T_00001",
                    ["format"] = DefaultFormat(),
                    ["writes"] = new object[]
                    {
                        new Dictionary<string, object>
                        {
                            ["row"] = 1,
                            ["values"] = new[] { FillA },
                            ["anchor"] = "指标"
                        }
                    }
                }),
                "anchor");
            ExpectFail(
                run,
                "wt-neg-fmt-no-xml",
                s.ApplyTableFormatFromXml(new Dictionary<string, object> { ["table_id"] = "T_00001" }),
                "path");
            ExpectFail(
                run,
                "wt-neg-opt-no-id",
                s.OptimizeTableDisplay(new Dictionary<string, object>()),
                "编号");
        }

        private static void RunCreateAndFill(TestRun run, WdSession s)
        {
            BurnFirstWpsOpen(s);

            string path, id;
            if (!TryOpen(run, s, out path, out id,
                "wt-create-ok", "wt-create-id", "wt-create-text", "wt-create-fields",
                "wt-read-ok", "wt-read-empty", "wt-fill-ok", "wt-fill-text", "wt-fill-keep",
                "wt-fill-count", "wt-fill-wrong-len"))
            {
                return;
            }

            s.ActivateDoc(path);
            s.GetContent(id);
            ToolResult created = s.CreateTableFromXml(CreateArgs(id, WriteXml(GridXml(LabelA, "", LabelB, ""))));
            string tableId = WdSession.Str(created, "table_id");
            Ok(run, "wt-create-ok", created.Success, created.Error ?? "建表失败");
            Ok(run, "wt-create-id",
                created.Success && tableId.StartsWith("T_", StringComparison.Ordinal),
                created.Success ? "table_id=" + tableId : (created.Error ?? ""));
            Ok(run, "wt-create-text",
                created.Success && StillHas(s, id, LabelA) && StillHas(s, id, LabelB),
                created.Success ? "建表正文无标签" : (created.Error ?? ""));
            Ok(run, "wt-create-fields",
                created.Success
                    && WdSession.Bool(created, "table_created")
                    && WdSession.IntOf(created, "rows") == 2
                    && WdSession.IntOf(created, "columns") == 2,
                created.Success ? "回包缺 table_created / 行列" : (created.Error ?? ""));
            if (!created.Success || string.IsNullOrEmpty(tableId))
            {
                FailMany(run, created.Error ?? "无 table_id",
                    "wt-read-ok", "wt-read-empty", "wt-fill-ok", "wt-fill-text", "wt-fill-keep",
                    "wt-fill-count", "wt-fill-wrong-len");
                s.CloseDocByPath(path);
                return;
            }

            ToolResult read = s.ReadTableRowValues(new Dictionary<string, object>
            {
                ["channel_id"] = id,
                ["table_id"] = tableId
            });
            JObject readObj = DataObj(read);
            Ok(run, "wt-read-ok",
                read.Success
                    && WdSession.Bool(read, "do_not_use_getcontent_table_html")
                    && readObj != null
                    && readObj["rows"] is JArray,
                read.Success ? "回包无 do_not_use / rows" : (read.Error ?? ""));

            int emptyRow;
            int emptyCount;
            bool hasEmpty = TryFirstEmptyRow(read, out emptyRow, out emptyCount);
            Ok(run, "wt-read-empty", hasEmpty && emptyCount > 0,
                read.Success ? "没有 empty_count>0 的行" : (read.Error ?? ""));
            if (!hasEmpty)
            {
                FailMany(run, "无空槽可填",
                    "wt-fill-ok", "wt-fill-text", "wt-fill-keep", "wt-fill-count", "wt-fill-wrong-len");
                s.CloseDocByPath(path);
                return;
            }

            var tooLong = new string[emptyCount + 1];
            for (int i = 0; i < tooLong.Length; i++)
            {
                tooLong[i] = "x";
            }

            ToolResult badLen = s.ApplyTableRowValues(FillArgs(id, tableId, emptyRow, tooLong));
            Ok(run, "wt-fill-wrong-len",
                !badLen.Success && StillHas(s, id, LabelA) && !StillHas(s, id, FillA),
                badLen.Success ? "values 个数不对仍写入" : (badLen.Error ?? "错长度却改了字"));

            var values = new string[emptyCount];
            values[0] = FillA;
            for (int i = 1; i < values.Length; i++)
            {
                values[i] = "";
            }

            ToolResult fill = s.ApplyTableRowValues(FillArgs(id, tableId, emptyRow, values));
            Ok(run, "wt-fill-ok", fill.Success, fill.Error ?? "填空失败");
            Ok(run, "wt-fill-text",
                fill.Success && StillHas(s, id, FillA),
                fill.Success ? "空槽未写出" : (fill.Error ?? ""));
            Ok(run, "wt-fill-keep",
                fill.Success && StillHas(s, id, LabelA) && StillHas(s, id, LabelB),
                fill.Success ? "填空改了标签格" : (fill.Error ?? ""));
            Ok(run, "wt-fill-count",
                fill.Success && WdSession.IntOf(fill, "cells_written") >= 1,
                fill.Success ? "cells_written=" + WdSession.IntOf(fill, "cells_written") : (fill.Error ?? ""));
            s.CloseDocByPath(path);
        }

        private static void RunFormatDeleteOpt(TestRun run, WdSession s)
        {
            string path, id;
            if (!TryOpen(run, s, out path, out id,
                "wt-create-after", "wt-create-start", "wt-create-merge",
                "wt-xfmt-ok", "wt-xfmt-file", "wt-afmt-ok", "wt-afmt-text", "wt-afmt-no-data",
                "wt-del-ok", "wt-del-gone", "wt-rebuild-ok", "wt-rebuild-text",
                "wt-opt-ok", "wt-opt-fields"))
            {
                return;
            }

            if (!s.TypePlain(path, AfterSent))
            {
                FailMany(run, "COM 写句失败",
                    "wt-create-after", "wt-create-start", "wt-create-merge",
                    "wt-xfmt-ok", "wt-xfmt-file", "wt-afmt-ok", "wt-afmt-text", "wt-afmt-no-data",
                    "wt-del-ok", "wt-del-gone", "wt-rebuild-ok", "wt-rebuild-text",
                    "wt-opt-ok", "wt-opt-fields");
                s.CloseDocByPath(path);
                return;
            }

            ToolResult got = s.GetContent(id);
            string sent = CodeContaining(WdSession.Str(got, "document_content"), AfterSent, "S");
            if (string.IsNullOrEmpty(sent))
            {
                FailMany(run, "无可抄含前句的 S_",
                    "wt-create-after", "wt-create-start", "wt-create-merge",
                    "wt-xfmt-ok", "wt-xfmt-file", "wt-afmt-ok", "wt-afmt-text", "wt-afmt-no-data",
                    "wt-del-ok", "wt-del-gone", "wt-rebuild-ok", "wt-rebuild-text",
                    "wt-opt-ok", "wt-opt-fields");
                s.CloseDocByPath(path);
                return;
            }

            var afterArgs = CreateArgs(id, WriteXml(GridXml(LabelA, "", LabelB, "")));
            afterArgs["target"] = new Dictionary<string, object> { ["codes"] = sent };
            ToolResult after = s.CreateTableFromXml(afterArgs);
            string afterId = WdSession.Str(after, "table_id");
            Ok(run, "wt-create-after",
                after.Success && StillHas(s, id, AfterSent) && StillHas(s, id, LabelA),
                after.Success ? "句后建表丢了句或表" : (after.Error ?? ""));

            string startPath = s.NewDocxPath();
            ToolResult startOpen = s.Open(startPath, true, s.App);
            if (!startOpen.Success)
            {
                FailMany(run, startOpen.Error ?? "再建空稿失败", "wt-create-start");
            }
            else
            {
                string startId = WdSession.Str(startOpen, "channel_id");
                var startArgs = CreateArgs(startId, WriteXml(GridXml(LabelA, "1", LabelB, "2")));
                startArgs["target"] = new Dictionary<string, object> { ["position"] = "start" };
                ToolResult start = s.CreateTableFromXml(startArgs);
                Ok(run, "wt-create-start",
                    start.Success && StillHas(s, startId, LabelA),
                    start.Success ? "文首建表无标签" : (start.Error ?? ""));
                s.CloseDocByPath(startPath);
            }

            string mergePath = s.NewDocxPath();
            ToolResult mergeOpen = s.Open(mergePath, true, s.App);
            if (!mergeOpen.Success)
            {
                FailMany(run, mergeOpen.Error ?? "再建空稿失败", "wt-create-merge");
            }
            else
            {
                string mergeId = WdSession.Str(mergeOpen, "channel_id");
                s.ActivateDoc(mergePath);
                ToolResult merge = s.CreateTableFromXml(CreateArgs(mergeId, WriteXml(MergeXml())));
                Ok(run, "wt-create-merge",
                    merge.Success
                        && StillHas(s, mergeId, MergeHead)
                        && StillHas(s, mergeId, MergeLeft)
                        && StillHas(s, mergeId, MergeRight),
                    merge.Success ? "合并表缺可见格" : (merge.Error ?? ""));
                s.CloseDocByPath(mergePath);
            }

            if (!after.Success || string.IsNullOrEmpty(afterId))
            {
                FailMany(run, after.Error ?? "句后表无 id",
                    "wt-xfmt-ok", "wt-xfmt-file", "wt-afmt-ok", "wt-afmt-text", "wt-afmt-no-data",
                    "wt-del-ok", "wt-del-gone", "wt-rebuild-ok", "wt-rebuild-text",
                    "wt-opt-ok", "wt-opt-fields");
                s.CloseDocByPath(path);
                return;
            }

            string extractPath = Path.Combine(Path.GetTempPath(), "ew-wt-fmt-" + Guid.NewGuid().ToString("N") + ".xml");
            ToolResult extract = s.ExtractTableFormat(new Dictionary<string, object>
            {
                ["channel_id"] = id,
                ["path"] = extractPath,
                ["table_selection"] = new Dictionary<string, object>
                {
                    ["type"] = "by_id",
                    ["table_id"] = afterId
                }
            });
            Ok(run, "wt-xfmt-ok", extract.Success, extract.Error ?? "抽格式失败");
            Ok(run, "wt-xfmt-file",
                extract.Success
                    && (WdSession.Bool(extract, "format_saved")
                        || !string.IsNullOrEmpty(WdSession.Str(extract, "format_content"))),
                extract.Success ? "无 format_saved / format_content" : (extract.Error ?? ""));

            ToolResult applyFmt = s.ApplyTableFormatFromXml(new Dictionary<string, object>
            {
                ["channel_id"] = id,
                ["table_id"] = afterId,
                ["xml_content"] = "<TableConfig><General><TableFontColor>000000</TableFontColor><Style><TableStyle>ThreeLine</TableStyle><HeaderRows>1</HeaderRows></Style></General></TableConfig>"
            });
            Ok(run, "wt-afmt-ok", applyFmt.Success, applyFmt.Error ?? "套表格式失败");
            Ok(run, "wt-afmt-text",
                applyFmt.Success && StillHas(s, id, LabelA) && StillHas(s, id, AfterSent),
                applyFmt.Success ? "套格式改了字" : (applyFmt.Error ?? ""));

            ToolResult applyData = s.ApplyTableFormatFromXml(new Dictionary<string, object>
            {
                ["channel_id"] = id,
                ["table_id"] = afterId,
                ["xml_content"] = GridXml("不该写入甲", "不该写入乙", "不该写入丙", "不该写入丁")
            });
            Ok(run, "wt-afmt-no-data",
                applyData.Success
                    && StillHas(s, id, LabelA)
                    && !StillHas(s, id, "不该写入甲"),
                applyData.Success ? "套格式吃了 Data" : (applyData.Error ?? ""));

            ToolResult opt = s.OptimizeTableDisplay(new Dictionary<string, object>
            {
                ["channel_id"] = id,
                ["table_id"] = afterId
            });
            Ok(run, "wt-opt-ok", opt.Success, opt.Error ?? "新建后第一次 optimize 失败");
            Ok(run, "wt-opt-fields",
                opt.Success && WdSession.HasKey(opt, "final_font_size"),
                opt.Success ? "回包无 final_font_size" : (opt.Error ?? ""));

            ToolResult del = s.DeleteTable(new Dictionary<string, object>
            {
                ["channel_id"] = id,
                ["table_id"] = afterId
            });
            Ok(run, "wt-del-ok", del.Success, del.Error ?? "删表失败");
            Ok(run, "wt-del-gone",
                del.Success && !StillHas(s, id, LabelA),
                del.Success ? "删表后标签还在" : (del.Error ?? ""));

            ToolResult rebuild = s.CreateTableFromXml(CreateArgs(id, WriteXml(GridXml(RebuildA, "1", RebuildB, "2"))));
            Ok(run, "wt-rebuild-ok", rebuild.Success, rebuild.Error ?? "删后重建失败");
            Ok(run, "wt-rebuild-text",
                rebuild.Success && StillHas(s, id, RebuildA) && StillHas(s, id, AfterSent),
                rebuild.Success ? "重建表或前句丢了" : (rebuild.Error ?? ""));
            s.CloseDocByPath(path);
        }

        private static void RunConsolidate(TestRun run, WdSession s)
        {
            string path, id;
            if (!TryOpen(run, s, out path, out id, ExtraIds))
            {
                return;
            }

            if (!s.TypePlain(path, ConsSent))
            {
                FailMany(run, "COM 写巩固句失败", ExtraIds);
                s.CloseDocByPath(path);
                return;
            }

            s.ActivateDoc(path);
            s.GetContent(id);

            ExpectFail(run, "wt-create-root",
                s.CreateTableFromXml(CreateArgs(id, WriteXml("<NotTable><Table><Data><tr><td>x</td></tr></Data></Table></NotTable>"))),
                "TableConfig");
            ExpectFail(run, "wt-create-colw-gen",
                s.CreateTableFromXml(CreateArgs(id, WriteXml(
                    "<TableConfig><General><ColWidths><float>120</float></ColWidths></General><Table><Properties/><Data><tr><td>x</td></tr></Data></Table></TableConfig>"))),
                "ColWidths");
            ExpectFail(run, "wt-create-colw-tbl",
                s.CreateTableFromXml(CreateArgs(id, WriteXml(
                    "<TableConfig><General/><Table><ColWidths><float>120</float></ColWidths><Data><tr><td>x</td></tr></Data></Table></TableConfig>"))),
                "ColWidths");
            ExpectFail(run, "wt-create-style-prop",
                s.CreateTableFromXml(CreateArgs(id, WriteXml(
                    "<TableConfig><General/><Table><Properties><Style><TableStyle>ThreeLine</TableStyle></Style></Properties><Data><tr><td>x</td></tr></Data></Table></TableConfig>"))),
                "Style");
            ExpectFail(run, "wt-create-no-data",
                s.CreateTableFromXml(CreateArgs(id, WriteXml(
                    "<TableConfig><General/><Table><Properties><ColWidths><float>120</float></ColWidths></Properties></Table></TableConfig>"))),
                "Data");

            var endArgs = CreateArgs(id, WriteXml(GridXml(ConsA, "", ConsB, "")));
            endArgs["target"] = new Dictionary<string, object> { ["position"] = "end" };
            ExpectFail(run, "wt-create-pos-end", s.CreateTableFromXml(endArgs), "start");

            var unkCode = CreateArgs(id, WriteXml(GridXml(ConsA, "", ConsB, "")));
            unkCode["target"] = new Dictionary<string, object> { ["codes"] = "S_99999" };
            ExpectFail(run, "wt-create-unk-code", s.CreateTableFromXml(unkCode), "locate");

            var unkTbl = CreateArgs(id, WriteXml(GridXml(ConsA, "", ConsB, "")));
            unkTbl["target"] = new Dictionary<string, object> { ["table"] = "T_99999" };
            ExpectFail(run, "wt-create-unk-tbl", s.CreateTableFromXml(unkTbl), "表格");

            ToolResult created = s.CreateTableFromXml(
                CreateArgsFilename(id, WriteXml(GridXml(ConsA, "", ConsB, ""))));
            string tableId = WdSession.Str(created, "table_id");
            Ok(run, "wt-create-filename",
                created.Success && StillHas(s, id, ConsA) && StillHas(s, id, ConsB),
                created.Success ? "filename 建表无标签" : (created.Error ?? ""));
            if (!created.Success || string.IsNullOrEmpty(tableId))
            {
                FailMany(run, created.Error ?? "filename 建表无 table_id",
                    "wt-create-1x1", "wt-create-3row", "wt-create-span", "wt-create-rowspan",
                    "wt-create-two", "wt-create-two-text", "wt-create-after-tbl",
                    "wt-fill-short-len", "wt-fill-index", "wt-fill-bad-row", "wt-fill-row0",
                    "wt-fill-writes-empty", "wt-fill-writes-obj", "wt-fill-unk-t", "wt-read-unk-t",
                    "wt-read-slots", "wt-read-rows-n", "wt-fill-blank", "wt-fill-row2",
                    "wt-fill-row2-text", "wt-read-after-fill", "wt-fill-no-empty", "wt-fill-del",
                    "wt-xfmt-no-sel", "wt-xfmt-ctx", "wt-xfmt-unk", "wt-xfmt-no-type",
                    "wt-afmt-no-id", "wt-afmt-unk", "wt-afmt-bad-xml", "wt-afmt-path",
                    "wt-afmt-applied", "wt-afmt-echo", "wt-del-unk", "wt-del-twice",
                    "wt-opt-unk", "wt-opt-after-del", "wt-opt-rebuild", "wt-del-keep-sent",
                    "wt-read-gone", "wt-two-del-first");
                s.CloseDocByPath(path);
                return;
            }

            var afterTbl = CreateArgs(id, WriteXml(GridXml(ConsC, "", ConsD, "")));
            afterTbl["target"] = new Dictionary<string, object> { ["table"] = tableId };
            ToolResult second = s.CreateTableFromXml(afterTbl);
            string secondId = WdSession.Str(second, "table_id");
            Ok(run, "wt-create-after-tbl",
                second.Success && StillHas(s, id, ConsA) && StillHas(s, id, ConsC),
                second.Success ? "表后建表丢了表" : (second.Error ?? ""));
            Ok(run, "wt-create-two",
                second.Success && secondId.StartsWith("T_", StringComparison.Ordinal) && secondId != tableId,
                second.Success ? "第二张 table_id 不对" : (second.Error ?? ""));
            Ok(run, "wt-create-two-text",
                second.Success && StillHas(s, id, ConsC) && StillHas(s, id, ConsD) && StillHas(s, id, ConsA),
                second.Success ? "两张表标签不全" : (second.Error ?? ""));

            Ok(run, "wt-create-1x1",
                CreateOkHas(s, id, OneCellXml(Cell1), Cell1),
                "单格未出现");
            Ok(run, "wt-create-3row",
                CreateOkHas(s, id, ThreeRowXml(), Row3A) && StillHas(s, id, Row3B) && StillHas(s, id, Row3C),
                "三行标签不全");
            Ok(run, "wt-create-span",
                CreateOkHas(s, id, SpanXml(), SpanTxt),
                "span 格字未出现");
            Ok(run, "wt-create-rowspan",
                CreateOkHas(s, id, RowspanXml(), RsLeft) && StillHas(s, id, RsUp) && StillHas(s, id, RsDown),
                "跨行合并可见格不全");

            ToolResult read = s.ReadTableRowValues(new Dictionary<string, object>
            {
                ["channel_id"] = id,
                ["table_id"] = tableId
            });
            JObject readObj = DataObj(read);
            JArray rows = readObj == null ? null : readObj["rows"] as JArray;
            Ok(run, "wt-read-slots",
                read.Success && rows != null && RowHasSlots(rows),
                read.Success ? "rows 无 slots" : (read.Error ?? ""));
            Ok(run, "wt-read-rows-n",
                read.Success && rows != null && rows.Count == 2,
                read.Success ? "rows.Count=" + (rows == null ? -1 : rows.Count) : (read.Error ?? ""));

            int emptyRow;
            int emptyCount;
            bool hasEmpty = TryFirstEmptyRow(read, out emptyRow, out emptyCount);
            int row2;
            int empty2;
            bool hasRow2 = TryLastEmptyRow(read, out row2, out empty2);

            ExpectFail(run, "wt-fill-unk-t",
                s.ApplyTableRowValues(FillArgs(id, "T_99999", 1, new[] { "x" })),
                "找不到");
            ExpectFail(run, "wt-read-unk-t",
                s.ReadTableRowValues(new Dictionary<string, object>
                {
                    ["channel_id"] = id,
                    ["table_id"] = "T_99999"
                }),
                "找不到");

            if (!hasEmpty)
            {
                FailMany(run, "巩固表无空槽",
                    "wt-fill-short-len", "wt-fill-index", "wt-fill-bad-row", "wt-fill-row0",
                    "wt-fill-writes-empty", "wt-fill-writes-obj", "wt-fill-blank", "wt-fill-row2",
                    "wt-fill-row2-text", "wt-read-after-fill", "wt-fill-no-empty");
            }
            else
            {
                var tooShort = emptyCount <= 1 ? new string[0] : new string[emptyCount - 1];
                for (int i = 0; i < tooShort.Length; i++)
                {
                    tooShort[i] = "x";
                }

                ToolResult shortLen = s.ApplyTableRowValues(FillArgs(id, tableId, emptyRow, tooShort));
                Ok(run, "wt-fill-short-len",
                    !shortLen.Success && StillHas(s, id, ConsA) && !StillHas(s, id, FillA),
                    shortLen.Success ? "values 偏短仍写入" : (shortLen.Error ?? "错长度却改了字"));

                var indexArgs = FillArgs(id, tableId, emptyRow, PadValues(emptyCount, FillA));
                ((Dictionary<string, object>)((object[])indexArgs["writes"])[0])["index"] = 1;
                ExpectFail(run, "wt-fill-index", s.ApplyTableRowValues(indexArgs), "index");
                ExpectFail(run, "wt-fill-bad-row",
                    s.ApplyTableRowValues(FillArgs(id, tableId, 99, new[] { "x" })),
                    "invalid_row");
                ExpectFail(run, "wt-fill-row0",
                    s.ApplyTableRowValues(FillArgs(id, tableId, 0, new[] { "x" })),
                    "invalid_row");
                ExpectFail(run, "wt-fill-writes-empty",
                    s.ApplyTableRowValues(new Dictionary<string, object>
                    {
                        ["channel_id"] = id,
                        ["table_id"] = tableId,
                        ["format"] = DefaultFormat(),
                        ["writes"] = new object[0]
                    }),
                    "writes");
                ExpectFail(run, "wt-fill-writes-obj",
                    s.ApplyTableRowValues(new Dictionary<string, object>
                    {
                        ["channel_id"] = id,
                        ["table_id"] = tableId,
                        ["format"] = DefaultFormat(),
                        ["writes"] = "nope"
                    }),
                    "数组");

                var blanks = new string[emptyCount];
                for (int i = 0; i < blanks.Length; i++)
                {
                    blanks[i] = "";
                }

                ToolResult blank = s.ApplyTableRowValues(FillArgs(id, tableId, emptyRow, blanks));
                Ok(run, "wt-fill-blank",
                    blank.Success && StillHas(s, id, ConsA) && StillHas(s, id, ConsB) && !StillHas(s, id, FillA),
                    blank.Success ? "全空写入改了标签或写出了数" : (blank.Error ?? ""));

                int fillRow = hasRow2 ? row2 : emptyRow;
                int fillN = hasRow2 ? empty2 : emptyCount;
                ToolResult fill2 = s.ApplyTableRowValues(FillArgs(id, tableId, fillRow, PadValues(fillN, FillB)));
                Ok(run, "wt-fill-row2", fill2.Success, fill2.Error ?? "第二空行填空失败");
                Ok(run, "wt-fill-row2-text",
                    fill2.Success && StillHas(s, id, FillB) && !StillHas(s, id, FillA) && StillHas(s, id, ConsA),
                    fill2.Success ? "第二空行字不对" : (fill2.Error ?? ""));

                ToolResult readAfter = s.ReadTableRowValues(new Dictionary<string, object>
                {
                    ["channel_id"] = id,
                    ["table_id"] = tableId
                });
                int afterN;
                TryRowEmptyCount(readAfter, fillRow, out afterN);
                Ok(run, "wt-read-after-fill",
                    readAfter.Success && afterN == 0,
                    readAfter.Success ? "填后 empty_count=" + afterN : (readAfter.Error ?? ""));
                ExpectFail(run, "wt-fill-no-empty",
                    s.ApplyTableRowValues(FillArgs(id, tableId, fillRow, new[] { "x" })),
                    "values_length");
            }

            ExpectFail(run, "wt-xfmt-no-sel",
                s.ExtractTableFormat(new Dictionary<string, object>
                {
                    ["channel_id"] = id,
                    ["path"] = Path.Combine(Path.GetTempPath(), "ew-wt-x-" + Guid.NewGuid().ToString("N") + ".xml")
                }),
                "table_selection");
            ExpectFail(run, "wt-xfmt-ctx",
                s.ExtractTableFormat(new Dictionary<string, object>
                {
                    ["channel_id"] = id,
                    ["path"] = Path.Combine(Path.GetTempPath(), "ew-wt-x-" + Guid.NewGuid().ToString("N") + ".xml"),
                    ["table_selection"] = new Dictionary<string, object>
                    {
                        ["type"] = "by_context",
                        ["table_id"] = tableId
                    }
                }),
                "by_context");
            ExpectFail(run, "wt-xfmt-unk",
                s.ExtractTableFormat(new Dictionary<string, object>
                {
                    ["channel_id"] = id,
                    ["path"] = Path.Combine(Path.GetTempPath(), "ew-wt-x-" + Guid.NewGuid().ToString("N") + ".xml"),
                    ["table_selection"] = new Dictionary<string, object>
                    {
                        ["type"] = "by_id",
                        ["table_id"] = "T_99999"
                    }
                }),
                "找不到");
            ExpectFail(run, "wt-xfmt-no-type",
                s.ExtractTableFormat(new Dictionary<string, object>
                {
                    ["channel_id"] = id,
                    ["path"] = Path.Combine(Path.GetTempPath(), "ew-wt-x-" + Guid.NewGuid().ToString("N") + ".xml"),
                    ["table_selection"] = new Dictionary<string, object> { ["table_id"] = tableId }
                }),
                "类型");

            ExpectFail(run, "wt-afmt-no-id",
                s.ApplyTableFormatFromXml(new Dictionary<string, object>
                {
                    ["channel_id"] = id,
                    ["xml_content"] = "<TableConfig><General><TableFontColor>000000</TableFontColor></General></TableConfig>"
                }),
                "编号");
            ExpectFail(run, "wt-afmt-unk",
                s.ApplyTableFormatFromXml(new Dictionary<string, object>
                {
                    ["channel_id"] = id,
                    ["table_id"] = "T_99999",
                    ["xml_content"] = "<TableConfig><General><TableFontColor>000000</TableFontColor></General></TableConfig>"
                }),
                "找不到");
            ExpectFail(run, "wt-afmt-bad-xml",
                s.ApplyTableFormatFromXml(new Dictionary<string, object>
                {
                    ["channel_id"] = id,
                    ["table_id"] = tableId,
                    ["xml_content"] = "<NotTable/>"
                }),
                "TableConfig");

            string fmtPath = WriteXml(
                "<TableConfig><General><TableFontColor>000000</TableFontColor><Style><TableStyle>ThreeLine</TableStyle><HeaderRows>1</HeaderRows></Style></General></TableConfig>");
            ToolResult afmt = s.ApplyTableFormatFromXml(new Dictionary<string, object>
            {
                ["channel_id"] = id,
                ["table_id"] = tableId,
                ["path"] = fmtPath
            });
            Ok(run, "wt-afmt-path",
                afmt.Success && StillHas(s, id, ConsA) && StillHas(s, id, ConsSent),
                afmt.Success ? "path 套格式改了字" : (afmt.Error ?? ""));
            Ok(run, "wt-afmt-applied",
                afmt.Success && WdSession.Bool(afmt, "table_format_applied"),
                afmt.Success ? "无 table_format_applied" : (afmt.Error ?? ""));
            Ok(run, "wt-afmt-echo",
                afmt.Success && Contains(WdSession.Str(afmt, "table_id"), tableId),
                afmt.Success ? "回包 table_id 不对" : (afmt.Error ?? ""));

            ExpectFail(run, "wt-del-unk",
                s.DeleteTable(new Dictionary<string, object> { ["channel_id"] = id, ["table_id"] = "T_99999" }),
                "找不到");
            ExpectFail(run, "wt-opt-unk",
                s.OptimizeTableDisplay(new Dictionary<string, object> { ["channel_id"] = id, ["table_id"] = "T_99999" }),
                "找不到");

            ToolResult del = s.DeleteTable(new Dictionary<string, object>
            {
                ["channel_id"] = id,
                ["table_id"] = tableId
            });
            Ok(run, "wt-two-del-first",
                del.Success && !StillHas(s, id, ConsA) && StillHas(s, id, ConsC),
                del.Success ? "删第一张时第二张也没了或第一张还在" : (del.Error ?? ""));
            Ok(run, "wt-del-keep-sent",
                del.Success && StillHas(s, id, ConsSent),
                del.Success ? "删表带走了巩固句" : (del.Error ?? ""));
            ExpectFail(run, "wt-read-gone",
                s.ReadTableRowValues(new Dictionary<string, object>
                {
                    ["channel_id"] = id,
                    ["table_id"] = tableId
                }),
                "找不到");
            ExpectFail(run, "wt-del-twice",
                s.DeleteTable(new Dictionary<string, object> { ["channel_id"] = id, ["table_id"] = tableId }),
                "找不到");
            ExpectFail(run, "wt-opt-after-del",
                s.OptimizeTableDisplay(new Dictionary<string, object> { ["channel_id"] = id, ["table_id"] = tableId }),
                "找不到");
            ExpectFail(run, "wt-fill-del",
                s.ApplyTableRowValues(FillArgs(id, tableId, 1, new[] { "x" })),
                "找不到");

            ToolResult rebuild = s.CreateTableFromXml(CreateArgs(id, WriteXml(GridXml(RebuildA, "1", RebuildB, "2"))));
            string newId = WdSession.Str(rebuild, "table_id");
            if (!rebuild.Success || string.IsNullOrEmpty(newId))
            {
                FailMany(run, rebuild.Error ?? "巩固重建失败", "wt-opt-rebuild");
            }
            else
            {
                ToolResult opt = s.OptimizeTableDisplay(new Dictionary<string, object>
                {
                    ["channel_id"] = id,
                    ["table_id"] = newId
                });
                Ok(run, "wt-opt-rebuild", opt.Success, opt.Error ?? "新表第一次 optimize 失败");
            }

            s.CloseDocByPath(path);
        }

        private static void RunNegLate(TestRun run, WdSession s)
        {
            string xml = WriteXml(GridXml(LabelA, "1", LabelB, "2"));
            string xlsx = Path.Combine(Path.GetTempPath(), "ew-wt-" + Guid.NewGuid().ToString("N") + ".xlsx");
            ToolResult xl = s.Open(xlsx, true, "excel");
            if (!xl.Success)
            {
                run.CaseFail("wt-neg-wrong-kind-excel", xl.Error ?? "建 xlsx 失败");
            }
            else
            {
                ExpectFail(run, "wt-neg-wrong-kind-excel",
                    s.CreateTableFromXml(CreateArgs(WdSession.Str(xl, "channel_id"), xml)), "excel");
                WdSession.CloseExcelByPath(WdSession.Str(xl, "path"));
            }

            string pptx = Path.Combine(Path.GetTempPath(), "ew-wt-" + Guid.NewGuid().ToString("N") + ".pptx");
            ToolResult ppt = s.Open(pptx, true, "powerpoint");
            if (!ppt.Success)
            {
                run.CaseFail("wt-neg-wrong-kind-ppt", ppt.Error ?? "建 pptx 失败");
            }
            else
            {
                ExpectFail(run, "wt-neg-wrong-kind-ppt",
                    s.CreateTableFromXml(CreateArgs(WdSession.Str(ppt, "channel_id"), xml)), "ppt");
                WdSession.ClosePptByPath(WdSession.Str(ppt, "path"));
            }

            string closePath = s.NewDocxPath();
            ToolResult opened = s.Open(closePath, true, s.App);
            if (!opened.Success)
            {
                run.CaseFail("wt-neg-closed", opened.Error ?? "再建稿失败");
                return;
            }

            string cid = WdSession.Str(opened, "channel_id");
            if (!s.CloseDocByPath(closePath))
            {
                run.CaseFail("wt-neg-closed", "COM 关稿失败");
                return;
            }

            ToolResult closed = s.CreateTableFromXml(CreateArgs(cid, xml));
            if (closed != null && !closed.Success
                && (Contains(closed.Error, "关闭")
                    || Contains(closed.Error, "找不到")
                    || Contains(closed.Error, "未知")
                    || Contains(closed.Error, "unsupported")))
            {
                run.CaseOk("wt-neg-closed");
            }
            else if (closed != null && !closed.Success)
            {
                run.CaseFail("wt-neg-closed", closed.Error ?? "");
            }
            else
            {
                run.CaseFail("wt-neg-closed", "关稿后仍成功");
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
            Ok(run, "wt-no-quit", wordOk && wpsOk, wordDetail + " " + wpsDetail);
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

        private static Dictionary<string, object> CreateArgsFilename(string channelId, string path)
        {
            var args = new Dictionary<string, object> { ["filename"] = path ?? "" };
            if (channelId != null)
            {
                args["channel_id"] = channelId;
            }

            return args;
        }

        private static bool CreateOkHas(WdSession s, string id, string xml, string needle)
        {
            ToolResult created = s.CreateTableFromXml(CreateArgs(id, WriteXml(xml)));
            return created.Success && StillHas(s, id, needle);
        }

        private static string[] PadValues(int emptyCount, string first)
        {
            var values = new string[Math.Max(0, emptyCount)];
            if (values.Length > 0)
            {
                values[0] = first ?? "";
            }

            for (int i = 1; i < values.Length; i++)
            {
                values[i] = "";
            }

            return values;
        }

        private static string OneCellXml(string text)
        {
            return
                "<TableConfig><General><TableFontColor>000000</TableFontColor></General><Table><Properties><ColWidths><float>120</float></ColWidths></Properties><Data>"
                + "<tr><td>" + text + "</td></tr>"
                + "</Data></Table></TableConfig>";
        }

        private static string ThreeRowXml()
        {
            return
                "<TableConfig><General><TableFontColor>000000</TableFontColor></General><Table><Properties><ColWidths><float>120</float></ColWidths></Properties><Data>"
                + "<tr><td>" + Row3A + "</td></tr>"
                + "<tr><td>" + Row3B + "</td></tr>"
                + "<tr><td>" + Row3C + "</td></tr>"
                + "</Data></Table></TableConfig>";
        }

        private static string SpanXml()
        {
            return
                "<TableConfig><General><TableFontColor>000000</TableFontColor></General><Table><Properties><ColWidths><float>180</float></ColWidths></Properties><Data>"
                + "<tr><td><span style='font-weight:bold'>" + SpanTxt + "</span></td></tr>"
                + "</Data></Table></TableConfig>";
        }

        private static string RowspanXml()
        {
            return
                "<TableConfig><General><TableFontColor>000000</TableFontColor></General><Table><Properties><ColWidths><float>120</float><float>180</float></ColWidths></Properties><Data>"
                + "<tr><td rowspan=\"2\">" + RsLeft + "</td><td>" + RsUp + "</td></tr>"
                + "<tr><td>" + RsDown + "</td></tr>"
                + "</Data></Table></TableConfig>";
        }

        private static bool RowHasSlots(JArray rows)
        {
            if (rows == null)
            {
                return false;
            }

            foreach (JToken token in rows)
            {
                JObject o = token as JObject;
                if (o != null && o["slots"] is JArray)
                {
                    return true;
                }
            }

            return false;
        }

        private static bool TryLastEmptyRow(ToolResult read, out int row, out int emptyCount)
        {
            row = 0;
            emptyCount = 0;
            JObject data = DataObj(read);
            JArray rows = data == null ? null : data["rows"] as JArray;
            if (rows == null)
            {
                return false;
            }

            bool found = false;
            foreach (JToken token in rows)
            {
                JObject o = token as JObject;
                if (o == null)
                {
                    continue;
                }

                int n = 0;
                try
                {
                    n = o["empty_count"] == null ? 0 : o["empty_count"].Value<int>();
                }
                catch
                {
                    n = 0;
                }

                if (n > 0)
                {
                    try
                    {
                        row = o["row"].Value<int>();
                    }
                    catch
                    {
                        return false;
                    }

                    emptyCount = n;
                    found = true;
                }
            }

            return found;
        }

        private static bool TryRowEmptyCount(ToolResult read, int row, out int emptyCount)
        {
            emptyCount = -1;
            JObject data = DataObj(read);
            JArray rows = data == null ? null : data["rows"] as JArray;
            if (rows == null)
            {
                return false;
            }

            foreach (JToken token in rows)
            {
                JObject o = token as JObject;
                if (o == null)
                {
                    continue;
                }

                int r;
                try
                {
                    r = o["row"].Value<int>();
                }
                catch
                {
                    continue;
                }

                if (r != row)
                {
                    continue;
                }

                try
                {
                    emptyCount = o["empty_count"] == null ? 0 : o["empty_count"].Value<int>();
                    return true;
                }
                catch
                {
                    return false;
                }
            }

            return false;
        }

        private static Dictionary<string, object> FillArgs(
            string channelId,
            string tableId,
            int row,
            IList<string> values)
        {
            var args = new Dictionary<string, object>
            {
                ["table_id"] = tableId ?? "",
                ["format"] = DefaultFormat(),
                ["writes"] = new object[]
                {
                    new Dictionary<string, object>
                    {
                        ["row"] = row,
                        ["values"] = values
                    }
                }
            };
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

        private static string GridXml(string a, string b, string c, string d)
        {
            return
                "<TableConfig><General><TableFontColor>000000</TableFontColor></General><Table><Properties><ColWidths><float>120</float><float>180</float></ColWidths></Properties><Data>"
                + "<tr><td>" + a + "</td><td>" + b + "</td></tr>"
                + "<tr><td>" + c + "</td><td>" + d + "</td></tr>"
                + "</Data></Table></TableConfig>";
        }

        private static string MergeXml()
        {
            return
                "<TableConfig><General><TableFontColor>000000</TableFontColor></General><Table><Properties><ColWidths><float>120</float><float>180</float></ColWidths></Properties><Data>"
                + "<tr><td colspan=\"2\">" + MergeHead + "</td></tr>"
                + "<tr><td>" + MergeLeft + "</td><td>" + MergeRight + "</td></tr>"
                + "</Data></Table></TableConfig>";
        }

        private static string WriteXml(string xml)
        {
            string path = Path.Combine(Path.GetTempPath(), "ew-wt-" + Guid.NewGuid().ToString("N") + ".xml");
            File.WriteAllText(path, xml ?? "");
            return path;
        }

        private static bool TryFirstEmptyRow(ToolResult read, out int row, out int emptyCount)
        {
            row = 0;
            emptyCount = 0;
            JObject data = DataObj(read);
            JArray rows = data == null ? null : data["rows"] as JArray;
            if (rows == null)
            {
                return false;
            }

            foreach (JToken token in rows)
            {
                JObject o = token as JObject;
                if (o == null)
                {
                    continue;
                }

                int n = 0;
                try
                {
                    n = o["empty_count"] == null ? 0 : o["empty_count"].Value<int>();
                }
                catch
                {
                    n = 0;
                }

                if (n > 0)
                {
                    try
                    {
                        row = o["row"].Value<int>();
                    }
                    catch
                    {
                        return false;
                    }

                    emptyCount = n;
                    return true;
                }
            }

            return false;
        }

        private static string FirstCode(string text, string prefix)
        {
            if (string.IsNullOrEmpty(text))
            {
                return "";
            }

            Match m = Regex.Match(text, @"\[" + prefix + @"_([0-9a-zA-Z]{5})-start\]");
            return m.Success ? prefix + "_" + m.Groups[1].Value : "";
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

        private static string ReadBody(WdSession s, string id)
        {
            ToolResult got = s.GetContent(id);
            return got.Success ? WdSession.Str(got, "document_content") : "";
        }

        private static bool StillHas(WdSession s, string id, string needle)
        {
            return ReadBody(s, id).IndexOf(needle, StringComparison.Ordinal) >= 0;
        }

        private static JObject DataObj(ToolResult r)
        {
            if (r == null || r.Data == null)
            {
                return null;
            }

            try
            {
                return JObject.Parse(JsonConvert.SerializeObject(r.Data));
            }
            catch
            {
                return null;
            }
        }

        private static void ExpectFail(TestRun run, string id, ToolResult r, string needle)
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
