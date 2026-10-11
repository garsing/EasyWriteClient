using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using WordAddIn1;

namespace WordDocumentRoundtripTest
{
    /// <summary>
    /// 第 6 波：XML 建表、填空槽、抽/套格式、删表重建、新建后第一次 optimize。
    /// </summary>
    internal static class WtCases
    {
        public const int Target = 44;

        private const string LabelA = "易写表指标";
        private const string LabelB = "易写表WACC";
        private const string FillA = "8.46%";
        private const string FillB = "12.00%";
        private const string MergeHead = "易写表合并头";
        private const string MergeLeft = "易写表左";
        private const string MergeRight = "易写表右";
        private const string AfterSent = "易写表前句。";
        private const string RebuildA = "易写表重建甲";
        private const string RebuildB = "易写表重建乙";

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
            string sent = FirstCode(WdSession.Str(got, "document_content"), "S");
            if (string.IsNullOrEmpty(sent))
            {
                FailMany(run, "无可抄 S_",
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
