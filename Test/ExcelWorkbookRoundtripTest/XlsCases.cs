using System;
using System.Collections.Generic;
using WordAddIn1;
using WordAddIn1.SpreadsheetHost;

namespace ExcelWorkbookRoundtripTest
{
    internal static class XlsCases
    {
        public static void Run(TestRun run, bool et)
        {
            var s = new XlSession(et);
            RunNeg(run, s);
            RunMerge(run, s);
            RunInsertDelete(run, s);
        }

        private static void RunNeg(TestRun run, XlSession s)
        {
            ChannelRegistry.ClearAll();
            ToolResult noCh = s.Structure(null, "利润表", "merge", "A1:B1");
            ExpectFail(run, "xls-neg-no-channel", noCh, "打开");

            var req = new SpreadsheetStructureRequest
            {
                Action = "merge",
                SheetName = "利润表",
                RangeA1 = "A1:B1",
                Count = 0
            };
            if (!SpreadsheetHostAdapter.TryApplyStructure(new FakeWordChannel(), req, out _, out ToolResult wrong))
            {
                ExpectFail(run, "xls-neg-wrong-kind", wrong, "格子结构仅支持");
            }
            else
            {
                run.CaseFail("xls-neg-wrong-kind", "word 渠道应失败");
            }

            WithBook(run, s, "xls-neg-no-sheet", (id, ch) =>
            {
                ToolResult r = s.Call("F_apply_excel_structure", new Dictionary<string, object>
                {
                    ["channel_id"] = id,
                    ["action"] = "merge",
                    ["range"] = "A1:B1"
                });
                return r.Success ? "缺 sheet 应失败" : (Has(r.Error, "sheet") ? null : r.Error);
            });

            WithBook(run, s, "xls-neg-no-action", (id, ch) =>
            {
                ToolResult r = s.Call("F_apply_excel_structure", new Dictionary<string, object>
                {
                    ["channel_id"] = id,
                    ["sheet"] = s.FirstSheetName(ch),
                    ["range"] = "A1:B1"
                });
                return r.Success ? "缺 action 应失败" : (Has(r.Error, "action") ? null : r.Error);
            });

            WithBook(run, s, "xls-neg-no-range", (id, ch) =>
            {
                ToolResult r = s.Call("F_apply_excel_structure", new Dictionary<string, object>
                {
                    ["channel_id"] = id,
                    ["sheet"] = s.FirstSheetName(ch),
                    ["action"] = "merge"
                });
                return r.Success ? "缺 range 应失败" : (Has(r.Error, "range") ? null : r.Error);
            });

            WithBook(run, s, "xls-neg-bad-action", (id, ch) =>
            {
                ToolResult r = s.Structure(id, s.FirstSheetName(ch), "resize", "A1:B1");
                return r.Success ? "非法 action 应失败" : (Has(r.Error, "action") ? null : r.Error);
            });

            WithBook(run, s, "xls-neg-bang", (id, ch) =>
            {
                string name = s.FirstSheetName(ch);
                ToolResult r = s.Structure(id, name, "merge", name + "!A1:C1");
                return r.Success ? "含!应失败" : (Has(r.Error, "!") || Has(r.Error, "纯 A1") ? null : r.Error);
            });

            WithBook(run, s, "xls-neg-missing-sheet", (id, ch) =>
            {
                ToolResult r = s.Structure(id, "不存在的表", "merge", "A1:B1");
                return r.Success
                    ? "缺表应失败"
                    : (Has(r.Error, "不存在") || Has(r.Error, "未找到") ? null : r.Error);
            });

            WithBook(run, s, "xls-neg-chart", (id, ch) =>
            {
                if (!s.SeedChartSheet(ch, out string chartName))
                {
                    return "铺图表 sheet 失败";
                }

                ToolResult r = s.Structure(id, chartName, "merge", "A1:B1");
                return r.Success ? "图表 sheet 应失败" : (Has(r.Error, "sheet_type=chart") ? null : r.Error);
            });

            WithBook(run, s, "xls-neg-bad-a1", (id, ch) =>
            {
                ToolResult r = s.Structure(id, s.FirstSheetName(ch), "merge", "ZZ");
                return r.Success ? "非法 range 应失败" : (Has(r.Error, "非法") || Has(r.Error, "range") ? null : r.Error);
            });
        }

        private static void RunMerge(TestRun run, XlSession s)
        {
            WithBook(run, s, "xls-merge-row", (id, ch) =>
            {
                string name = s.FirstSheetName(ch);
                if (!s.SeedCell(ch, name, "A1", "标题")
                    || !s.SeedCell(ch, name, "B1", "x")
                    || !s.SeedCell(ch, name, "C1", "y"))
                {
                    return "铺底失败";
                }

                ToolResult r = s.Structure(id, name, "merge", "A1:C1");
                if (!r.Success)
                {
                    return r.Error;
                }

                if (!Has(Str(r, "affected"), "A1:C1"))
                {
                    return "affected=" + Str(r, "affected");
                }

                ToolResult rd = s.Read(id, name, "A1:C1");
                if (!rd.Success)
                {
                    return rd.Error;
                }

                string table = Str(rd, "table");
                if (!Has(table, "标题"))
                {
                    return "应留标题 " + table;
                }

                if (HasAddr(table, "B1") || HasAddr(table, "C1"))
                {
                    return "续格不该独立 " + table;
                }

                return HasAddr(table, "A1") && (Has(table, "area=") || Has(table, "colspan"))
                    ? null
                    : "缺 A1/area " + table;
            });

            WithBook(run, s, "xls-merge-block", (id, ch) =>
            {
                string name = s.FirstSheetName(ch);
                if (!s.SeedCell(ch, name, "A1", "左上")
                    || !s.SeedCell(ch, name, "B1", "右上")
                    || !s.SeedCell(ch, name, "A2", "左下")
                    || !s.SeedCell(ch, name, "B2", "右下"))
                {
                    return "铺底失败";
                }

                ToolResult r = s.Structure(id, name, "merge", "A1:B2");
                if (!r.Success)
                {
                    return r.Error;
                }

                ToolResult rd = s.Read(id, name, "A1:B2");
                if (!rd.Success)
                {
                    return rd.Error;
                }

                string table = Str(rd, "table");
                return Has(table, "左上") && (Has(table, "area=\"A1:B2\"") || Has(table, "colspan") || Has(table, "rowspan"))
                    ? null
                    : "缺块合并 " + table;
            });

            WithBook(run, s, "xls-neg-merge-single", (id, ch) =>
            {
                ToolResult r = s.Structure(id, s.FirstSheetName(ch), "merge", "A1");
                return r.Success ? "单格 merge 应失败" : (Has(r.Error, "2") || Has(r.Error, "至少") ? null : r.Error);
            });

            WithBook(run, s, "xls-neg-merge-oversize", (id, ch) =>
            {
                ToolResult r = s.Structure(id, s.FirstSheetName(ch), "merge", "A1:A201");
                if (r.Success)
                {
                    return "超限 merge 应失败";
                }

                return Has(r.Error, "过大") || Has(r.Error, "200") ? null : r.Error;
            });

            WithBook(run, s, "xls-unmerge-origin", (id, ch) =>
            {
                string name = s.FirstSheetName(ch);
                if (!s.SeedMerge(ch, name))
                {
                    return "铺合并失败";
                }

                ToolResult r = s.Structure(id, name, "unmerge", "A1");
                if (!r.Success)
                {
                    return r.Error;
                }

                ToolResult rd = s.Read(id, name, "A1:C1");
                if (!rd.Success)
                {
                    return rd.Error;
                }

                string table = Str(rd, "table");
                return !Has(table, "area=\"A1:C1\"") && !Has(table, "colspan")
                    ? null
                    : "拆开后仍是整块 " + table;
            });

            WithBook(run, s, "xls-unmerge-body", (id, ch) =>
            {
                string name = s.FirstSheetName(ch);
                if (!s.SeedMerge(ch, name))
                {
                    return "铺合并失败";
                }

                ToolResult r = s.Structure(id, name, "unmerge", "B1");
                if (!r.Success)
                {
                    return r.Error;
                }

                ToolResult rd = s.Read(id, name, "A1:C1");
                if (!rd.Success)
                {
                    return rd.Error;
                }

                string table = Str(rd, "table");
                return !Has(table, "area=\"A1:C1\"") && !Has(table, "colspan")
                    ? null
                    : "非左上角拆开失败 " + table;
            });

            WithBook(run, s, "xls-merge-keep-value", (id, ch) =>
            {
                string name = s.FirstSheetName(ch);
                if (!s.SeedCell(ch, name, "A1", "甲") || !s.SeedCell(ch, name, "B1", "乙"))
                {
                    return "铺底失败";
                }

                ToolResult r = s.Structure(id, name, "merge", "A1:B1");
                if (!r.Success)
                {
                    return r.Error;
                }

                ToolResult rd = s.Read(id, name, "A1:B1");
                if (!rd.Success)
                {
                    return rd.Error;
                }

                string table = Str(rd, "table");
                if (!Has(table, "甲"))
                {
                    return "左上角值丢了 " + table;
                }

                return !HasAddr(table, "B1")
                    ? null
                    : "B1 仍独立 " + table;
            });

            WithBook(run, s, "xls-merge-hidden", (id, ch) =>
            {
                if (!s.SeedHidden(ch, "可见表", "隐藏底稿") || !s.SeedCell(ch, "隐藏底稿", "B1", "续"))
                {
                    return "铺隐藏表失败";
                }

                ToolResult r = s.Structure(id, "隐藏底稿", "merge", "A1:B1");
                if (!r.Success)
                {
                    return r.Error;
                }

                ToolResult rd = s.Read(id, "隐藏底稿", "A1:B1");
                if (!rd.Success)
                {
                    return rd.Error;
                }

                string table = Str(rd, "table");
                return Has(table, "隐") && (Has(table, "area=") || Has(table, "colspan"))
                    ? null
                    : "隐藏表合并未见 " + table;
            });

            WithBook(run, s, "xls-zh-sheet", (id, ch) =>
            {
                if (!s.SeedUsed(ch, "利润表") || !s.SeedCell(ch, "利润表", "C1", "旁"))
                {
                    return "铺底失败";
                }

                ToolResult r = s.Structure(id, "利润表", "merge", "A1:C1");
                if (!r.Success)
                {
                    return r.Error;
                }

                if (!string.Equals(Str(r, "sheet"), "利润表", StringComparison.Ordinal))
                {
                    return "sheet=" + Str(r, "sheet");
                }

                ToolResult rd = s.Read(id, "利润表", "A1:C1");
                if (!rd.Success)
                {
                    return rd.Error;
                }

                string table = Str(rd, "table");
                return Has(table, "地区") && (Has(table, "area=") || Has(table, "colspan"))
                    ? null
                    : "中文表合并未见 " + table;
            });

            WithBook(run, s, "xls-unmerge-rect", (id, ch) =>
            {
                string name = s.FirstSheetName(ch);
                if (!s.SeedMerge(ch, name))
                {
                    return "铺合并失败";
                }

                ToolResult r = s.Structure(id, name, "unmerge", "A1:C1");
                if (!r.Success)
                {
                    return r.Error;
                }

                ToolResult rd = s.Read(id, name, "A1:C1");
                if (!rd.Success)
                {
                    return rd.Error;
                }

                string table = Str(rd, "table");
                return !Has(table, "area=\"A1:C1\"") && !Has(table, "colspan")
                    ? null
                    : "整块拆开失败 " + table;
            });
        }

        private static void RunInsertDelete(TestRun run, XlSession s)
        {
            WithBook(run, s, "xls-ins-row", (id, ch) =>
            {
                string name = s.FirstSheetName(ch);
                if (!s.SeedCell(ch, name, "A3", "下"))
                {
                    return "铺底失败";
                }

                ToolResult r = s.Structure(id, name, "insert_rows", "A3");
                if (!r.Success)
                {
                    return r.Error;
                }

                if (Int(r, "count") != 1)
                {
                    return "count=" + Int(r, "count");
                }

                ToolResult rd = s.Read(id, name, "A3:A4");
                if (!rd.Success)
                {
                    return rd.Error;
                }

                string table = Str(rd, "table");
                return HasCellText(table, "A4", "下") && !HasCellText(table, "A3", "下")
                    ? null
                    : "下未下移 " + table;
            });

            WithBook(run, s, "xls-ins-rows-count", (id, ch) =>
            {
                string name = s.FirstSheetName(ch);
                if (!s.SeedCell(ch, name, "A5", "锚"))
                {
                    return "铺底失败";
                }

                ToolResult r = s.Structure(id, name, "insert_rows", "A5", 3);
                if (!r.Success)
                {
                    return r.Error;
                }

                if (!Has(Str(r, "affected"), "rows=5:7"))
                {
                    return "affected=" + Str(r, "affected");
                }

                ToolResult rd = s.Read(id, name, "A5:A8");
                if (!rd.Success)
                {
                    return rd.Error;
                }

                string table = Str(rd, "table");
                return HasCellText(table, "A8", "锚") && !HasCellText(table, "A5", "锚")
                    ? null
                    : "锚未下移 3 行 " + table;
            });

            WithBook(run, s, "xls-del-row", (id, ch) =>
            {
                string name = s.FirstSheetName(ch);
                if (!s.SeedCell(ch, name, "A2", "删") || !s.SeedCell(ch, name, "A3", "留"))
                {
                    return "铺底失败";
                }

                ToolResult r = s.Structure(id, name, "delete_rows", "A2", 1);
                if (!r.Success)
                {
                    return r.Error;
                }

                ToolResult rd = s.Read(id, name, "A2:A3");
                if (!rd.Success)
                {
                    return rd.Error;
                }

                string table = Str(rd, "table");
                return HasCellText(table, "A2", "留") && !HasCellText(table, "A2", "删")
                    ? null
                    : "删行后地址不对 " + table;
            });

            WithBook(run, s, "xls-ins-col", (id, ch) =>
            {
                string name = s.FirstSheetName(ch);
                if (!s.SeedCell(ch, name, "C1", "右"))
                {
                    return "铺底失败";
                }

                ToolResult r = s.Structure(id, name, "insert_cols", "C1");
                if (!r.Success)
                {
                    return r.Error;
                }

                ToolResult rd = s.Read(id, name, "C1:D1");
                if (!rd.Success)
                {
                    return rd.Error;
                }

                string table = Str(rd, "table");
                return HasCellText(table, "D1", "右") && !HasCellText(table, "C1", "右")
                    ? null
                    : "右未右移 " + table;
            });

            WithBook(run, s, "xls-del-col", (id, ch) =>
            {
                string name = s.FirstSheetName(ch);
                if (!s.SeedCell(ch, name, "B1", "删列") || !s.SeedCell(ch, name, "C1", "右邻"))
                {
                    return "铺底失败";
                }

                ToolResult r = s.Structure(id, name, "delete_cols", "B1");
                if (!r.Success)
                {
                    return r.Error;
                }

                ToolResult rd = s.Read(id, name, "A1:C1");
                if (!rd.Success)
                {
                    return rd.Error;
                }

                string table = Str(rd, "table");
                return HasCellText(table, "B1", "右邻") && !HasCellText(table, "B1", "删列")
                    ? null
                    : "删列后地址不对 " + table;
            });

            WithBook(run, s, "xls-neg-ins-rect", (id, ch) =>
            {
                ToolResult r = s.Structure(id, s.FirstSheetName(ch), "insert_rows", "A5:C10");
                return r.Success ? "插删矩形应失败" : (Has(r.Error, "单格") ? null : r.Error);
            });

            WithBook(run, s, "xls-neg-del-rect", (id, ch) =>
            {
                ToolResult r = s.Structure(id, s.FirstSheetName(ch), "delete_cols", "B1:D1");
                return r.Success ? "删列矩形应失败" : (Has(r.Error, "单格") ? null : r.Error);
            });

            WithBook(run, s, "xls-neg-count-over", (id, ch) =>
            {
                string name = s.FirstSheetName(ch);
                if (!s.SeedCell(ch, name, "A2", "旧"))
                {
                    return "铺底失败";
                }

                ToolResult r = s.Structure(id, name, "insert_rows", "A2", 51);
                if (r.Success)
                {
                    return "count=51 应失败";
                }

                if (!Has(r.Error, "50") && !Has(r.Error, "过大"))
                {
                    return r.Error;
                }

                ToolResult rd = s.Read(id, name, "A2");
                if (!rd.Success)
                {
                    return rd.Error;
                }

                return HasCellText(Str(rd, "table"), "A2", "旧")
                    ? null
                    : "超限后原格变了 " + Str(rd, "table");
            });

            WithBook(run, s, "xls-reread-addr", (id, ch) =>
            {
                string name = s.FirstSheetName(ch);
                if (!s.SeedCell(ch, name, "A2", "旧址"))
                {
                    return "铺底失败";
                }

                ToolResult r = s.Structure(id, name, "insert_rows", "A2");
                if (!r.Success)
                {
                    return r.Error;
                }

                ToolResult oldAddr = s.Read(id, name, "A2");
                if (!oldAddr.Success)
                {
                    return oldAddr.Error;
                }

                if (HasCellText(Str(oldAddr, "table"), "A2", "旧址"))
                {
                    return "旧 A2 仍是旧址，D22 未钉住";
                }

                ToolResult moved = s.Read(id, name, "A3");
                if (!moved.Success)
                {
                    return moved.Error;
                }

                return HasCellText(Str(moved, "table"), "A3", "旧址")
                    ? null
                    : "下移后没抄到 " + Str(moved, "table");
            });

            WithBook(run, s, "xls-default-count", (id, ch) =>
            {
                string name = s.FirstSheetName(ch);
                if (!s.SeedCell(ch, name, "B1", "锚"))
                {
                    return "铺底失败";
                }

                ToolResult r = s.Structure(id, name, "insert_cols", "B1");
                if (!r.Success)
                {
                    return r.Error;
                }

                if (Int(r, "count") != 1)
                {
                    return "默认 count=" + Int(r, "count");
                }

                if (!Has(Str(r, "affected"), "cols=B:B"))
                {
                    return "affected=" + Str(r, "affected");
                }

                ToolResult rd = s.Read(id, name, "B1:C1");
                if (!rd.Success)
                {
                    return rd.Error;
                }

                string table = Str(rd, "table");
                return HasCellText(table, "C1", "锚") && !HasCellText(table, "B1", "锚")
                    ? null
                    : "默认只插一列未对上 " + table;
            });
        }

        private static void WithBook(
            TestRun run,
            XlSession s,
            string id,
            Func<string, IOperationChannel, string> body)
        {
            string path = s.NewXlsxPath();
            ToolResult opened = s.Open(path, true);
            if (!opened.Success)
            {
                run.CaseFail(id, "打开失败: " + opened.Error);
                return;
            }

            string channelId = Str(opened, "channel_id");
            if (!s.TryGetChannel(channelId, out IOperationChannel ch))
            {
                run.CaseFail(id, "渠道不在注册表");
                s.CloseBook(channelId);
                return;
            }

            string err;
            try
            {
                err = body(channelId, ch);
            }
            catch (Exception ex)
            {
                err = ex.Message;
            }

            s.CloseBook(channelId);
            if (err != null && err.StartsWith("SKIP:", StringComparison.Ordinal))
            {
                run.CaseSkip(id, err.Substring(5).Trim());
                return;
            }

            if (err == null)
            {
                run.CaseOk(id);
            }
            else
            {
                run.CaseFail(id, err);
            }
        }

        private static void ExpectFail(TestRun run, string id, ToolResult r, string mustContain)
        {
            if (r != null && !r.Success && (mustContain == null || Has(r.Error, mustContain)))
            {
                run.CaseOk(id);
                return;
            }

            run.CaseFail(id, r == null ? "无回包" : (r.Success ? "应失败" : r.Error));
        }

        private static bool Has(string text, string part)
        {
            return !string.IsNullOrEmpty(text)
                && text.IndexOf(part, StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private static bool HasAddr(string table, string a1)
        {
            return Has(table, "addr=\"" + a1 + "\"");
        }

        private static bool HasCellText(string table, string a1, string text)
        {
            if (string.IsNullOrEmpty(table) || string.IsNullOrEmpty(a1))
            {
                return false;
            }

            int i = table.IndexOf("addr=\"" + a1 + "\"", StringComparison.Ordinal);
            if (i < 0)
            {
                return false;
            }

            int gt = table.IndexOf('>', i);
            int close = table.IndexOf("</", gt + 1);
            if (gt < 0 || close < 0)
            {
                return false;
            }

            return table.Substring(gt + 1, close - gt - 1).IndexOf(text, StringComparison.Ordinal) >= 0;
        }

        private static Dictionary<string, object> Data(ToolResult r)
        {
            return r?.Data as Dictionary<string, object>;
        }

        private static string Str(ToolResult r, string key)
        {
            Dictionary<string, object> data = Data(r);
            if (data == null || !data.ContainsKey(key) || data[key] == null)
            {
                return "";
            }

            return Convert.ToString(data[key]) ?? "";
        }

        private static int Int(ToolResult r, string key)
        {
            Dictionary<string, object> data = Data(r);
            if (data == null || !data.ContainsKey(key) || data[key] == null)
            {
                return 0;
            }

            try
            {
                return Convert.ToInt32(data[key]);
            }
            catch
            {
                return 0;
            }
        }
    }
}
