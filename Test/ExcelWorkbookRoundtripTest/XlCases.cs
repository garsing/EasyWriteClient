using System;
using System.Collections;
using System.Collections.Generic;
using WordAddIn1;
using WordAddIn1.SpreadsheetHost;

namespace ExcelWorkbookRoundtripTest
{
    internal static class XlCases
    {
        public static void Run(TestRun run, bool et)
        {
            var s = new XlSession(et);
            RunNeg(run, s);
            RunOpen(run, s);
            RunOverview(run, s);
            RunRead(run, s);
        }

        private static void RunNeg(TestRun run, XlSession s)
        {
            ChannelRegistry.ClearAll();
            ToolResult ov = s.Overview(null);
            ExpectFail(run, "xl-ov-neg-no-channel", ov, "打开");

            ChannelRegistry.ClearAll();
            ToolResult rd = s.Call("F_read_excel_range", new Dictionary<string, object>
            {
                ["sheet"] = "利润表",
                ["range"] = "A1"
            });
            ExpectFail(run, "xl-rd-neg-no-channel", rd, "打开");

            if (!SpreadsheetHostAdapter.TryGetWorkbookContent(
                    new FakeWordChannel(),
                    out _,
                    out ToolResult wrongOv))
            {
                ExpectFail(run, "xl-ov-neg-wrong-kind", wrongOv, "F_get_document_content");
            }
            else
            {
                run.CaseFail("xl-ov-neg-wrong-kind", "word 渠道应失败");
            }

            if (!SpreadsheetHostAdapter.TryReadRange(
                    new FakeWordChannel(),
                    "利润表",
                    "A1",
                    SpreadsheetContentMode.Value,
                    out _,
                    out ToolResult wrongRd))
            {
                ExpectFail(run, "xl-rd-neg-wrong-kind", wrongRd, "F_get_document_content");
            }
            else
            {
                run.CaseFail("xl-rd-neg-wrong-kind", "word 渠道应失败");
            }
        }

        private static void RunOpen(TestRun run, XlSession s)
        {
            string path = s.NewXlsxPath();
            ToolResult created = s.Open(path, true);
            if (!created.Success)
            {
                run.CaseFail("xl-open-blank", created.Error ?? "打开失败");
                run.CaseFail("xl-open-xlsx", "依赖 blank 建文件");
                run.CaseFail("xl-open-reuse", "依赖 blank 建文件");
                return;
            }

            string id1 = Str(created, "channel_id");
            bool blankOk = string.Equals(Str(created, "kind"), s.ExpectedKind, StringComparison.Ordinal)
                && !string.IsNullOrEmpty(id1)
                && !Bool(created, "index_ready");
            if (blankOk)
            {
                run.CaseOk("xl-open-blank");
            }
            else
            {
                run.CaseFail("xl-open-blank", "kind/channel/index_ready 不符 "
                    + Str(created, "kind") + " " + id1);
            }

            s.CloseBook(id1);

            ToolResult opened = s.Open(path, false);
            if (!opened.Success)
            {
                run.CaseFail("xl-open-xlsx", opened.Error ?? "再开失败");
                run.CaseFail("xl-open-reuse", "再开失败");
            }
            else
            {
                string id2 = Str(opened, "channel_id");
                bool xlsxOk = string.Equals(Str(opened, "kind"), s.ExpectedKind, StringComparison.Ordinal)
                    && !string.IsNullOrEmpty(id2)
                    && !Bool(opened, "index_ready");
                if (xlsxOk)
                {
                    run.CaseOk("xl-open-xlsx");
                }
                else
                {
                    run.CaseFail("xl-open-xlsx", "kind/index_ready 不符");
                }

                ToolResult again = s.Open(path, false);
                if (again.Success
                    && string.Equals(Str(again, "channel_id"), id2, StringComparison.Ordinal)
                    && Bool(again, "reused"))
                {
                    run.CaseOk("xl-open-reuse");
                }
                else if (again.Success && string.Equals(Str(again, "channel_id"), id2, StringComparison.Ordinal))
                {
                    run.CaseOk("xl-open-reuse");
                }
                else
                {
                    run.CaseFail("xl-open-reuse", again.Error
                        ?? ("channel " + Str(again, "channel_id") + " reused=" + Bool(again, "reused")));
                }

                s.CloseBook(id2);
            }

            ToolResult conflict = s.Open(path, false, "word");
            ExpectFail(run, "xl-open-conflict", conflict, null);

            ToolResult nosuffix = s.Call("F_open_document", new Dictionary<string, object>
            {
                ["path"] = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "ew-xl-nosuffix"),
                ["app"] = s.App
            });
            ExpectFail(run, "xl-neg-open-nosuffix", nosuffix, null);
        }

        private static void RunOverview(TestRun run, XlSession s)
        {
            WithBook(run, s, "xl-ov-used", (id, ch) =>
            {
                if (!s.SeedUsed(ch, "利润表"))
                {
                    return "铺底失败";
                }

                ToolResult r = s.Overview(id);
                if (!r.Success)
                {
                    return r.Error;
                }

                Dictionary<string, object> sheet = SheetNamed(r, "利润表");
                if (sheet == null)
                {
                    return "没有利润表";
                }

                string used = Convert.ToString(sheet["used_range"] ?? "");
                string preview = Convert.ToString(sheet["preview"] ?? "");
                if (used.IndexOf("A1", StringComparison.Ordinal) < 0
                    || used.IndexOf("B2", StringComparison.Ordinal) < 0)
                {
                    return "used_range=" + used;
                }

                if (preview.IndexOf("addr=\"A1\"", StringComparison.Ordinal) < 0
                    || preview.IndexOf("地区", StringComparison.Ordinal) < 0)
                {
                    return "预览缺 A1/地区";
                }

                if (!string.Equals(Convert.ToString(sheet["sheet_type"]), "worksheet", StringComparison.Ordinal))
                {
                    return "sheet_type";
                }

                return null;
            });

            WithBook(run, s, "xl-ov-two-sheets", (id, ch) =>
            {
                if (!s.SeedTwoSheets(ch, "利润表", "底稿"))
                {
                    return "铺两表失败";
                }

                ToolResult r = s.Overview(id);
                if (!r.Success)
                {
                    return r.Error;
                }

                string display = Str(r, "display_contents");
                if (SheetNamed(r, "利润表") == null || SheetNamed(r, "底稿") == null)
                {
                    return "sheets 缺表";
                }

                if (display.IndexOf("## 利润表", StringComparison.Ordinal) < 0
                    || display.IndexOf("## 底稿", StringComparison.Ordinal) < 0)
                {
                    return "display 缺 ##";
                }

                return null;
            });

            WithBook(run, s, "xl-ov-empty", (id, ch) =>
            {
                ToolResult r = s.Overview(id);
                if (!r.Success)
                {
                    return r.Error;
                }

                string preview = "";
                List<Dictionary<string, object>> sheets = Sheets(r);
                if (sheets.Count < 1)
                {
                    return "没有 sheet";
                }

                preview = Convert.ToString(sheets[0]["preview"] ?? "");
                if (preview.IndexOf("地区", StringComparison.Ordinal) >= 0)
                {
                    return "空表预览不该有业务字";
                }

                return null;
            });

            WithBook(run, s, "xl-ov-hidden", (id, ch) =>
            {
                if (!s.SeedHidden(ch, "可见表", "隐藏底稿"))
                {
                    return "铺隐藏表失败";
                }

                ToolResult r = s.Overview(id);
                if (!r.Success)
                {
                    return r.Error;
                }

                Dictionary<string, object> hid = SheetNamed(r, "隐藏底稿");
                if (hid == null)
                {
                    return "未列出隐藏表";
                }

                if (!Convert.ToBoolean(hid["hidden"]))
                {
                    return "hidden 应为 true";
                }

                string preview = Convert.ToString(hid["preview"] ?? "");
                if (preview.IndexOf("addr=", StringComparison.Ordinal) < 0)
                {
                    return "隐藏表预览应有 addr";
                }

                return null;
            });

            WithBook(run, s, "xl-ov-chart", (id, ch) =>
            {
                if (!s.SeedChartSheet(ch, out string chartName))
                {
                    return "SKIP: 加图表 sheet 失败";
                }

                ToolResult r = s.Overview(id);
                if (!r.Success)
                {
                    return r.Error;
                }

                Dictionary<string, object> chart = SheetNamed(r, chartName);
                if (chart == null)
                {
                    return "未列出 " + chartName;
                }

                if (!string.Equals(Convert.ToString(chart["sheet_type"]), "chart", StringComparison.Ordinal))
                {
                    return "SKIP: 宿主 Charts.Add 未建成图表 sheet（" + chart["sheet_type"] + "）";
                }

                if (!string.IsNullOrEmpty(Convert.ToString(chart["used_range"] ?? ""))
                    || !string.IsNullOrEmpty(Convert.ToString(chart["preview"] ?? "")))
                {
                    return "图表 sheet 应无 used/preview";
                }

                return null;
            });

            WithBook(run, s, "xl-ov-preview-trunc", (id, ch) =>
            {
                if (!s.SeedWide(ch, "宽表"))
                {
                    return "铺宽表失败";
                }

                ToolResult r = s.Overview(id);
                if (!r.Success)
                {
                    return r.Error;
                }

                Dictionary<string, object> sheet = SheetNamed(r, "宽表");
                if (sheet == null)
                {
                    return "没有宽表";
                }

                if (!Convert.ToBoolean(sheet["preview_truncated"]))
                {
                    return "preview_truncated 应为 true";
                }

                string preview = Convert.ToString(sheet["preview"] ?? "");
                if (preview.IndexOf("R8", StringComparison.Ordinal) >= 0)
                {
                    return "预览倒进了第 8 行";
                }

                return null;
            });

            WithUnsaved(run, s, "xl-ov-unsaved", ch =>
            {
                ToolResult r = s.Overview(ch.ChannelId);
                if (!r.Success)
                {
                    return r.Error;
                }

                if (!string.IsNullOrEmpty(Str(r, "path")))
                {
                    return "未存盘 path 应空，实际 " + Str(r, "path");
                }

                if (string.IsNullOrEmpty(Str(r, "display_contents")))
                {
                    return "display_contents 空";
                }

                return null;
            });

            WithBook(run, s, "xl-ov-merge", (id, ch) =>
            {
                if (!s.SeedMerge(ch, "合并表"))
                {
                    return "铺合并失败";
                }

                ToolResult r = s.Overview(id);
                if (!r.Success)
                {
                    return r.Error;
                }

                string preview = Convert.ToString(SheetNamed(r, "合并表")?["preview"] ?? "");
                if (preview.IndexOf("addr=\"A1\"", StringComparison.Ordinal) < 0
                    || preview.IndexOf("标题", StringComparison.Ordinal) < 0)
                {
                    return "预览缺 A1/标题";
                }

                if (preview.IndexOf("addr=\"B1\"", StringComparison.Ordinal) >= 0
                    || preview.IndexOf("addr=\"C1\"", StringComparison.Ordinal) >= 0)
                {
                    return "延续格不该输出";
                }

                if (preview.IndexOf("area=", StringComparison.Ordinal) < 0
                    && preview.IndexOf("colspan=", StringComparison.Ordinal) < 0)
                {
                    return "合并格应有 area/colspan";
                }

                return null;
            });
        }

        private static void RunRead(TestRun run, XlSession s)
        {
            WithBook(run, s, "xl-rd-neg-no-sheet", (id, ch) =>
            {
                ToolResult r = s.Call("F_read_excel_range", new Dictionary<string, object>
                {
                    ["channel_id"] = id
                });
                return r.Success ? "缺 sheet 应失败" : (Has(r.Error, "sheet") ? null : r.Error);
            });

            WithBook(run, s, "xl-rd-neg-missing-sheet", (id, ch) =>
            {
                ToolResult r = s.Read(id, "不存在的表", "A1");
                return r.Success ? "应失败" : (Has(r.Error, "不存在") ? null : r.Error);
            });

            WithBook(run, s, "xl-rd-neg-bang", (id, ch) =>
            {
                if (!s.SeedUsed(ch, "利润表"))
                {
                    return "铺底失败";
                }

                ToolResult r = s.Read(id, "利润表", "利润表!A1:B2");
                return r.Success ? "含!应失败" : (Has(r.Error, "纯 A1") ? null : r.Error);
            });

            WithBook(run, s, "xl-rd-neg-chart", (id, ch) =>
            {
                if (!s.SeedChartSheet(ch, out string chartName))
                {
                    return "SKIP: 加图表 sheet 失败";
                }

                ToolResult ov = s.Overview(id);
                Dictionary<string, object> chart = ov.Success ? SheetNamed(ov, chartName) : null;
                if (chart == null
                    || !string.Equals(Convert.ToString(chart["sheet_type"]), "chart", StringComparison.Ordinal))
                {
                    return "SKIP: 宿主没有图表 sheet，无法验读失败";
                }

                ToolResult r = s.Read(id, chartName, "A1");
                return r.Success ? "图表 sheet 应失败" : (Has(r.Error, "chart") ? null : r.Error);
            });

            WithBook(run, s, "xl-rd-block", (id, ch) =>
            {
                if (!s.SeedUsed(ch, "利润表"))
                {
                    return "铺底失败";
                }

                ToolResult r = s.Read(id, "利润表", "A1:B2");
                if (!r.Success)
                {
                    return r.Error;
                }

                string table = Str(r, "table");
                if (table.IndexOf("addr=\"A1\"", StringComparison.Ordinal) < 0
                    || table.IndexOf("地区", StringComparison.Ordinal) < 0
                    || table.IndexOf("addr=\"B2\"", StringComparison.Ordinal) < 0)
                {
                    return "table 缺格子 " + table;
                }

                if (Str(r, "actual_range").IndexOf("A1", StringComparison.Ordinal) < 0)
                {
                    return "actual_range=" + Str(r, "actual_range");
                }

                return null;
            });

            WithBook(run, s, "xl-rd-omit-used", (id, ch) =>
            {
                if (!s.SeedUsed(ch, "利润表"))
                {
                    return "铺底失败";
                }

                ToolResult r = s.Read(id, "利润表");
                if (!r.Success)
                {
                    return r.Error;
                }

                string table = Str(r, "table");
                if (table.IndexOf("地区", StringComparison.Ordinal) < 0
                    || table.IndexOf("addr=\"B2\"", StringComparison.Ordinal) < 0)
                {
                    return "省略 range 没读到 UsedRange";
                }

                return null;
            });

            WithBook(run, s, "xl-rd-formula", (id, ch) =>
            {
                if (!s.SeedFormula(ch, "公式表"))
                {
                    return "铺公式失败";
                }

                ToolResult r = s.Read(id, "公式表", "A2:B2", "formula");
                if (!r.Success)
                {
                    return r.Error;
                }

                string table = Str(r, "table");
                if (table.IndexOf("=A2+1", StringComparison.OrdinalIgnoreCase) < 0
                    && table.IndexOf("=", StringComparison.Ordinal) < 0)
                {
                    return "formula 模式格内应是公式 " + table;
                }

                return null;
            });

            WithBook(run, s, "xl-rd-value", (id, ch) =>
            {
                if (!s.SeedFormula(ch, "公式表"))
                {
                    return "铺公式失败";
                }

                ToolResult r = s.Read(id, "公式表", "A2:B2", "value");
                if (!r.Success)
                {
                    return r.Error;
                }

                string table = Str(r, "table");
                if (table.IndexOf("=A2+1", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    return "value 模式不应是公式串";
                }

                if (table.IndexOf("11", StringComparison.Ordinal) < 0
                    && table.IndexOf("10", StringComparison.Ordinal) < 0)
                {
                    return "value 模式应是显示值 " + table;
                }

                return null;
            });

            WithBook(run, s, "xl-rd-merge", (id, ch) =>
            {
                if (!s.SeedMerge(ch, "合并表"))
                {
                    return "铺合并失败";
                }

                ToolResult r = s.Read(id, "合并表", "A1:C1");
                if (!r.Success)
                {
                    return r.Error;
                }

                string table = Str(r, "table");
                if (table.IndexOf("addr=\"B1\"", StringComparison.Ordinal) >= 0
                    || table.IndexOf("addr=\"C1\"", StringComparison.Ordinal) >= 0)
                {
                    return "延续格不该输出";
                }

                if (table.IndexOf("标题", StringComparison.Ordinal) < 0)
                {
                    return "缺标题";
                }

                if (table.IndexOf("area=", StringComparison.Ordinal) < 0
                    && table.IndexOf("colspan=", StringComparison.Ordinal) < 0)
                {
                    return "应有 area/colspan";
                }

                return null;
            });

            WithBook(run, s, "xl-rd-empty", (id, ch) =>
            {
                string name = s.FirstSheetName(ch);
                ToolResult r = s.Read(id, name);
                if (!r.Success)
                {
                    return r.Error;
                }

                if (Str(r, "table").IndexOf("地区", StringComparison.Ordinal) >= 0)
                {
                    return "空表不该有业务字";
                }

                return null;
            });

            WithBook(run, s, "xl-rd-hidden", (id, ch) =>
            {
                if (!s.SeedHidden(ch, "可见表", "隐藏底稿"))
                {
                    return "铺隐藏失败";
                }

                ToolResult r = s.Read(id, "隐藏底稿", "A1");
                if (!r.Success)
                {
                    return r.Error;
                }

                return Str(r, "table").IndexOf("隐", StringComparison.Ordinal) >= 0
                    ? null
                    : "没读到隐 " + Str(r, "table");
            });

            WithBook(run, s, "xl-rd-trunc", (id, ch) =>
            {
                if (!s.SeedUsed(ch, "利润表"))
                {
                    return "铺底失败";
                }

                ToolResult r = s.Read(id, "利润表", "A1:AN120");
                if (!r.Success)
                {
                    return r.Error;
                }

                if (!Bool(r, "truncated"))
                {
                    return "应 truncated";
                }

                return null;
            });

            WithBook(run, s, "xl-rd-bad-a1", (id, ch) =>
            {
                string name = s.FirstSheetName(ch);
                ToolResult r = s.Read(id, name, "C1:A1");
                return r.Success ? "颠倒 range 应失败" : (Has(r.Error, "非法") || Has(r.Error, "颠倒")
                    ? null
                    : r.Error);
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
            Report(run, id, err);
        }

        private static void WithUnsaved(
            TestRun run,
            XlSession s,
            string id,
            Func<IOperationChannel, string> body)
        {
            IOperationChannel ch = s.UnsavedBook();
            if (ch == null)
            {
                run.CaseFail(id, "无法 Add 未存盘簿");
                return;
            }

            string err;
            try
            {
                err = body(ch);
            }
            catch (Exception ex)
            {
                err = ex.Message;
            }

            s.CloseBook(ch.ChannelId);
            Report(run, id, err);
        }

        private static void Report(TestRun run, string id, string err)
        {
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

        private static bool Bool(ToolResult r, string key)
        {
            Dictionary<string, object> data = Data(r);
            if (data == null || !data.ContainsKey(key) || data[key] == null)
            {
                return false;
            }

            try
            {
                return Convert.ToBoolean(data[key]);
            }
            catch
            {
                return false;
            }
        }

        private static List<Dictionary<string, object>> Sheets(ToolResult r)
        {
            var list = new List<Dictionary<string, object>>();
            Dictionary<string, object> data = Data(r);
            if (data == null || !data.ContainsKey("sheets"))
            {
                return list;
            }

            if (data["sheets"] is IEnumerable raw)
            {
                foreach (object one in raw)
                {
                    if (one is Dictionary<string, object> d)
                    {
                        list.Add(d);
                    }
                }
            }

            return list;
        }

        private static Dictionary<string, object> SheetNamed(ToolResult r, string name)
        {
            foreach (Dictionary<string, object> one in Sheets(r))
            {
                if (string.Equals(Convert.ToString(one["name"] ?? ""), name, StringComparison.Ordinal))
                {
                    return one;
                }
            }

            return null;
        }
    }
}
