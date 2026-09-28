using System;
using System.Collections;
using System.Collections.Generic;
using WordAddIn1;
using WordAddIn1.SpreadsheetHost;

namespace ExcelWorkbookRoundtripTest
{
    internal static class XlmCases
    {
        public static void Run(TestRun run, bool et)
        {
            var s = new XlSession(et);
            RunNeg(run, s);
            RunAddRename(run, s);
            RunDeleteExtra(run, s);
        }

        private static void RunNeg(TestRun run, XlSession s)
        {
            ChannelRegistry.ClearAll();
            ToolResult noCh = s.Manage(null, "add", name: "分析");
            ExpectFail(run, "xlm-neg-no-channel", noCh, "打开");

            var req = new SpreadsheetManageSheetRequest
            {
                Action = "add",
                Name = "分析"
            };
            if (!SpreadsheetHostAdapter.TryManageSheet(new FakeWordChannel(), req, out _, out ToolResult wrong))
            {
                ExpectFail(run, "xlm-neg-wrong-kind", wrong, "unsupported");
            }
            else
            {
                run.CaseFail("xlm-neg-wrong-kind", "word 渠道应失败");
            }

            WithBook(run, s, "xlm-neg-no-action", (id, ch) =>
            {
                ToolResult r = s.Call("F_manage_excel_sheet", new Dictionary<string, object>
                {
                    ["channel_id"] = id,
                    ["name"] = "分析"
                });
                return r.Success ? "缺 action 应失败" : (Has(r.Error, "action") ? null : r.Error);
            });

            WithBook(run, s, "xlm-neg-bad-action", (id, ch) =>
            {
                ToolResult r = s.Manage(id, "merge", name: "分析");
                return r.Success ? "非法 action 应失败" : (Has(r.Error, "add") || Has(r.Error, "action") ? null : r.Error);
            });

            WithBook(run, s, "xlm-neg-add-no-name", (id, ch) =>
            {
                ToolResult r = s.Manage(id, "add");
                return r.Success ? "缺 name 应失败" : (Has(r.Error, "空") || Has(r.Error, "name") ? null : r.Error);
            });

            WithBook(run, s, "xlm-neg-rename-no-sheet", (id, ch) =>
            {
                ToolResult r = s.Manage(id, "rename", name: "新名");
                return r.Success ? "缺 sheet 应失败" : (Has(r.Error, "sheet") ? null : r.Error);
            });

            WithBook(run, s, "xlm-neg-delete-no-confirm", (id, ch) =>
            {
                if (!s.SeedTwoSheets(ch, "留", "删"))
                {
                    return "铺两表失败";
                }

                ToolResult r = s.Manage(id, "delete", sheet: "删", confirm: false);
                if (r.Success)
                {
                    return "无 confirm 应失败";
                }

                if (!Has(r.Error, "confirm"))
                {
                    return r.Error;
                }

                ToolResult ov = s.Overview(id);
                return ov.Success && SheetNamed(ov, "删") != null
                    ? null
                    : "表被删了 " + (ov.Error ?? "概览无删");
            });

            WithBook(run, s, "xlm-neg-delete-last", (id, ch) =>
            {
                string name = s.FirstSheetName(ch);
                ToolResult r = s.Manage(id, "delete", sheet: name, confirm: true);
                if (r.Success)
                {
                    return "唯一表应删失败";
                }

                if (!Has(r.Error, "唯一"))
                {
                    return r.Error;
                }

                ToolResult ov = s.Overview(id);
                return ov.Success && SheetNamed(ov, name) != null
                    ? null
                    : "唯一表没了";
            });

            WithBook(run, s, "xlm-neg-missing-sheet", (id, ch) =>
            {
                if (!s.SeedTwoSheets(ch, "利润表", "底稿"))
                {
                    return "铺两表失败";
                }

                ToolResult r = s.Manage(id, "delete", sheet: "不存在的表", confirm: true);
                return r.Success ? "缺表应失败" : (Has(r.Error, "不存在") ? null : r.Error);
            });

            WithBook(run, s, "xlm-neg-dup-add", (id, ch) =>
            {
                if (!s.SeedUsed(ch, "利润表"))
                {
                    return "铺底失败";
                }

                ToolResult r = s.Manage(id, "add", name: "利润表");
                if (r.Success)
                {
                    return "重名 add 应失败";
                }

                return Has(r.Error, "已存在") ? null : r.Error;
            });
        }

        private static void RunAddRename(TestRun run, XlSession s)
        {
            WithBook(run, s, "xlm-add-tail", (id, ch) =>
            {
                ToolResult r = s.Manage(id, "add", name: "分析");
                if (!r.Success)
                {
                    return r.Error;
                }

                if (!string.Equals(Str(r, "sheet"), "分析", StringComparison.Ordinal)
                    || !NameList(r).Contains("分析"))
                {
                    return "回包 sheet/sheets 不对";
                }

                ToolResult ov = s.Overview(id);
                if (!ov.Success)
                {
                    return ov.Error;
                }

                Dictionary<string, object> added = SheetNamed(ov, "分析");
                if (added == null)
                {
                    return "概览没有分析";
                }

                return !IsHidden(added)
                    ? null
                    : "新表被隐藏了";
            });

            WithBook(run, s, "xlm-add-before", (id, ch) =>
            {
                string first = s.FirstSheetName(ch);
                ToolResult renamed = s.Manage(id, "rename", name: "利润表", sheet: first);
                if (!renamed.Success)
                {
                    return renamed.Error;
                }

                ToolResult r = s.Manage(id, "add", name: "汇总", before: "利润表");
                if (!r.Success)
                {
                    return r.Error;
                }

                List<string> names = NameList(r);
                int iSum = names.IndexOf("汇总");
                int iProfit = names.IndexOf("利润表");
                return iSum >= 0 && iProfit >= 0 && iSum < iProfit
                    ? null
                    : "顺序 " + string.Join(",", names);
            });

            WithBook(run, s, "xlm-add-zh", (id, ch) =>
            {
                ToolResult r = s.Manage(id, "add", name: "底稿");
                return r.Success && string.Equals(Str(r, "sheet"), "底稿", StringComparison.Ordinal)
                    ? null
                    : (r.Success ? "sheet=" + Str(r, "sheet") : r.Error);
            });

            WithBook(run, s, "xlm-neg-add-illegal", (id, ch) =>
            {
                ToolResult r = s.Manage(id, "add", name: "a/b");
                return r.Success ? "非法名应失败" : (Has(r.Error, "非法") ? null : r.Error);
            });

            WithBook(run, s, "xlm-neg-add-long", (id, ch) =>
            {
                ToolResult r = s.Manage(id, "add", name: new string('甲', 32));
                return r.Success ? "超长应失败" : (Has(r.Error, "31") || Has(r.Error, "过长") ? null : r.Error);
            });

            WithBook(run, s, "xlm-neg-before-missing", (id, ch) =>
            {
                ToolResult r = s.Manage(id, "add", name: "汇总", before: "不存在的表");
                return r.Success ? "缺 before 应失败" : (Has(r.Error, "不存在") ? null : r.Error);
            });

            WithBook(run, s, "xlm-neg-before-chart", (id, ch) =>
            {
                if (!s.SeedChartSheet(ch, out string chartName))
                {
                    return "铺图表 sheet 失败";
                }

                ToolResult r = s.Manage(id, "add", name: "汇总", before: chartName);
                return r.Success ? "before 图表应失败" : (Has(r.Error, "普通") || Has(r.Error, "图表") ? null : r.Error);
            });

            WithBook(run, s, "xlm-rename", (id, ch) =>
            {
                string old = s.FirstSheetName(ch);
                if (!s.SeedCell(ch, old, "A1", "甲"))
                {
                    return "铺底失败";
                }

                ToolResult r = s.Manage(id, "rename", name: "原始数据", sheet: old);
                if (!r.Success)
                {
                    return r.Error;
                }

                List<string> names = NameList(r);
                if (names.Contains(old) || !names.Contains("原始数据"))
                {
                    return "sheets " + string.Join(",", names);
                }

                ToolResult rd = s.Read(id, "原始数据", "A1");
                if (!rd.Success || !Has(Str(rd, "table"), "甲"))
                {
                    return rd.Success ? "值丢了 " + Str(rd, "table") : rd.Error;
                }

                ToolResult oldRd = s.Read(id, old, "A1");
                return oldRd.Success ? "旧名还应可读" : null;
            });

            WithBook(run, s, "xlm-neg-rename-same", (id, ch) =>
            {
                string name = s.FirstSheetName(ch);
                ToolResult r = s.Manage(id, "rename", name: name, sheet: name);
                return r.Success ? "同名 rename 应失败" : (Has(r.Error, "相同") ? null : r.Error);
            });

            WithBook(run, s, "xlm-neg-rename-dup", (id, ch) =>
            {
                if (!s.SeedTwoSheets(ch, "利润表", "底稿"))
                {
                    return "铺两表失败";
                }

                ToolResult r = s.Manage(id, "rename", name: "利润表", sheet: "底稿");
                if (r.Success)
                {
                    return "重名 rename 应失败";
                }

                if (!Has(r.Error, "已存在"))
                {
                    return r.Error;
                }

                ToolResult ov = s.Overview(id);
                return ov.Success && SheetNamed(ov, "利润表") != null && SheetNamed(ov, "底稿") != null
                    ? null
                    : "表名被改掉了";
            });
        }

        private static void RunDeleteExtra(TestRun run, XlSession s)
        {
            WithBook(run, s, "xlm-delete", (id, ch) =>
            {
                if (!s.SeedTwoSheets(ch, "留", "删") || !s.SeedCell(ch, "删", "A1", "废"))
                {
                    return "铺底失败";
                }

                ToolResult r = s.Manage(id, "delete", sheet: "删", confirm: true);
                if (!r.Success)
                {
                    return r.Error;
                }

                if (!string.Equals(Str(r, "sheet"), "删", StringComparison.Ordinal))
                {
                    return "sheet=" + Str(r, "sheet");
                }

                List<string> names = NameList(r);
                if (names.Contains("删") || !names.Contains("留"))
                {
                    return "sheets " + string.Join(",", names);
                }

                ToolResult gone = s.Read(id, "删", "A1");
                return gone.Success ? "删表还能读" : null;
            });

            WithBook(run, s, "xlm-delete-keep", (id, ch) =>
            {
                if (!s.SeedTwoSheets(ch, "留", "删") || !s.SeedCell(ch, "留", "A1", "留字"))
                {
                    return "铺底失败";
                }

                ToolResult r = s.Manage(id, "delete", sheet: "删", confirm: true);
                if (!r.Success)
                {
                    return r.Error;
                }

                ToolResult rd = s.Read(id, "留", "A1");
                return rd.Success && Has(Str(rd, "table"), "留字")
                    ? null
                    : (rd.Error ?? ("留表被清 " + Str(rd, "table")));
            });

            WithBook(run, s, "xlm-neg-delete-no-sheet", (id, ch) =>
            {
                ToolResult r = s.Call("F_manage_excel_sheet", new Dictionary<string, object>
                {
                    ["channel_id"] = id,
                    ["action"] = "delete",
                    ["confirm"] = true
                });
                return r.Success ? "缺 sheet 应失败" : (Has(r.Error, "sheet") ? null : r.Error);
            });

            WithBook(run, s, "xlm-rename-hidden", (id, ch) =>
            {
                if (!s.SeedHidden(ch, "可见表", "隐藏底稿"))
                {
                    return "铺隐藏表失败";
                }

                ToolResult r = s.Manage(id, "rename", name: "归档", sheet: "隐藏底稿");
                if (!r.Success)
                {
                    return r.Error;
                }

                ToolResult ov = s.Overview(id);
                if (!ov.Success)
                {
                    return ov.Error;
                }

                if (SheetNamed(ov, "隐藏底稿") != null)
                {
                    return "旧名还在";
                }

                Dictionary<string, object> archived = SheetNamed(ov, "归档");
                return archived != null && IsHidden(archived)
                    ? null
                    : "归档未列出或未隐藏";
            });

            WithBook(run, s, "xlm-delete-hidden", (id, ch) =>
            {
                if (!s.SeedHidden(ch, "可见表", "隐藏底稿"))
                {
                    return "铺隐藏表失败";
                }

                ToolResult r = s.Manage(id, "delete", sheet: "隐藏底稿", confirm: true);
                if (!r.Success)
                {
                    return r.Error;
                }

                List<string> names = NameList(r);
                return !names.Contains("隐藏底稿") && names.Contains("可见表")
                    ? null
                    : "sheets " + string.Join(",", names);
            });

            WithBook(run, s, "xlm-rename-chart", (id, ch) =>
            {
                if (!s.SeedChartSheet(ch, out string chartName))
                {
                    return "铺图表 sheet 失败";
                }

                ToolResult r = s.Manage(id, "rename", name: "图页", sheet: chartName);
                if (!r.Success)
                {
                    return r.Error;
                }

                List<string> names = NameList(r);
                if (names.Contains(chartName) || !names.Contains("图页"))
                {
                    return "sheets " + string.Join(",", names);
                }

                ToolResult rd = s.Read(id, "图页", "A1");
                return rd.Success ? "图表 sheet 读区域应失败" : (Has(rd.Error, "chart") ? null : rd.Error);
            });

            WithBook(run, s, "xlm-delete-chart", (id, ch) =>
            {
                string keep = s.FirstSheetName(ch);
                if (!s.SeedChartSheet(ch, out string chartName))
                {
                    return "铺图表 sheet 失败";
                }

                ToolResult r = s.Manage(id, "delete", sheet: chartName, confirm: true);
                if (!r.Success)
                {
                    return r.Error;
                }

                List<string> names = NameList(r);
                return !names.Contains(chartName) && names.Contains(keep)
                    ? null
                    : "sheets " + string.Join(",", names);
            });

            WithBook(run, s, "xlm-add-overview", (id, ch) =>
            {
                ToolResult r = s.Manage(id, "add", name: "分析");
                if (!r.Success)
                {
                    return r.Error;
                }

                ToolResult ov = s.Overview(id);
                if (!ov.Success)
                {
                    return ov.Error;
                }

                Dictionary<string, object> added = SheetNamed(ov, "分析");
                if (added == null)
                {
                    return "概览没有分析";
                }

                string typ = Convert.ToString(added.ContainsKey("sheet_type") ? added["sheet_type"] : "") ?? "";
                return !string.Equals(typ, "chart", StringComparison.OrdinalIgnoreCase)
                    ? null
                    : "sheet_type=" + typ;
            });

            WithBook(run, s, "xlm-rename-then-read", (id, ch) =>
            {
                if (!s.SeedUsed(ch, "利润表") || !s.SeedCell(ch, "利润表", "A1", "甲"))
                {
                    return "铺底失败";
                }

                ToolResult r = s.Manage(id, "rename", name: "明细", sheet: "利润表");
                if (!r.Success)
                {
                    return r.Error;
                }

                ToolResult rd = s.Read(id, "明细", "A1");
                if (!rd.Success || !Has(Str(rd, "table"), "甲"))
                {
                    return rd.Success ? "明细无甲 " + Str(rd, "table") : rd.Error;
                }

                ToolResult oldRd = s.Read(id, "利润表", "A1");
                return oldRd.Success ? "旧名利润表还应可读" : null;
            });

            WithBook(run, s, "xlm-add-then-write", (id, ch) =>
            {
                ToolResult r = s.Manage(id, "add", name: "底稿");
                if (!r.Success)
                {
                    return r.Error;
                }

                if (!s.SeedCell(ch, "底稿", "A1", "格"))
                {
                    return "写底稿失败";
                }

                ToolResult rd = s.Read(id, "底稿", "A1");
                return rd.Success && Has(Str(rd, "table"), "格")
                    ? null
                    : (rd.Error ?? ("读不到格 " + Str(rd, "table")));
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

        private static List<string> NameList(ToolResult r)
        {
            var list = new List<string>();
            Dictionary<string, object> data = Data(r);
            if (data == null || !data.ContainsKey("sheets") || data["sheets"] == null)
            {
                return list;
            }

            if (data["sheets"] is IEnumerable raw)
            {
                foreach (object one in raw)
                {
                    if (one == null)
                    {
                        continue;
                    }

                    list.Add(Convert.ToString(one) ?? "");
                }
            }

            return list;
        }

        private static List<Dictionary<string, object>> Sheets(ToolResult r)
        {
            var list = new List<Dictionary<string, object>>();
            Dictionary<string, object> data = Data(r);
            if (data == null || !data.ContainsKey("sheets") || !(data["sheets"] is IEnumerable raw))
            {
                return list;
            }

            foreach (object one in raw)
            {
                if (one is Dictionary<string, object> d)
                {
                    list.Add(d);
                }
            }

            return list;
        }

        private static Dictionary<string, object> SheetNamed(ToolResult r, string name)
        {
            foreach (Dictionary<string, object> one in Sheets(r))
            {
                if (string.Equals(Convert.ToString(one.ContainsKey("name") ? one["name"] : ""), name, StringComparison.Ordinal))
                {
                    return one;
                }
            }

            return null;
        }

        private static bool IsHidden(Dictionary<string, object> sheet)
        {
            if (sheet == null || !sheet.ContainsKey("hidden") || sheet["hidden"] == null)
            {
                return false;
            }

            try
            {
                return Convert.ToBoolean(sheet["hidden"]);
            }
            catch
            {
                return false;
            }
        }
    }
}
