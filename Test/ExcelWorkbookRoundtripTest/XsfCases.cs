using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using WordAddIn1;
using WordAddIn1.OpenFiles;
using WordAddIn1.SpreadsheetHost;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelWorkbookRoundtripTest
{
    /// <summary>排序和筛选。前缀 xsf-。不要并进已绿的 xl- / xlw- / xlp-。</summary>
    internal static class XsfCases
    {
        public const int Target = 24;

        public static void Run(TestRun run, bool et, string filter)
        {
            var session = new XlSession(et);
            RunContract(run, session, filter);
            RunSheet(run, session, filter);
        }

        private static void RunContract(TestRun run, XlSession session, string filter)
        {
            Gate(run, filter, "xsf-neg-word", () =>
            {
                var request = new SpreadsheetSortFilterRequest
                {
                    Action = "list",
                    Sheet = "Sheet1",
                    Sorts = new List<SortFilterSortKey>(),
                    Filters = new List<SortFilterColumn>()
                };
                if (SpreadsheetHostAdapter.TrySortFilter(new FakeWordChannel(), request, out _, out ToolResult wrong))
                {
                    return "word 渠道应失败";
                }

                return Has(wrong == null ? null : wrong.Error, "excel/et") ? null : (wrong == null ? "无回包" : wrong.Error);
            });
        }

        private static void RunSheet(TestRun run, XlSession session, string filter)
        {
            Book(run, session, filter, "xsf-neg-action", (id, ch, sheet) =>
            {
                ToolResult result = Call(session, id, sheet, "sort");
                return result.Success ? "非法 action 应失败" : (Has(result.Error, "action") ? null : result.Error);
            });

            Book(run, session, filter, "xsf-neg-range", (id, ch, sheet) =>
            {
                ToolResult result = Call(session, id, sheet, "add", "C1:F2");
                return result.Success ? "跨行 range 应失败" : (Has(result.Error, "同一行") ? null : result.Error);
            });

            Book(run, session, filter, "xsf-neg-expr", (id, ch, sheet) =>
            {
                string seeded = SeedMain(session, ch, sheet);
                if (seeded != null)
                {
                    return seeded;
                }

                ToolResult result = Call(session, id, sheet, "add", "D1:D1", filters: new object[]
                {
                    FilterExpr("金额", ">=100 AND <=500")
                });
                if (result.Success)
                {
                    return "大写 AND 应失败";
                }

                if (!Has(result.Error, "小写"))
                {
                    return result.Error;
                }

                return PanelCount(Call(session, id, sheet, "list")) == 0 ? null : "失败仍写入了面板";
            });

            Book(run, session, filter, "xsf-neg-both", (id, ch, sheet) =>
            {
                ToolResult result = Call(session, id, sheet, "add", "C1:C1", filters: new object[]
                {
                    new Dictionary<string, object>
                    {
                        ["column"] = "地区",
                        ["in"] = new List<string> { "华东" },
                        ["expr"] = "!=\"华南\""
                    }
                });
                return result.Success ? "in 与 expr 应失败" : (Has(result.Error, "不能同时") ? null : result.Error);
            });

            Book(run, session, filter, "xsf-neg-index", (id, ch, sheet) =>
            {
                ToolResult result = Call(session, id, sheet, "add", "C1:C1", index: 1);
                return result.Success ? "add 带 index 应失败" : (Has(result.Error, "index") ? null : result.Error);
            });

            Book(run, session, filter, "xsf-list-empty", (id, ch, sheet) =>
            {
                ToolResult result = Call(session, id, sheet, "list");
                if (!result.Success)
                {
                    return result.Error;
                }

                return PanelCount(result) == 0 ? null : "空表不应有面板";
            });

            Book(run, session, filter, "xsf-add-in", (id, ch, sheet) =>
            {
                string seeded = SeedMain(session, ch, sheet);
                if (seeded != null)
                {
                    return seeded;
                }

                ToolResult result = Call(session, id, sheet, "add", "C1:C1", filters: new object[]
                {
                    FilterIn("地区", "华东")
                });
                if (!result.Success)
                {
                    return result.Error;
                }

                if (IntOf(result, "visible_rows") != 2 || IntOf(result, "data_rows") != 4)
                {
                    return "可见 " + IntOf(result, "visible_rows") + " 数据 " + IntOf(result, "data_rows");
                }

                if (!string.Equals(Str(result, "range"), "C1", StringComparison.Ordinal))
                {
                    return "range=" + Str(result, "range");
                }

                if (!RowHidden(ch, sheet, 3) || RowHidden(ch, sheet, 2))
                {
                    return "华北应隐藏，华东应可见";
                }

                ToolResult read = session.Read(id, sheet, "C2:C5");
                if (!read.Success || !Has(Str(read, "table"), "华北"))
                {
                    return "隐藏行应仍能读到 华北";
                }

                return IsNum(session.ReadValue2(ch, sheet, "A2"), 1) ? null : "筛选改动了编号列";
            });

            Book(run, session, filter, "xsf-sort-part", (id, ch, sheet) =>
            {
                string seeded = SeedMain(session, ch, sheet);
                if (seeded != null)
                {
                    return seeded;
                }

                ToolResult result = Call(session, id, sheet, "add", "D1:D1", sorts: new object[]
                {
                    SortKey("金额", "desc")
                });
                if (!result.Success)
                {
                    return result.Error;
                }

                if (CountOf(PanelField(result, "sorts")) != 0)
                {
                    return "工作表筛选不应编造 sorts";
                }

                if (!IsNum(session.ReadValue2(ch, sheet, "D2"), 3000))
                {
                    return "金额列没有倒序";
                }

                if (!IsNum(session.ReadValue2(ch, sheet, "A2"), 1)
                    || !string.Equals(Text(session.ReadValue2(ch, sheet, "B2")), "甲", StringComparison.Ordinal))
                {
                    return "range 外的编号/姓名跟着动了";
                }

                return null;
            });

            Book(run, session, filter, "xsf-sort-all", (id, ch, sheet) =>
            {
                string seeded = SeedMain(session, ch, sheet);
                if (seeded != null)
                {
                    return seeded;
                }

                ToolResult result = Call(session, id, sheet, "add", "A1:D1", sorts: new object[]
                {
                    SortKey("金额", "desc")
                });
                if (!result.Success)
                {
                    return result.Error;
                }

                if (!IsNum(session.ReadValue2(ch, sheet, "A2"), 2)
                    || !string.Equals(Text(session.ReadValue2(ch, sheet, "B2")), "乙", StringComparison.Ordinal)
                    || !string.Equals(Text(session.ReadValue2(ch, sheet, "C2")), "华北", StringComparison.Ordinal)
                    || !IsNum(session.ReadValue2(ch, sheet, "D2"), 3000))
                {
                    return "整行没有按金额倒序";
                }

                return null;
            });

            Book(run, session, filter, "xsf-expr-and", (id, ch, sheet) =>
            {
                string seeded = SeedMain(session, ch, sheet);
                if (seeded != null)
                {
                    return seeded;
                }

                ToolResult result = Call(session, id, sheet, "add", "D1:D1", filters: new object[]
                {
                    FilterExpr("金额", ">=100 and <=500")
                });
                if (!result.Success)
                {
                    return result.Error;
                }

                if (IntOf(result, "visible_rows") != 1)
                {
                    return "介于应只剩 100，可见 " + IntOf(result, "visible_rows");
                }

                string expr = FilterText(result, "expr");
                return expr != null && expr.IndexOf(">=100", StringComparison.Ordinal) >= 0
                    && expr.IndexOf("and", StringComparison.Ordinal) >= 0
                    && expr.IndexOf("<=500", StringComparison.Ordinal) >= 0
                    ? null
                    : "expr=" + expr;
            });

            Book(run, session, filter, "xsf-expr-or", (id, ch, sheet) =>
            {
                string seeded = SeedMain(session, ch, sheet);
                if (seeded != null)
                {
                    return seeded;
                }

                ToolResult result = Call(session, id, sheet, "add", "D1:D1", filters: new object[]
                {
                    FilterExpr("金额", "<=0 or >=1000")
                });
                if (!result.Success)
                {
                    return result.Error;
                }

                if (IntOf(result, "visible_rows") != 2)
                {
                    return "两端应剩两行，可见 " + IntOf(result, "visible_rows");
                }

                string expr = FilterText(result, "expr");
                return expr != null && expr.IndexOf(" or ", StringComparison.Ordinal) >= 0
                    ? null
                    : "expr=" + expr;
            });

            Book(run, session, filter, "xsf-sort-then", (id, ch, sheet) =>
            {
                string seeded = SeedMain(session, ch, sheet);
                if (seeded != null)
                {
                    return seeded;
                }

                ToolResult result = Call(session, id, sheet, "add", "C1:D1",
                    sorts: new object[] { SortKey("金额", "desc") },
                    filters: new object[] { FilterIn("地区", "华东") });
                if (!result.Success)
                {
                    return result.Error;
                }

                if (IntOf(result, "visible_rows") != 2)
                {
                    return "筛后应剩两个华东，可见 " + IntOf(result, "visible_rows");
                }

                if (!IsNum(session.ReadValue2(ch, sheet, "A2"), 1))
                {
                    return "编号列跟着排了";
                }

                if (!IsNum(session.ReadValue2(ch, sheet, "D2"), 3000)
                    || !string.Equals(Text(session.ReadValue2(ch, sheet, "C2")), "华北", StringComparison.Ordinal))
                {
                    return "先筛后排：首行应仍是排完的华北/3000";
                }

                if (!IsNum(session.ReadValue2(ch, sheet, "D4"), 100))
                {
                    return "华东 100 不在排序后的位置";
                }

                return null;
            });

            Book(run, session, filter, "xsf-update", (id, ch, sheet) =>
            {
                string seeded = SeedMain(session, ch, sheet);
                if (seeded != null)
                {
                    return seeded;
                }

                ToolResult added = Call(session, id, sheet, "add", "C1:C1", filters: new object[]
                {
                    FilterIn("地区", "华东")
                });
                if (!added.Success)
                {
                    return added.Error;
                }

                ToolResult updated = Call(session, id, sheet, "update", "C1:C1", index: 1, filters: new object[]
                {
                    FilterIn("地区", "华北")
                });
                if (!updated.Success)
                {
                    return updated.Error;
                }

                if (IntOf(updated, "visible_rows") != 1 || IntOf(updated, "index") != 1)
                {
                    return "update 后可见 " + IntOf(updated, "visible_rows");
                }

                string values = FilterText(updated, "in");
                if (values == null || values.IndexOf("华北", StringComparison.Ordinal) < 0
                    || values.IndexOf("华东", StringComparison.Ordinal) >= 0)
                {
                    return "筛选仍是旧值 " + values;
                }

                return RowHidden(ch, sheet, 2) && !RowHidden(ch, sheet, 3) ? null : "隐藏行没有换成华北";
            });

            Book(run, session, filter, "xsf-clear-order", (id, ch, sheet) =>
            {
                string seeded = SeedMain(session, ch, sheet);
                if (seeded != null)
                {
                    return seeded;
                }

                ToolResult added = Call(session, id, sheet, "add", "A1:D1", sorts: new object[]
                {
                    SortKey("金额", "desc")
                });
                if (!added.Success)
                {
                    return added.Error;
                }

                ToolResult cleared = Call(session, id, sheet, "clear", index: 1);
                if (!cleared.Success)
                {
                    return cleared.Error;
                }

                if (PanelCount(Call(session, id, sheet, "list")) != 0)
                {
                    return "clear 后面板还在";
                }

                return IsNum(session.ReadValue2(ch, sheet, "A2"), 2)
                    ? null
                    : "clear 把排序恢复成了原来的行序";
            });

            Book(run, session, filter, "xsf-replace", (id, ch, sheet) =>
            {
                string seeded = SeedMain(session, ch, sheet);
                if (seeded != null)
                {
                    return seeded;
                }

                ToolResult first = Call(session, id, sheet, "add", "A1:B1", filters: new object[]
                {
                    FilterIn("姓名", "甲")
                });
                if (!first.Success)
                {
                    return first.Error;
                }

                ToolResult second = Call(session, id, sheet, "add", "C1:D1", filters: new object[]
                {
                    FilterIn("地区", "华东")
                });
                if (!second.Success)
                {
                    return second.Error;
                }

                ToolResult list = Call(session, id, sheet, "list");
                if (PanelCount(list) != 1)
                {
                    return "普通区域应只留一套，现有 " + PanelCount(list);
                }

                Dictionary<string, object> panel = PanelAt(list, 0);
                return string.Equals(Text(panel["range"]), "C1:D1", StringComparison.Ordinal)
                    ? null
                    : "留下的 range=" + Text(panel["range"]);
            });

            Book(run, session, filter, "xsf-keeps-table", (id, ch, sheet) =>
            {
                string seeded = SeedMain(session, ch, sheet);
                if (seeded != null)
                {
                    return seeded;
                }

                string table = AddTable(ch, sheet, "C1:D5");
                if (table != null)
                {
                    return table;
                }

                ToolResult added = Call(session, id, sheet, "add", "A1:B1", filters: new object[]
                {
                    FilterIn("姓名", "甲")
                });
                if (!added.Success)
                {
                    return added.Error;
                }

                ToolResult list = Call(session, id, sheet, "list");
                if (PanelCount(list) != 2)
                {
                    return "表和普通区域应各留一套，现有 " + PanelCount(list);
                }

                return string.Equals(Text(PanelAt(list, 0)["range"]), "A1:B1", StringComparison.Ordinal)
                    && string.Equals(Text(PanelAt(list, 1)["range"]), "C1:D1", StringComparison.Ordinal)
                    ? null
                    : "序号或 range 不对";
            });

            Book(run, session, filter, "xsf-two", (id, ch, sheet) =>
            {
                string seeded = SeedPair(session, ch, sheet);
                if (seeded != null)
                {
                    return seeded;
                }

                string left = AddTable(ch, sheet, "A1:B5");
                if (left != null)
                {
                    return left;
                }

                string right = AddTable(ch, sheet, "D1:E5");
                if (right != null)
                {
                    return right;
                }

                ToolResult list = Call(session, id, sheet, "list");
                if (PanelCount(list) != 2)
                {
                    return "两张插入表应有两块面板，现有 " + PanelCount(list);
                }

                if (!string.Equals(Text(PanelAt(list, 0)["range"]), "A1:B1", StringComparison.Ordinal)
                    || !string.Equals(Text(PanelAt(list, 1)["range"]), "D1:E1", StringComparison.Ordinal))
                {
                    return "序号没有按从左到右";
                }

                ToolResult cleared = Call(session, id, sheet, "clear", index: 1);
                if (!cleared.Success)
                {
                    return cleared.Error;
                }

                ToolResult again = Call(session, id, sheet, "list");
                if (PanelCount(again) != 1)
                {
                    return "清掉左边后应只剩一块";
                }

                return string.Equals(Text(PanelAt(again, 0)["range"]), "D1:E1", StringComparison.Ordinal)
                    ? null
                    : "右边的表也被清了";
            });

            Book(run, session, filter, "xsf-table-partial", (id, ch, sheet) =>
            {
                string seeded = SeedMain(session, ch, sheet);
                if (seeded != null)
                {
                    return seeded;
                }

                string table = AddTable(ch, sheet, "A1:D5");
                if (table != null)
                {
                    return table;
                }

                ToolResult result = Call(session, id, sheet, "add", "C1:D1", filters: new object[]
                {
                    FilterIn("地区", "华东")
                });
                if (result.Success || !Has(result.Error, "整行"))
                {
                    return result.Success ? "半行表头应失败" : result.Error;
                }

                ToolResult list = Call(session, id, sheet, "list");
                if (PanelCount(list) != 1
                    || !string.Equals(Text(PanelAt(list, 0)["range"]), "A1:D1", StringComparison.Ordinal))
                {
                    return "失败不应改掉原表的整行表头";
                }

                return IsNum(session.ReadValue2(ch, sheet, "A2"), 1) ? null : "失败仍排序了";
            });

            Book(run, session, filter, "xsf-table-sort", (id, ch, sheet) =>
            {
                string seeded = SeedMain(session, ch, sheet);
                if (seeded != null)
                {
                    return seeded;
                }

                string table = AddTable(ch, sheet, "A1:D5");
                if (table != null)
                {
                    return table;
                }

                ToolResult result = Call(session, id, sheet, "add", "A1:D1", sorts: new object[]
                {
                    SortKey("金额", "desc")
                });
                if (!result.Success)
                {
                    return result.Error;
                }

                if (!IsNum(session.ReadValue2(ch, sheet, "A2"), 2)
                    || !IsNum(session.ReadValue2(ch, sheet, "D2"), 3000))
                {
                    return "插入表整行没有跟着排序";
                }

                object sorts = PanelField(result, "sorts");
                int count = CountOf(sorts);
                if (count == 0)
                {
                    return null;
                }

                Dictionary<string, object> key = ItemAt(sorts, 0);
                return key != null
                    && string.Equals(Text(key["column"]), "金额", StringComparison.Ordinal)
                    && string.Equals(Text(key["order"]), "desc", StringComparison.Ordinal)
                    ? null
                    : "sorts 读回了别的键";
            });

            Book(run, session, filter, "xsf-nodata", (id, ch, sheet) =>
            {
                if (!session.SeedCell(ch, sheet, "A1", "地区"))
                {
                    return "铺表头失败";
                }

                ToolResult result = Call(session, id, sheet, "add", "A1:A1", filters: new object[]
                {
                    FilterIn("地区", "华东")
                });
                if (result.Success || !Has(result.Error, "没有数据"))
                {
                    return result.Success ? "没有数据应失败" : result.Error;
                }

                return PanelCount(Call(session, id, sheet, "list")) == 0 ? null : "失败仍打开了筛选";
            });

            Book(run, session, filter, "xsf-bad-col", (id, ch, sheet) =>
            {
                string seeded = SeedMain(session, ch, sheet);
                if (seeded != null)
                {
                    return seeded;
                }

                ToolResult result = Call(session, id, sheet, "add", "C1:D1", filters: new object[]
                {
                    FilterIn("不存在", "华东")
                });
                if (result.Success || !Has(result.Error, "不在这一行"))
                {
                    return result.Success ? "未知列应失败" : result.Error;
                }

                return PanelCount(Call(session, id, sheet, "list")) == 0 ? null : "失败仍写入了面板";
            });

            Book(run, session, filter, "xsf-gap", (id, ch, sheet) =>
            {
                if (!session.SeedCell(ch, sheet, "A1", "地区")
                    || !session.SeedCell(ch, sheet, "A2", "华东")
                    || !session.SeedCell(ch, sheet, "A3", "华北")
                    || !session.SeedCell(ch, sheet, "A5", "华南"))
                {
                    return "铺空行失败";
                }

                ToolResult result = Call(session, id, sheet, "add", "A1:A1", filters: new object[]
                {
                    FilterIn("地区", "华南")
                });
                if (!result.Success)
                {
                    return result.Error;
                }

                if (IntOf(result, "data_rows") != 2 || IntOf(result, "visible_rows") != 0)
                {
                    return "空行以下不应算进面板，数据 " + IntOf(result, "data_rows")
                        + " 可见 " + IntOf(result, "visible_rows");
                }

                return string.Equals(Text(session.ReadValue2(ch, sheet, "A5")), "华南", StringComparison.Ordinal)
                    ? null
                    : "空行下的华南被改了";
            });

            Book(run, session, filter, "xsf-miss", (id, ch, sheet) =>
            {
                string seeded = SeedMain(session, ch, sheet);
                if (seeded != null)
                {
                    return seeded;
                }

                ToolResult result = Call(session, id, sheet, "update", "C1:C1", index: 2, filters: new object[]
                {
                    FilterIn("地区", "华东")
                });
                if (result.Success || !Has(result.Error, "序号不存在"))
                {
                    return result.Success ? "缺序号应失败" : result.Error;
                }

                return PanelCount(Call(session, id, sheet, "list")) == 0 ? null : "失败仍新建了面板";
            });
        }

        private static void Gate(TestRun run, string filter, string id, Func<string> body)
        {
            if (!Want(filter, id))
            {
                return;
            }

            string error;
            try
            {
                error = body();
            }
            catch (Exception ex)
            {
                error = ex.Message;
            }

            if (error == null)
            {
                run.CaseOk(id);
            }
            else
            {
                run.CaseFail(id, error);
            }
        }

        private static void Book(
            TestRun run,
            XlSession session,
            string filter,
            string id,
            Func<string, IOperationChannel, string, string> body)
        {
            if (!Want(filter, id))
            {
                return;
            }

            string path = session.NewXlsxPath();
            ToolResult opened = session.Open(path, true);
            if (!opened.Success)
            {
                run.CaseFail(id, "打开失败: " + opened.Error);
                return;
            }

            string channelId = Str(opened, "channel_id");
            if (!session.TryGetChannel(channelId, out IOperationChannel channel))
            {
                run.CaseFail(id, "渠道不在注册表");
                session.CloseBook(channelId);
                return;
            }

            string sheet = session.FirstSheetName(channel);
            string error;
            try
            {
                error = body(channelId, channel, sheet);
            }
            catch (Exception ex)
            {
                error = ex.Message;
            }

            session.CloseBook(channelId);
            if (error == null)
            {
                run.CaseOk(id);
            }
            else
            {
                run.CaseFail(id, error);
            }
        }

        private static bool Want(string filter, string id)
        {
            if (string.IsNullOrEmpty(filter)
                || string.Equals(filter, "xsf", StringComparison.OrdinalIgnoreCase)
                || string.Equals(filter, "xsf-", StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            return id.StartsWith(filter, StringComparison.OrdinalIgnoreCase);
        }

        private static string SeedMain(XlSession session, IOperationChannel channel, string sheet)
        {
            bool ok = session.SeedCell(channel, sheet, "A1", "编号")
                && session.SeedCell(channel, sheet, "B1", "姓名")
                && session.SeedCell(channel, sheet, "C1", "地区")
                && session.SeedCell(channel, sheet, "D1", "金额")
                && session.SeedCell(channel, sheet, "A2", 1)
                && session.SeedCell(channel, sheet, "B2", "甲")
                && session.SeedCell(channel, sheet, "C2", "华东")
                && session.SeedCell(channel, sheet, "D2", 100)
                && session.SeedCell(channel, sheet, "A3", 2)
                && session.SeedCell(channel, sheet, "B3", "乙")
                && session.SeedCell(channel, sheet, "C3", "华北")
                && session.SeedCell(channel, sheet, "D3", 3000)
                && session.SeedCell(channel, sheet, "A4", 3)
                && session.SeedCell(channel, sheet, "B4", "丙")
                && session.SeedCell(channel, sheet, "C4", "华东")
                && session.SeedCell(channel, sheet, "D4", 50)
                && session.SeedCell(channel, sheet, "A5", 4)
                && session.SeedCell(channel, sheet, "B5", "丁")
                && session.SeedCell(channel, sheet, "C5", "华南")
                && session.SeedCell(channel, sheet, "D5", 2000);
            return ok ? null : "铺明细失败";
        }

        private static string SeedPair(XlSession session, IOperationChannel channel, string sheet)
        {
            bool ok = session.SeedCell(channel, sheet, "A1", "编号")
                && session.SeedCell(channel, sheet, "B1", "姓名")
                && session.SeedCell(channel, sheet, "A2", 1)
                && session.SeedCell(channel, sheet, "B2", "甲")
                && session.SeedCell(channel, sheet, "A3", 2)
                && session.SeedCell(channel, sheet, "B3", "乙")
                && session.SeedCell(channel, sheet, "A4", 3)
                && session.SeedCell(channel, sheet, "B4", "丙")
                && session.SeedCell(channel, sheet, "A5", 4)
                && session.SeedCell(channel, sheet, "B5", "丁")
                && session.SeedCell(channel, sheet, "D1", "地区")
                && session.SeedCell(channel, sheet, "E1", "金额")
                && session.SeedCell(channel, sheet, "D2", "华东")
                && session.SeedCell(channel, sheet, "E2", 100)
                && session.SeedCell(channel, sheet, "D3", "华北")
                && session.SeedCell(channel, sheet, "E3", 3000)
                && session.SeedCell(channel, sheet, "D4", "华东")
                && session.SeedCell(channel, sheet, "E4", 50)
                && session.SeedCell(channel, sheet, "D5", "华南")
                && session.SeedCell(channel, sheet, "E5", 2000);
            return ok ? null : "铺两块失败";
        }

        private static string AddTable(IOperationChannel channel, string sheet, string a1)
        {
            try
            {
                if (channel is ExcelChannel excel && excel.TryGetLiveWorkbook(out Excel.Workbook book))
                {
                    Excel.Worksheet worksheet = FindExcelSheet(book, sheet);
                    Excel.ListObject table = worksheet.ListObjects.Add(
                        Excel.XlListObjectSourceType.xlSrcRange,
                        worksheet.Range[a1],
                        Type.Missing,
                        Excel.XlYesNoGuess.xlYes);
                    table.ShowAutoFilter = true;
                    return null;
                }

                if (channel is EtChannel et && et.TryGetLiveWorkbook(out object etBook))
                {
                    object worksheet = FindEtSheet(etBook, sheet);
                    object range = EtCom.InvokeFlex(worksheet, "Range", a1);
                    object tables = EtCom.GetProperty(worksheet, "ListObjects");
                    object table = EtCom.InvokeFlex(tables, "Add", 1, range, Type.Missing, 1);
                    if (table == null)
                    {
                        return "插入表失败";
                    }

                    EtCom.TrySetProperty(table, "ShowAutoFilter", true);
                    return null;
                }
            }
            catch (Exception ex)
            {
                return "插入表失败: " + ex.Message;
            }

            return "没有工作簿";
        }

        private static bool RowHidden(IOperationChannel channel, string sheet, int row)
        {
            if (channel is ExcelChannel excel && excel.TryGetLiveWorkbook(out Excel.Workbook book))
            {
                Excel.Worksheet worksheet = FindExcelSheet(book, sheet);
                Excel.Range line = (Excel.Range)worksheet.Rows[row];
                return line != null && line.Hidden;
            }

            if (channel is EtChannel et && et.TryGetLiveWorkbook(out object etBook))
            {
                object worksheet = FindEtSheet(etBook, sheet);
                object rows = EtCom.GetProperty(worksheet, "Rows");
                object line = EtCom.GetIndexed(rows, row);
                object hidden = line == null ? null : EtCom.GetProperty(line, "Hidden");
                return hidden != null && Convert.ToBoolean(hidden);
            }

            return false;
        }

        private static Excel.Worksheet FindExcelSheet(Excel.Workbook book, string sheet)
        {
            foreach (Excel.Worksheet worksheet in book.Worksheets)
            {
                if (string.Equals(worksheet.Name, sheet, StringComparison.OrdinalIgnoreCase))
                {
                    return worksheet;
                }
            }

            return null;
        }

        private static object FindEtSheet(object book, string sheet)
        {
            foreach (object worksheet in EtCom.EnumerateSheets(book))
            {
                string name = Convert.ToString(EtCom.GetProperty(worksheet, "Name")) ?? "";
                if (string.Equals(name, sheet, StringComparison.OrdinalIgnoreCase))
                {
                    return worksheet;
                }
            }

            return null;
        }

        private static ToolResult Call(
            XlSession session,
            string channelId,
            string sheet,
            string action,
            string range = null,
            int? index = null,
            object sorts = null,
            object filters = null)
        {
            var args = new Dictionary<string, object>
            {
                ["channel_id"] = channelId,
                ["sheet"] = sheet,
                ["action"] = action
            };
            if (range != null)
            {
                args["range"] = range;
            }

            if (index.HasValue)
            {
                args["index"] = index.Value;
            }

            if (sorts != null)
            {
                args["sorts"] = sorts;
            }

            if (filters != null)
            {
                args["filters"] = filters;
            }

            return session.Call("F_excel_sort_filter", args);
        }

        private static Dictionary<string, object> FilterIn(string column, params string[] values)
        {
            return new Dictionary<string, object>
            {
                ["column"] = column,
                ["in"] = new List<string>(values)
            };
        }

        private static Dictionary<string, object> FilterExpr(string column, string expr)
        {
            return new Dictionary<string, object>
            {
                ["column"] = column,
                ["expr"] = expr
            };
        }

        private static Dictionary<string, object> SortKey(string column, string order)
        {
            return new Dictionary<string, object>
            {
                ["column"] = column,
                ["order"] = order
            };
        }

        private static int PanelCount(ToolResult result)
        {
            Dictionary<string, object> data = Data(result);
            return CountOf(data != null && data.ContainsKey("panels") ? data["panels"] : null);
        }

        private static Dictionary<string, object> PanelAt(ToolResult result, int index)
        {
            Dictionary<string, object> data = Data(result);
            object panels = data != null && data.ContainsKey("panels") ? data["panels"] : null;
            return ItemAt(panels, index);
        }

        private static object PanelField(ToolResult result, string key)
        {
            Dictionary<string, object> data = Data(result);
            if (data != null && data.ContainsKey(key))
            {
                return data[key];
            }

            return null;
        }

        private static string FilterText(ToolResult result, string key)
        {
            object filters = PanelField(result, "filters");
            Dictionary<string, object> first = ItemAt(filters, 0);
            if (first == null || !first.ContainsKey(key) || first[key] == null)
            {
                return null;
            }

            if (first[key] is string text)
            {
                return text;
            }

            var parts = new List<string>();
            if (first[key] is IEnumerable enumerable && !(first[key] is string))
            {
                foreach (object item in enumerable)
                {
                    parts.Add(Convert.ToString(item, CultureInfo.InvariantCulture));
                }
            }

            return string.Join(",", parts);
        }

        private static Dictionary<string, object> ItemAt(object raw, int index)
        {
            if (!(raw is IEnumerable enumerable) || raw is string)
            {
                return null;
            }

            int at = 0;
            foreach (object item in enumerable)
            {
                if (at == index && item is Dictionary<string, object> dict)
                {
                    return dict;
                }

                at++;
            }

            return null;
        }

        private static int CountOf(object raw)
        {
            if (raw is ICollection collection)
            {
                return collection.Count;
            }

            if (!(raw is IEnumerable enumerable) || raw is string)
            {
                return 0;
            }

            int count = 0;
            foreach (object unused in enumerable)
            {
                count++;
            }

            return count;
        }

        private static int IntOf(ToolResult result, string key)
        {
            Dictionary<string, object> data = Data(result);
            if (data == null || !data.ContainsKey(key) || data[key] == null)
            {
                return -1;
            }

            return Convert.ToInt32(data[key], CultureInfo.InvariantCulture);
        }

        private static Dictionary<string, object> Data(ToolResult result)
        {
            return result == null ? null : result.Data as Dictionary<string, object>;
        }

        private static string Str(ToolResult result, string key)
        {
            Dictionary<string, object> data = Data(result);
            if (data == null || !data.ContainsKey(key) || data[key] == null)
            {
                return "";
            }

            return Convert.ToString(data[key], CultureInfo.InvariantCulture) ?? "";
        }

        private static string Text(object value)
        {
            return value == null ? "" : (Convert.ToString(value, CultureInfo.InvariantCulture) ?? "");
        }

        private static bool IsNum(object value, double expected)
        {
            if (value == null)
            {
                return false;
            }

            double number;
            try
            {
                number = Convert.ToDouble(value, CultureInfo.InvariantCulture);
            }
            catch (Exception)
            {
                return false;
            }

            return Math.Abs(number - expected) < 0.001;
        }

        private static bool Has(string text, string part)
        {
            return !string.IsNullOrEmpty(text)
                && text.IndexOf(part, StringComparison.OrdinalIgnoreCase) >= 0;
        }
    }
}
