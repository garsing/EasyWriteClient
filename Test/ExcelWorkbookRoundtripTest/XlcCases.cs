using System;
using System.Collections;
using System.Collections.Generic;
using WordAddIn1;
using WordAddIn1.SpreadsheetHost;

namespace ExcelWorkbookRoundtripTest
{
    internal static class XlcCases
    {
        public static void Run(TestRun run, bool et)
        {
            var s = new XlSession(et);
            RunNeg(run, s);
            RunApply(run, s);
            RunIntersect(run, s);
            RunSolid(run, s);
        }

        private static void RunNeg(TestRun run, XlSession s)
        {
            ChannelRegistry.ClearAll();
            ToolResult noCh = s.CondFormat(null, "利润表", "A1", "add", LessRule());
            ExpectFail(run, "xlc-neg-no-channel", noCh, "打开");

            var req = new SpreadsheetConditionalFormatRequest
            {
                SheetName = "利润表",
                RangeA1 = "A1",
                Action = "add",
                Rule = new SpreadsheetConditionalRule
                {
                    Type = "cell_value",
                    Operator = "less",
                    Formula1 = "0",
                    Format = new SpreadsheetConditionalAppearance { FillColorHex = "FFCCCC" }
                }
            };
            if (!SpreadsheetHostAdapter.TryApplyConditionalFormat(
                    new FakeWordChannel(), req, out _, out ToolResult wrong))
            {
                ExpectFail(run, "xlc-neg-wrong-kind", wrong, "excel");
            }
            else
            {
                run.CaseFail("xlc-neg-wrong-kind", "word 渠道应失败");
            }

            WithBook(run, s, "xlc-neg-no-sheet", (id, ch) =>
            {
                ToolResult r = s.Call("F_apply_excel_conditional_format", new Dictionary<string, object>
                {
                    ["channel_id"] = id,
                    ["range"] = "A1",
                    ["action"] = "add",
                    ["rule"] = LessRule()
                });
                return r.Success ? "缺 sheet 应失败" : (Has(r.Error, "sheet") ? null : r.Error);
            });

            WithBook(run, s, "xlc-neg-no-range", (id, ch) =>
            {
                ToolResult r = s.Call("F_apply_excel_conditional_format", new Dictionary<string, object>
                {
                    ["channel_id"] = id,
                    ["sheet"] = s.FirstSheetName(ch),
                    ["action"] = "list"
                });
                return r.Success ? "缺 range 应失败" : (Has(r.Error, "range") ? null : r.Error);
            });

            WithBook(run, s, "xlc-neg-no-action", (id, ch) =>
            {
                ToolResult r = s.Call("F_apply_excel_conditional_format", new Dictionary<string, object>
                {
                    ["channel_id"] = id,
                    ["sheet"] = s.FirstSheetName(ch),
                    ["range"] = "A1"
                });
                return r.Success ? "缺 action 应失败" : (Has(r.Error, "action") ? null : r.Error);
            });

            WithBook(run, s, "xlc-neg-bad-action", (id, ch) =>
            {
                ToolResult r = s.CondFormat(id, s.FirstSheetName(ch), "A1", "apply", LessRule());
                return r.Success ? "非法 action 应失败" : (Has(r.Error, "action") ? null : r.Error);
            });

            WithBook(run, s, "xlc-neg-add-no-rule", (id, ch) =>
            {
                ToolResult r = s.CondFormat(id, s.FirstSheetName(ch), "A1", "add");
                return r.Success ? "add 无 rule 应失败" : (Has(r.Error, "rule") ? null : r.Error);
            });

            WithBook(run, s, "xlc-neg-readonly-rule", (id, ch) =>
            {
                string name = s.FirstSheetName(ch);
                ToolResult list = s.CondFormat(id, name, "A1", "list", LessRule());
                if (list.Success || !Has(list.Error, "rule"))
                {
                    return list.Success ? "list 带 rule 应失败" : list.Error;
                }

                ToolResult clear = s.CondFormat(id, name, "A1", "clear", LessRule());
                return clear.Success
                    ? "clear 带 rule 应失败"
                    : (Has(clear.Error, "rule") ? null : clear.Error);
            });

            WithBook(run, s, "xlc-neg-bang", (id, ch) =>
            {
                string name = s.FirstSheetName(ch);
                ToolResult r = s.CondFormat(id, name, name + "!C2:C20", "list");
                return r.Success ? "含!应失败" : (Has(r.Error, "纯 A1") || Has(r.Error, "!") ? null : r.Error);
            });

            WithBook(run, s, "xlc-neg-missing-sheet", (id, ch) =>
            {
                ToolResult r = s.CondFormat(id, "不存在的表", "A1", "list");
                return r.Success
                    ? "缺表应失败"
                    : (Has(r.Error, "不存在") || Has(r.Error, "未找到") ? null : r.Error);
            });
        }

        private static void RunApply(TestRun run, XlSession s)
        {
            WithBook(run, s, "xlc-cell-less", (id, ch) =>
            {
                string name = s.FirstSheetName(ch);
                if (!s.SeedCell(ch, name, "C2", -1) || !s.SeedCell(ch, name, "C3", 1))
                {
                    return "铺底失败";
                }

                ToolResult r = s.CondFormat(id, name, "C2:C3", "replace_all", LessRule());
                if (!r.Success || Int(r, "rules_count") != 1)
                {
                    return r.Success ? "rules_count=" + Int(r, "rules_count") : r.Error;
                }

                ToolResult g = s.CondFormat(id, name, "C2:C3", "list");
                if (!g.Success || !HasType(g, "cell_value") || !HasOp(g, "less")
                    || !HasFormulaField(g, "formula1", "0"))
                {
                    return g.Success ? "list 对不上 less/0" : g.Error;
                }

                return SameNumber(s, ch, name, "C2", -1) && SameNumber(s, ch, name, "C3", 1)
                    ? null
                    : "值被改了";
            });

            WithBook(run, s, "xlc-cell-between", (id, ch) =>
            {
                string name = s.FirstSheetName(ch);
                if (!s.SeedCell(ch, name, "B2", 5))
                {
                    return "铺底失败";
                }

                ToolResult r = s.CondFormat(id, name, "B2", "replace_all", CellValue(
                    "between", "1", "10"));
                if (!r.Success)
                {
                    return r.Error;
                }

                ToolResult g = s.CondFormat(id, name, "B2", "list");
                return g.Success && HasOp(g, "between")
                    && HasFormulaField(g, "formula1", "1")
                    && HasFormulaField(g, "formula2", "10")
                    ? null
                    : (g.Error ?? "between 回读不对");
            });

            WithBook(run, s, "xlc-cell-equal", (id, ch) =>
            {
                string name = s.FirstSheetName(ch);
                if (!s.SeedCell(ch, name, "A1", "完成"))
                {
                    return "铺底失败";
                }

                ToolResult r = s.CondFormat(id, name, "A1", "replace_all", CellValue(
                    "equal", "\"完成\""));
                if (!r.Success)
                {
                    return r.Error;
                }

                ToolResult g = s.CondFormat(id, name, "A1", "list");
                if (!g.Success || !HasOp(g, "equal"))
                {
                    return g.Error ?? "operator 不是 equal";
                }

                ToolResult rd = s.Read(id, name, "A1");
                return rd.Success && Has(Str(rd, "table"), "完成")
                    ? null
                    : (rd.Error ?? "值丢了");
            });

            WithBook(run, s, "xlc-formula", (id, ch) =>
            {
                string name = s.FirstSheetName(ch);
                if (!s.SeedCell(ch, name, "C2", -1) || !s.SeedCell(ch, name, "C3", 1))
                {
                    return "铺底失败";
                }

                ToolResult r = s.CondFormat(id, name, "A2:D3", "add", FormulaRule("=$C2<0"));
                if (!r.Success)
                {
                    return r.Error;
                }

                ToolResult g = s.CondFormat(id, name, "A2:D3", "list");
                if (!g.Success || !HasType(g, "formula") || !HasFormulaField(g, "formula", "C2"))
                {
                    return g.Success ? "formula 回读无 C2" : g.Error;
                }

                return SameNumber(s, ch, name, "C2", -1) && SameNumber(s, ch, name, "C3", 1)
                    ? null
                    : "值被改了";
            });

            WithBook(run, s, "xlc-add-stack", (id, ch) =>
            {
                string name = s.FirstSheetName(ch);
                ToolResult a = s.CondFormat(id, name, "A1", "add", LessRule());
                if (!a.Success)
                {
                    return a.Error;
                }

                ToolResult b = s.CondFormat(id, name, "A1", "add", CellValue("greater", "10"));
                if (!b.Success || Int(b, "rules_count") != 2)
                {
                    return b.Success ? "rules_count=" + Int(b, "rules_count") : b.Error;
                }

                ToolResult g = s.CondFormat(id, name, "A1", "list");
                return g.Success && HasOp(g, "less") && HasOp(g, "greater")
                    ? null
                    : (g.Error ?? "两条没叠上");
            });

            WithBook(run, s, "xlc-replace-all", (id, ch) =>
            {
                string name = s.FirstSheetName(ch);
                ToolResult a = s.CondFormat(id, name, "A1", "add", LessRule());
                ToolResult b = s.CondFormat(id, name, "A1", "add", CellValue("greater", "10"));
                if (!a.Success || !b.Success)
                {
                    return a.Error ?? b.Error;
                }

                ToolResult r = s.CondFormat(id, name, "A1", "replace_all", CellValue("equal", "0"));
                if (!r.Success || Int(r, "rules_count") != 1 || Int(r, "cleared_count") < 1)
                {
                    return r.Success
                        ? "count=" + Int(r, "rules_count") + " cleared=" + Int(r, "cleared_count")
                        : r.Error;
                }

                ToolResult g = s.CondFormat(id, name, "A1", "list");
                return g.Success && Int(g, "rules_count") == 1 && HasOp(g, "equal") && !HasOp(g, "less")
                    ? null
                    : (g.Error ?? "replace_all 后不是只剩 equal");
            });

            WithBook(run, s, "xlc-appearance", (id, ch) =>
            {
                string name = s.FirstSheetName(ch);
                ToolResult r = s.CondFormat(id, name, "A1", "replace_all", new Dictionary<string, object>
                {
                    ["type"] = "cell_value",
                    ["operator"] = "less",
                    ["formula1"] = "0",
                    ["format"] = new Dictionary<string, object>
                    {
                        ["fill_color"] = "#FFCCCC",
                        ["font_color"] = "#990000",
                        ["bold"] = true
                    }
                });
                if (!r.Success)
                {
                    return r.Error;
                }

                ToolResult g = s.CondFormat(id, name, "A1", "list");
                Dictionary<string, object> fmt = RuleFmt(FirstRule(g));
                if (!HexEq(Str(fmt, "fill_color"), "FFCCCC") || !IsTrue(fmt, "bold"))
                {
                    return g.Success
                        ? "外观 " + Str(fmt, "fill_color") + " bold=" + Str(fmt, "bold")
                        : g.Error;
                }

                string font = Str(fmt, "font_color");
                return string.IsNullOrEmpty(font) || HexEq(font, "990000")
                    ? null
                    : "字色 " + font;
            });

            WithBook(run, s, "xlc-static-kept", (id, ch) =>
            {
                string name = s.FirstSheetName(ch);
                if (!s.SeedCell(ch, name, "A1", "甲"))
                {
                    return "铺底失败";
                }

                ToolResult ap = s.ApplyFormat(id, name, "A1", new Dictionary<string, object>
                {
                    ["bold"] = true
                });
                if (!ap.Success)
                {
                    return ap.Error;
                }

                ToolResult r = s.CondFormat(id, name, "A1", "replace_all", LessRule());
                if (!r.Success)
                {
                    return r.Error;
                }

                ToolResult g = s.GetFormat(id, name, "A1");
                Dictionary<string, object> fmt = ItemFmt(g);
                return g.Success && IsTrue(fmt, "bold")
                    ? null
                    : (g.Error ?? "静态 bold 丢了");
            });

            WithBook(run, s, "xlc-hidden", (id, ch) =>
            {
                if (!s.SeedHidden(ch, "可见", "隐藏底稿"))
                {
                    return "铺隐藏表失败";
                }

                ToolResult r = s.CondFormat(id, "隐藏底稿", "A1", "replace_all", LessRule());
                if (!r.Success)
                {
                    return r.Error;
                }

                ToolResult g = s.CondFormat(id, "隐藏底稿", "A1", "list");
                return g.Success && Int(g, "rules_count") >= 1
                    ? null
                    : (g.Error ?? "隐藏表没挂上");
            });

            WithBook(run, s, "xlc-clear", (id, ch) =>
            {
                string name = s.FirstSheetName(ch);
                ToolResult add = s.CondFormat(id, name, "A1", "replace_all", LessRule());
                if (!add.Success)
                {
                    return add.Error;
                }

                ToolResult r = s.CondFormat(id, name, "A1", "clear");
                if (!r.Success || Int(r, "cleared_count") < 1 || Int(r, "rules_count") != 0)
                {
                    return r.Success
                        ? "cleared=" + Int(r, "cleared_count") + " left=" + Int(r, "rules_count")
                        : r.Error;
                }

                ToolResult g = s.CondFormat(id, name, "A1", "list");
                return g.Success && Int(g, "rules_count") == 0
                    ? null
                    : (g.Error ?? "clear 后 list 非空");
            });
        }

        private static void RunIntersect(TestRun run, XlSession s)
        {
            WithBook(run, s, "xlc-list-applies", (id, ch) =>
            {
                string name = s.FirstSheetName(ch);
                ToolResult r = s.CondFormat(id, name, "C2:C5", "replace_all", LessRule());
                if (!r.Success)
                {
                    return r.Error;
                }

                ToolResult g = s.CondFormat(id, name, "C2:C5", "list");
                string applies = Str(FirstRule(g), "applies_to");
                return g.Success && string.Equals(applies, "C2:C5", StringComparison.OrdinalIgnoreCase)
                    && applies.IndexOf('$') < 0 && applies.IndexOf('!') < 0
                    ? null
                    : (g.Error ?? ("applies_to=" + applies));
            });

            WithBook(run, s, "xlc-clear-intersect", (id, ch) =>
            {
                string name = s.FirstSheetName(ch);
                ToolResult add = s.CondFormat(id, name, "A1:C10", "replace_all", LessRule());
                if (!add.Success)
                {
                    return add.Error;
                }

                ToolResult r = s.CondFormat(id, name, "B1:B5", "clear");
                if (!r.Success || Int(r, "cleared_count") < 1)
                {
                    return r.Success ? "cleared=" + Int(r, "cleared_count") : r.Error;
                }

                ToolResult g = s.CondFormat(id, name, "A1:C10", "list");
                return g.Success && Int(g, "rules_count") == 0
                    ? null
                    : (g.Error ?? "相交未整条删掉");
            });

            WithBook(run, s, "xlc-list-disjoint", (id, ch) =>
            {
                string name = s.FirstSheetName(ch);
                ToolResult add = s.CondFormat(id, name, "A1:A10", "replace_all", LessRule());
                if (!add.Success)
                {
                    return add.Error;
                }

                ToolResult miss = s.CondFormat(id, name, "B1:B5", "list");
                if (!miss.Success || Int(miss, "rules_count") != 0)
                {
                    return miss.Success ? "不相交不应看到" : miss.Error;
                }

                ToolResult hit = s.CondFormat(id, name, "A5", "list");
                return hit.Success && Int(hit, "rules_count") >= 1
                    ? null
                    : (hit.Error ?? "A5 应看到规则");
            });

            WithBook(run, s, "xlc-neg-chart", (id, ch) =>
            {
                if (!s.SeedChartSheet(ch, out string chartName))
                {
                    return "铺图表 sheet 失败";
                }

                ToolResult r = s.CondFormat(id, chartName, "A1", "list");
                return r.Success ? "图表 sheet 应失败" : (Has(r.Error, "sheet_type=chart") ? null : r.Error);
            });

            WithBook(run, s, "xlc-neg-oversize", (id, ch) =>
            {
                ToolResult r = s.CondFormat(id, s.FirstSheetName(ch), "A1:A501", "add", LessRule());
                return r.Success ? "超限应失败" : (Has(r.Error, "过大") || Has(r.Error, "500") ? null : r.Error);
            });

            WithBook(run, s, "xlc-neg-static-key", (id, ch) =>
            {
                string name = s.FirstSheetName(ch);
                ToolResult r = s.CondFormat(id, name, "A1", "add", new Dictionary<string, object>
                {
                    ["type"] = "cell_value",
                    ["operator"] = "less",
                    ["formula1"] = "0",
                    ["format"] = new Dictionary<string, object>
                    {
                        ["fill_color"] = "#FFCCCC",
                        ["number_format"] = "0.00"
                    }
                });
                if (r.Success || !Has(r.Error, "F_apply_excel_format"))
                {
                    return r.Success ? "静态键应失败" : r.Error;
                }

                ToolResult g = s.CondFormat(id, name, "A1", "list");
                return g.Success && Int(g, "rules_count") == 0
                    ? null
                    : (g.Error ?? "半截挂上了");
            });

            WithBook(run, s, "xlc-neg-bad-type", (id, ch) =>
            {
                ToolResult r = s.CondFormat(id, s.FirstSheetName(ch), "A1", "add", new Dictionary<string, object>
                {
                    ["type"] = "color_scale",
                    ["format"] = FillOnly()
                });
                return r.Success ? "color_scale 应失败" : (Has(r.Error, "cell_value") || Has(r.Error, "type")
                    ? null
                    : r.Error);
            });

            WithBook(run, s, "xlc-neg-empty-format", (id, ch) =>
            {
                ToolResult r = s.CondFormat(id, s.FirstSheetName(ch), "A1", "add", new Dictionary<string, object>
                {
                    ["type"] = "cell_value",
                    ["operator"] = "less",
                    ["formula1"] = "0",
                    ["format"] = new Dictionary<string, object>()
                });
                return r.Success ? "空 format 应失败" : (Has(r.Error, "至少一项") || Has(r.Error, "format")
                    ? null
                    : r.Error);
            });

            WithBook(run, s, "xlc-neg-between-no-f2", (id, ch) =>
            {
                ToolResult r = s.CondFormat(id, s.FirstSheetName(ch), "A1", "add", new Dictionary<string, object>
                {
                    ["type"] = "cell_value",
                    ["operator"] = "between",
                    ["formula1"] = "1",
                    ["format"] = FillOnly()
                });
                return r.Success ? "缺 formula2 应失败" : (Has(r.Error, "formula2") ? null : r.Error);
            });

            WithBook(run, s, "xlc-zh-sheet", (id, ch) =>
            {
                if (!s.SeedUsed(ch, "利润表"))
                {
                    return "改名失败";
                }

                ToolResult r = s.CondFormat(id, "利润表", "B2", "replace_all", LessRule());
                if (!r.Success || !string.Equals(Str(r, "sheet"), "利润表", StringComparison.Ordinal))
                {
                    return r.Success ? "replace sheet=" + Str(r, "sheet") : r.Error;
                }

                ToolResult g = s.CondFormat(id, "利润表", "B2", "list");
                return g.Success && string.Equals(Str(g, "sheet"), "利润表", StringComparison.Ordinal)
                    ? null
                    : (g.Success ? "list sheet=" + Str(g, "sheet") : g.Error);
            });
        }

        private static void RunSolid(TestRun run, XlSession s)
        {
            WithBook(run, s, "xlc-list-less-span", (id, ch) =>
            {
                string name = s.FirstSheetName(ch);
                ToolResult r = s.CondFormat(id, name, "C2:C10", "replace_all", LessRule());
                if (!r.Success || Int(r, "rules_count") != 1)
                {
                    return r.Success ? "rules_count=" + Int(r, "rules_count") : r.Error;
                }

                ToolResult g = s.CondFormat(id, name, "C2:C10", "list");
                return g.Success && Int(g, "rules_count") == 1 && HasOp(g, "less")
                    && HasFormulaField(g, "formula1", "0")
                    ? null
                    : (g.Error ?? "less 多格被 Formula2 丢掉");
            });

            WithBook(run, s, "xlc-list-not-equal", (id, ch) =>
            {
                string name = s.FirstSheetName(ch);
                if (!s.SeedCell(ch, name, "A1", "甲"))
                {
                    return "铺底失败";
                }

                ToolResult r = s.CondFormat(id, name, "A1", "replace_all", CellValue("not_equal", "\"乙\""));
                if (!r.Success)
                {
                    return r.Error;
                }

                ToolResult g = s.CondFormat(id, name, "A1", "list");
                return g.Success && HasOp(g, "not_equal")
                    ? null
                    : (g.Error ?? "not_equal 回读丢掉");
            });

            WithBook(run, s, "xlc-geq", (id, ch) =>
            {
                string name = s.FirstSheetName(ch);
                ToolResult r = s.CondFormat(id, name, "B2:B8", "replace_all", CellValue("greater_or_equal", "10"));
                if (!r.Success || Int(r, "rules_count") != 1)
                {
                    return r.Success ? "rules_count=" + Int(r, "rules_count") : r.Error;
                }

                ToolResult g = s.CondFormat(id, name, "B2:B8", "list");
                return g.Success && HasOp(g, "greater_or_equal") && HasFormulaField(g, "formula1", "10")
                    ? null
                    : (g.Error ?? "greater_or_equal 回读丢掉");
            });

            WithBook(run, s, "xlc-leq", (id, ch) =>
            {
                string name = s.FirstSheetName(ch);
                ToolResult r = s.CondFormat(id, name, "B2:B8", "replace_all", CellValue("less_or_equal", "0"));
                if (!r.Success)
                {
                    return r.Error;
                }

                ToolResult g = s.CondFormat(id, name, "B2:B8", "list");
                return g.Success && HasOp(g, "less_or_equal")
                    ? null
                    : (g.Error ?? "less_or_equal 回读丢掉");
            });

            WithBook(run, s, "xlc-not-between", (id, ch) =>
            {
                string name = s.FirstSheetName(ch);
                if (!s.SeedCell(ch, name, "B2", 5))
                {
                    return "铺底失败";
                }

                ToolResult r = s.CondFormat(id, name, "B2", "replace_all", CellValue(
                    "not_between", "1", "10"));
                if (!r.Success)
                {
                    return r.Error;
                }

                ToolResult g = s.CondFormat(id, name, "B2", "list");
                return g.Success && HasOp(g, "not_between")
                    && HasFormulaField(g, "formula1", "1")
                    && HasFormulaField(g, "formula2", "10")
                    ? null
                    : (g.Error ?? "not_between 回读不对");
            });

            WithBook(run, s, "xlc-formula-no-shift", (id, ch) =>
            {
                string name = s.FirstSheetName(ch);
                if (!s.SeedCell(ch, name, "C2", -1) || !s.SeedCell(ch, name, "C3", 1)
                    || !s.SeedCell(ch, name, "C4", 1))
                {
                    return "铺底失败";
                }

                ToolResult r = s.CondFormat(id, name, "A2:D5", "add", FormulaRule("=$C2<0"));
                if (!r.Success)
                {
                    return r.Error;
                }

                ToolResult g = s.CondFormat(id, name, "A2:D5", "list");
                string formula = Str(FirstRule(g), "formula");
                if (!g.Success || !HasType(g, "formula") || !Has(formula, "C2") || Has(formula, "C3"))
                {
                    return g.Success ? "公式错位 " + formula : g.Error;
                }

                return SameNumber(s, ch, name, "C2", -1) && SameNumber(s, ch, name, "C3", 1)
                    ? null
                    : "值被改了";
            });

            WithBook(run, s, "xlc-replace-disjoint", (id, ch) =>
            {
                string name = s.FirstSheetName(ch);
                ToolResult a = s.CondFormat(id, name, "A1:A10", "add", LessRule());
                if (!a.Success)
                {
                    return a.Error;
                }

                ToolResult r = s.CondFormat(id, name, "C1:C10", "replace_all", CellValue("greater", "10"));
                if (!r.Success)
                {
                    return r.Error;
                }

                ToolResult colA = s.CondFormat(id, name, "A1:A10", "list");
                ToolResult colC = s.CondFormat(id, name, "C1:C10", "list");
                if (!colA.Success || Int(colA, "rules_count") != 1 || !HasOp(colA, "less"))
                {
                    return colA.Error ?? "A 列规则被 replace 误伤";
                }

                return colC.Success && HasOp(colC, "greater") && !HasOp(colC, "less")
                    ? null
                    : (colC.Error ?? "C 列不是 greater");
            });

            WithBook(run, s, "xlc-clear-span-less", (id, ch) =>
            {
                string name = s.FirstSheetName(ch);
                ToolResult add = s.CondFormat(id, name, "C2:C20", "replace_all", LessRule());
                if (!add.Success || Int(add, "rules_count") != 1)
                {
                    return add.Success ? "挂上后 count=" + Int(add, "rules_count") : add.Error;
                }

                ToolResult r = s.CondFormat(id, name, "C2:C20", "clear");
                if (!r.Success || Int(r, "cleared_count") < 1)
                {
                    return r.Success ? "cleared=" + Int(r, "cleared_count") : r.Error;
                }

                ToolResult g = s.CondFormat(id, name, "C2:C20", "list");
                return g.Success && Int(g, "rules_count") == 0
                    ? null
                    : (g.Error ?? "clear 多格 less 后仍看得到");
            });
        }

        private static Dictionary<string, object> FillOnly()
        {
            return new Dictionary<string, object> { ["fill_color"] = "#FFCCCC" };
        }

        private static Dictionary<string, object> LessRule()
        {
            return CellValue("less", "0");
        }

        private static Dictionary<string, object> CellValue(string op, string formula1, string formula2 = null)
        {
            var d = new Dictionary<string, object>
            {
                ["type"] = "cell_value",
                ["operator"] = op,
                ["formula1"] = formula1,
                ["format"] = FillOnly()
            };
            if (formula2 != null)
            {
                d["formula2"] = formula2;
            }

            return d;
        }

        private static Dictionary<string, object> FormulaRule(string formula)
        {
            return new Dictionary<string, object>
            {
                ["type"] = "formula",
                ["formula"] = formula,
                ["format"] = FillOnly()
            };
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
            return Str(Data(r), key);
        }

        private static string Str(Dictionary<string, object> d, string key)
        {
            if (d == null || !d.ContainsKey(key) || d[key] == null)
            {
                return "";
            }

            return Convert.ToString(d[key]) ?? "";
        }

        private static int Int(ToolResult r, string key)
        {
            Dictionary<string, object> d = Data(r);
            if (d == null || !d.ContainsKey(key) || d[key] == null)
            {
                return 0;
            }

            try
            {
                return Convert.ToInt32(d[key]);
            }
            catch
            {
                return 0;
            }
        }

        private static List<Dictionary<string, object>> Rules(ToolResult r)
        {
            var list = new List<Dictionary<string, object>>();
            Dictionary<string, object> data = Data(r);
            if (data == null || !data.ContainsKey("rules") || !(data["rules"] is IEnumerable raw))
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

        private static Dictionary<string, object> FirstRule(ToolResult r)
        {
            List<Dictionary<string, object>> rules = Rules(r);
            return rules.Count > 0 ? rules[0] : null;
        }

        private static bool HasType(ToolResult r, string type)
        {
            foreach (Dictionary<string, object> one in Rules(r))
            {
                if (string.Equals(Str(one, "type"), type, StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
            }

            return false;
        }

        private static bool HasOp(ToolResult r, string op)
        {
            foreach (Dictionary<string, object> one in Rules(r))
            {
                if (string.Equals(Str(one, "operator"), op, StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
            }

            return false;
        }

        private static bool HasFormulaField(ToolResult r, string key, string part)
        {
            foreach (Dictionary<string, object> one in Rules(r))
            {
                if (Has(Str(one, key), part))
                {
                    return true;
                }
            }

            return false;
        }

        private static Dictionary<string, object> RuleFmt(Dictionary<string, object> rule)
        {
            if (rule == null || !rule.ContainsKey("format"))
            {
                return new Dictionary<string, object>();
            }

            return rule["format"] as Dictionary<string, object> ?? new Dictionary<string, object>();
        }

        private static Dictionary<string, object> ItemFmt(ToolResult r)
        {
            Dictionary<string, object> data = Data(r);
            if (data == null || !data.ContainsKey("items") || !(data["items"] is IEnumerable raw))
            {
                return new Dictionary<string, object>();
            }

            foreach (object one in raw)
            {
                if (one is Dictionary<string, object> item && item.ContainsKey("format"))
                {
                    return item["format"] as Dictionary<string, object> ?? new Dictionary<string, object>();
                }
            }

            return new Dictionary<string, object>();
        }

        private static bool IsTrue(Dictionary<string, object> d, string key)
        {
            if (d == null || !d.ContainsKey(key) || d[key] == null)
            {
                return false;
            }

            try
            {
                return Convert.ToBoolean(d[key]);
            }
            catch
            {
                return false;
            }
        }

        private static bool HexEq(string raw, string rgb)
        {
            string a = (raw ?? "").Replace("#", "").Trim();
            return string.Equals(a, rgb, StringComparison.OrdinalIgnoreCase);
        }

        private static bool SameNumber(XlSession s, IOperationChannel ch, string sheet, string a1, double expect)
        {
            object raw = s.ReadValue2(ch, sheet, a1);
            try
            {
                return Math.Abs(Convert.ToDouble(raw) - expect) < 0.0001;
            }
            catch
            {
                return false;
            }
        }
    }
}
