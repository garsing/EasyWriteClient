using System;
using System.Collections;
using System.Collections.Generic;
using WordAddIn1;
using WordAddIn1.SpreadsheetHost;

namespace ExcelWorkbookRoundtripTest
{
    internal static class XlpCases
    {
        public static void Run(TestRun run, bool et)
        {
            var s = new XlSession(et);
            RunNeg(run, s);
            RunLife(run, s);
            RunGuard(run, s);
        }

        private static void RunNeg(TestRun run, XlSession s)
        {
            ChannelRegistry.ClearAll();
            ToolResult noCh = s.Pivot(
                null,
                "create",
                Src("利润表", "A1:C4"),
                Dest("利润表", "H1"),
                Rows(),
                null,
                SumVals());
            ExpectFail(run, "xlp-neg-no-channel", noCh, "打开");

            var req = new SpreadsheetPivotRequest
            {
                Action = "create",
                SourceSheet = "利润表",
                SourceRangeA1 = "A1:C4",
                DestSheet = "利润表",
                DestCellA1 = "H1",
                Rows = new List<string> { "地区" },
                Columns = new List<string>(),
                Values = new List<SpreadsheetPivotValueField>
                {
                    new SpreadsheetPivotValueField { Field = "金额", Agg = "sum" }
                }
            };
            if (!SpreadsheetHostAdapter.TryPivot(new FakeWordChannel(), req, out _, out ToolResult wrong))
            {
                ExpectFail(run, "xlp-neg-wrong-kind", wrong, "excel");
            }
            else
            {
                run.CaseFail("xlp-neg-wrong-kind", "word 渠道应失败");
            }

            WithBook(run, s, "xlp-neg-no-action", (id, ch) =>
            {
                ToolResult r = s.Call("F_excel_pivot", new Dictionary<string, object>
                {
                    ["channel_id"] = id
                });
                return r.Success ? "缺 action 应失败" : (Has(r.Error, "action") ? null : r.Error);
            });

            WithBook(run, s, "xlp-neg-bad-action", (id, ch) =>
            {
                ToolResult r = s.Pivot(id, "add");
                return r.Success ? "非法 action 应失败" : (Has(r.Error, "action") ? null : r.Error);
            });

            WithBook(run, s, "xlp-neg-create-no-source", (id, ch) =>
            {
                string name = s.FirstSheetName(ch);
                ToolResult r = s.Pivot(id, "create", dest: Dest(name, "H1"), values: SumVals());
                return r.Success ? "缺 source 应失败" : (Has(r.Error, "source") ? null : r.Error);
            });

            WithBook(run, s, "xlp-neg-create-no-dest", (id, ch) =>
            {
                string name = s.FirstSheetName(ch);
                ToolResult r = s.Pivot(id, "create", source: Src(name, "A1:C4"), values: SumVals());
                return r.Success ? "缺 dest 应失败" : (Has(r.Error, "dest") ? null : r.Error);
            });

            WithBook(run, s, "xlp-neg-create-no-values", (id, ch) =>
            {
                string name = s.FirstSheetName(ch);
                if (!s.SeedDetail(ch, name))
                {
                    return "铺底失败";
                }

                ToolResult r = s.Pivot(
                    id,
                    "create",
                    Src(name, "A1:C4"),
                    Dest(name, "H1"),
                    Rows(),
                    null,
                    new List<Dictionary<string, object>>());
                return r.Success ? "空 values 应失败" : (Has(r.Error, "values") ? null : r.Error);
            });

            WithBook(run, s, "xlp-neg-bang", (id, ch) =>
            {
                string name = s.FirstSheetName(ch);
                ToolResult r = s.Pivot(
                    id,
                    "create",
                    Src(name, name + "!A1:C4"),
                    Dest(name, "H1"),
                    Rows(),
                    null,
                    SumVals());
                return r.Success ? "含!应失败" : (Has(r.Error, "!") || Has(r.Error, "分开") ? null : r.Error);
            });

            WithBook(run, s, "xlp-neg-dest-rect", (id, ch) =>
            {
                string name = s.FirstSheetName(ch);
                ToolResult r = s.Pivot(
                    id,
                    "create",
                    Src(name, "A1:C4"),
                    Dest(name, "H1:L20"),
                    Rows(),
                    null,
                    SumVals());
                return r.Success ? "dest 矩形应失败" : (Has(r.Error, "单格") ? null : r.Error);
            });

            WithBook(run, s, "xlp-neg-missing-sheet", (id, ch) =>
            {
                string name = s.FirstSheetName(ch);
                if (!s.SeedDetail(ch, name))
                {
                    return "铺底失败";
                }

                ToolResult r = s.Pivot(
                    id,
                    "create",
                    Src(name, "A1:C4"),
                    Dest("不存在的表", "H1"),
                    Rows(),
                    null,
                    SumVals());
                return r.Success
                    ? "缺表应失败"
                    : (Has(r.Error, "不存在") || Has(r.Error, "未找到") ? null : r.Error);
            });
        }

        private static void RunLife(TestRun run, XlSession s)
        {
            WithBook(run, s, "xlp-create-sum", (id, ch) =>
            {
                string name = s.FirstSheetName(ch);
                if (!s.SeedDetail(ch, name))
                {
                    return "铺底失败";
                }

                ToolResult r = CreateSum(s, id, name);
                if (!r.Success || string.IsNullOrEmpty(Str(r, "name"))
                    || !string.Equals(Str(r, "dest_cell"), "H1", StringComparison.OrdinalIgnoreCase))
                {
                    return r.Success
                        ? "name/dest=" + Str(r, "name") + "/" + Str(r, "dest_cell")
                        : r.Error;
                }

                ToolResult g = s.Pivot(id, "list");
                if (!g.Success || !HasDest(g, "H1"))
                {
                    return g.Error ?? "list 看不到";
                }

                ToolResult rd = s.Read(id, name, "A1:C4");
                return rd.Success && Has(Str(rd, "table"), "10") && Has(Str(rd, "table"), "金额")
                    ? null
                    : (rd.Error ?? "源被改了");
            });

            WithBook(run, s, "xlp-create-avg", (id, ch) =>
            {
                string name = s.FirstSheetName(ch);
                if (!s.SeedDetail(ch, name))
                {
                    return "铺底失败";
                }

                ToolResult r = s.Pivot(
                    id,
                    "create",
                    Src(name, "A1:C4"),
                    Dest(name, "H1"),
                    Rows(),
                    null,
                    Vals("金额", "average"));
                if (!r.Success || string.IsNullOrEmpty(Str(r, "name")))
                {
                    return r.Error ?? "无 name";
                }

                ToolResult rd = s.Read(id, name, "C2");
                return rd.Success && Has(Str(rd, "table"), "10")
                    ? null
                    : (rd.Error ?? "源被改了");
            });

            WithBook(run, s, "xlp-create-named", (id, ch) =>
            {
                string name = s.FirstSheetName(ch);
                if (!s.SeedDetail(ch, name))
                {
                    return "铺底失败";
                }

                ToolResult r = s.Pivot(
                    id,
                    "create",
                    Src(name, "A1:C4"),
                    Dest(name, "H1"),
                    Rows(),
                    null,
                    SumVals(),
                    "销售透视");
                if (!r.Success || !string.Equals(Str(r, "name"), "销售透视", StringComparison.Ordinal))
                {
                    return r.Success ? "name=" + Str(r, "name") : r.Error;
                }

                ToolResult g = s.Pivot(id, "list");
                return g.Success && HasName(g, "销售透视")
                    ? null
                    : (g.Error ?? "list 无销售透视");
            });

            WithBook(run, s, "xlp-rows-cols", (id, ch) =>
            {
                string name = s.FirstSheetName(ch);
                if (!s.SeedDetail(ch, name))
                {
                    return "铺底失败";
                }

                ToolResult r = s.Pivot(
                    id,
                    "create",
                    Src(name, "A1:C4"),
                    Dest(name, "H1"),
                    new[] { "地区" },
                    new[] { "产品" },
                    SumVals());
                if (!r.Success)
                {
                    return r.Error;
                }

                ToolResult g = s.Pivot(id, "list");
                Dictionary<string, object> one = FindDest(g, "H1");
                string area = Str(one, "table_range");
                if (string.IsNullOrEmpty(area))
                {
                    return g.Error ?? "无 table_range";
                }

                ToolResult rd = s.Read(id, name, area);
                string table = Str(rd, "table");
                return rd.Success && (Has(table, "东") || Has(table, "甲"))
                    ? null
                    : (rd.Error ?? ("结果区无标签 " + table));
            });

            WithBook(run, s, "xlp-refresh", (id, ch) =>
            {
                string name = s.FirstSheetName(ch);
                if (!s.SeedDetail(ch, name))
                {
                    return "铺底失败";
                }

                ToolResult created = CreateSum(s, id, name);
                if (!created.Success)
                {
                    return created.Error;
                }

                string csv = s.WriteCsv("100");
                ToolResult w = s.Write(id, name, "C2", csv);
                if (!w.Success)
                {
                    return w.Error;
                }

                ToolResult rf = s.Pivot(id, "refresh", dest: Dest(name, "H1"));
                if (!rf.Success)
                {
                    return rf.Error;
                }

                object raw = s.ReadValue2(ch, name, "C2");
                if (!NearNumber(raw, 100))
                {
                    return "源 C2 不是 100: " + raw;
                }

                ToolResult g = s.Pivot(id, "list");
                string area = Str(FindDest(g, "H1"), "table_range");
                if (string.IsNullOrEmpty(area))
                {
                    return "无 table_range";
                }

                ToolResult rd = s.Read(id, name, area);
                string table = Str(rd, "table");
                return rd.Success && (Has(table, "100") || Has(table, "150"))
                    ? null
                    : (rd.Error ?? ("refresh 后结果未变 " + table));
            });

            WithBook(run, s, "xlp-delete", (id, ch) =>
            {
                string name = s.FirstSheetName(ch);
                if (!s.SeedDetail(ch, name))
                {
                    return "铺底失败";
                }

                ToolResult created = CreateSum(s, id, name, "销售透视");
                if (!created.Success)
                {
                    return created.Error;
                }

                ToolResult d = s.Pivot(id, "delete", name: "销售透视");
                if (!d.Success)
                {
                    return d.Error;
                }

                ToolResult g = s.Pivot(id, "list");
                if (!g.Success || HasName(g, "销售透视") || Int(g, "pivots_count") != 0)
                {
                    return g.Error ?? "delete 后还在";
                }

                ToolResult rd = s.Read(id, name, "A1:C4");
                return rd.Success && Has(Str(rd, "table"), "10")
                    ? null
                    : (rd.Error ?? "源被删了");
            });

            WithBook(run, s, "xlp-list-book", (id, ch) =>
            {
                string name = s.FirstSheetName(ch);
                if (!s.SeedDetail(ch, name))
                {
                    return "铺底失败";
                }

                ToolResult a = CreateSum(s, id, name, destCell: "H1");
                ToolResult b = CreateSum(s, id, name, destCell: "N1");
                if (!a.Success || !b.Success)
                {
                    return a.Error ?? b.Error;
                }

                ToolResult g = s.Pivot(id, "list");
                return g.Success && Int(g, "pivots_count") == 2 && HasDest(g, "H1") && HasDest(g, "N1")
                    ? null
                    : (g.Error ?? "整簿 list 不是 2");
            });

            WithBook(run, s, "xlp-list-sheet", (id, ch) =>
            {
                string src = s.FirstSheetName(ch);
                if (!s.SeedTwoSheets(ch, src, "分析") || !s.SeedDetail(ch, src) || !s.SeedDetail(ch, "分析"))
                {
                    return "铺两表失败";
                }

                ToolResult a = CreateSum(s, id, src, destSheet: src, destCell: "H1");
                ToolResult b = CreateSum(s, id, "分析", destSheet: "分析", destCell: "H1");
                if (!a.Success || !b.Success)
                {
                    return a.Error ?? b.Error;
                }

                ToolResult g = s.Pivot(id, "list", sheet: "分析");
                if (!g.Success || Int(g, "pivots_count") != 1)
                {
                    return g.Error ?? ("count=" + Int(g, "pivots_count"));
                }

                Dictionary<string, object> one = FirstPivot(g);
                return string.Equals(Str(one, "sheet"), "分析", StringComparison.Ordinal)
                    ? null
                    : "sheet=" + Str(one, "sheet");
            });

            WithBook(run, s, "xlp-hidden", (id, ch) =>
            {
                if (!s.SeedHidden(ch, "可见", "底稿") || !s.SeedDetail(ch, "底稿"))
                {
                    return "铺隐藏表失败";
                }

                ToolResult r = CreateSum(s, id, "底稿", destSheet: "底稿", destCell: "H1", name: "隐透视");
                if (!r.Success)
                {
                    return r.Error;
                }

                ToolResult g = s.Pivot(id, "list", sheet: "底稿");
                return g.Success && HasName(g, "隐透视")
                    ? null
                    : (g.Error ?? "隐藏表透视看不到");
            });

            WithBook(run, s, "xlp-zh-sheet", (id, ch) =>
            {
                if (!s.SeedTwoSheets(ch, "明细", "分析") || !s.SeedDetail(ch, "明细"))
                {
                    return "改名失败";
                }

                ToolResult r = CreateSum(s, id, "明细", destSheet: "分析", destCell: "H1");
                if (!r.Success || !string.Equals(Str(r, "sheet"), "分析", StringComparison.Ordinal))
                {
                    return r.Success ? "create sheet=" + Str(r, "sheet") : r.Error;
                }

                ToolResult g = s.Pivot(id, "list", sheet: "分析");
                Dictionary<string, object> one = FirstPivot(g);
                return g.Success && string.Equals(Str(one, "sheet"), "分析", StringComparison.Ordinal)
                    ? null
                    : (g.Success ? "list sheet=" + Str(one, "sheet") : g.Error);
            });
        }

        private static void RunGuard(TestRun run, XlSession s)
        {
            WithBook(run, s, "xlp-neg-conflict", (id, ch) =>
            {
                string name = s.FirstSheetName(ch);
                if (!s.SeedDetail(ch, name))
                {
                    return "铺底失败";
                }

                ToolResult first = CreateSum(s, id, name);
                if (!first.Success)
                {
                    return first.Error;
                }

                ToolResult r = CreateSum(s, id, name);
                return r.Success ? "同 dest 应失败" : (Has(r.Error, "已存在") ? null : r.Error);
            });

            WithBook(run, s, "xlp-neg-dup-name", (id, ch) =>
            {
                string name = s.FirstSheetName(ch);
                if (!s.SeedDetail(ch, name))
                {
                    return "铺底失败";
                }

                ToolResult first = CreateSum(s, id, name, "销售透视", destCell: "H1");
                if (!first.Success)
                {
                    return first.Error;
                }

                ToolResult r = CreateSum(s, id, name, "销售透视", destCell: "N1");
                return r.Success ? "同名应失败" : (Has(r.Error, "同名") || Has(r.Error, "已存在") ? null : r.Error);
            });

            WithBook(run, s, "xlp-neg-bad-field", (id, ch) =>
            {
                string name = s.FirstSheetName(ch);
                if (!s.SeedDetail(ch, name))
                {
                    return "铺底失败";
                }

                ToolResult r = s.Pivot(
                    id,
                    "create",
                    Src(name, "A1:C4"),
                    Dest(name, "H1"),
                    new[] { "不存在的列" },
                    null,
                    SumVals());
                return r.Success ? "坏字段应失败" : (Has(r.Error, "表头") || Has(r.Error, "字段") ? null : r.Error);
            });

            WithBook(run, s, "xlp-neg-empty-header", (id, ch) =>
            {
                string name = s.FirstSheetName(ch);
                if (!s.SeedCell(ch, name, "A1", "地区")
                    || !s.SeedCell(ch, name, "C1", "金额")
                    || !s.SeedCell(ch, name, "A2", "东")
                    || !s.SeedCell(ch, name, "C2", 10))
                {
                    return "铺底失败";
                }

                ToolResult r = s.Pivot(
                    id,
                    "create",
                    Src(name, "A1:C2"),
                    Dest(name, "H1"),
                    Rows(),
                    null,
                    SumVals());
                return r.Success ? "空表头应失败" : (Has(r.Error, "非空") || Has(r.Error, "表头") ? null : r.Error);
            });

            WithBook(run, s, "xlp-neg-header-only", (id, ch) =>
            {
                string name = s.FirstSheetName(ch);
                if (!s.SeedCell(ch, name, "A1", "地区")
                    || !s.SeedCell(ch, name, "B1", "产品")
                    || !s.SeedCell(ch, name, "C1", "金额"))
                {
                    return "铺底失败";
                }

                ToolResult r = s.Pivot(
                    id,
                    "create",
                    Src(name, "A1:C1"),
                    Dest(name, "H1"),
                    Rows(),
                    null,
                    SumVals());
                return r.Success ? "仅表头应失败" : (Has(r.Error, "一行数据") || Has(r.Error, "表头") ? null : r.Error);
            });

            WithBook(run, s, "xlp-neg-merge-source", (id, ch) =>
            {
                string name = s.FirstSheetName(ch);
                if (!s.SeedMerge(ch, name)
                    || !s.SeedCell(ch, name, "A2", "东")
                    || !s.SeedCell(ch, name, "C2", 10))
                {
                    return "铺合并失败";
                }

                ToolResult r = s.Pivot(
                    id,
                    "create",
                    Src(name, "A1:C2"),
                    Dest(name, "H1"),
                    Rows(),
                    null,
                    SumVals());
                return r.Success ? "合并源应失败" : (Has(r.Error, "合并") ? null : r.Error);
            });

            WithBook(run, s, "xlp-neg-row-col-same", (id, ch) =>
            {
                string name = s.FirstSheetName(ch);
                if (!s.SeedDetail(ch, name))
                {
                    return "铺底失败";
                }

                ToolResult r = s.Pivot(
                    id,
                    "create",
                    Src(name, "A1:C4"),
                    Dest(name, "H1"),
                    new[] { "地区" },
                    new[] { "地区" },
                    SumVals());
                return r.Success
                    ? "rows=columns 应失败"
                    : (Has(r.Error, "同时") || Has(r.Error, "rows") ? null : r.Error);
            });

            WithBook(run, s, "xlp-neg-oversize", (id, ch) =>
            {
                string name = s.FirstSheetName(ch);
                ToolResult r = s.Pivot(
                    id,
                    "create",
                    Src(name, "A1:A5001"),
                    Dest(name, "H1"),
                    Rows(),
                    null,
                    SumVals());
                return r.Success ? "超限应失败" : (Has(r.Error, "过大") || Has(r.Error, "5000") ? null : r.Error);
            });

            WithBook(run, s, "xlp-neg-chart", (id, ch) =>
            {
                string name = s.FirstSheetName(ch);
                if (!s.SeedDetail(ch, name) || !s.SeedChartSheet(ch, out string chartName))
                {
                    return "铺图表 sheet 失败";
                }

                ToolResult r = s.Pivot(
                    id,
                    "create",
                    Src(name, "A1:C4"),
                    Dest(chartName, "H1"),
                    Rows(),
                    null,
                    SumVals());
                return r.Success ? "图表 dest 应失败" : (Has(r.Error, "sheet_type=chart") ? null : r.Error);
            });

            WithBook(run, s, "xlp-neg-refresh-missing", (id, ch) =>
            {
                string name = s.FirstSheetName(ch);
                ToolResult r = s.Pivot(id, "refresh", dest: Dest(name, "H1"));
                return r.Success ? "无透视 refresh 应失败" : (Has(r.Error, "未找到") ? null : r.Error);
            });
        }

        private static ToolResult CreateSum(
            XlSession s,
            string id,
            string sourceSheet,
            string name = null,
            string destSheet = null,
            string destCell = "H1")
        {
            return s.Pivot(
                id,
                "create",
                Src(sourceSheet, "A1:C4"),
                Dest(destSheet ?? sourceSheet, destCell),
                Rows(),
                null,
                SumVals(),
                name);
        }

        private static Dictionary<string, object> Src(string sheet, string range)
        {
            return new Dictionary<string, object>
            {
                ["sheet"] = sheet,
                ["range"] = range
            };
        }

        private static Dictionary<string, object> Dest(string sheet, string cell)
        {
            return new Dictionary<string, object>
            {
                ["sheet"] = sheet,
                ["cell"] = cell
            };
        }

        private static string[] Rows()
        {
            return new[] { "地区" };
        }

        private static List<Dictionary<string, object>> SumVals()
        {
            return Vals("金额", "sum");
        }

        private static List<Dictionary<string, object>> Vals(string field, string agg)
        {
            return new List<Dictionary<string, object>>
            {
                new Dictionary<string, object>
                {
                    ["field"] = field,
                    ["agg"] = agg
                }
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

        private static List<Dictionary<string, object>> Pivots(ToolResult r)
        {
            var list = new List<Dictionary<string, object>>();
            Dictionary<string, object> data = Data(r);
            if (data == null || !data.ContainsKey("pivots") || !(data["pivots"] is IEnumerable raw))
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

        private static Dictionary<string, object> FirstPivot(ToolResult r)
        {
            List<Dictionary<string, object>> list = Pivots(r);
            return list.Count > 0 ? list[0] : null;
        }

        private static Dictionary<string, object> FindDest(ToolResult r, string destCell)
        {
            foreach (Dictionary<string, object> one in Pivots(r))
            {
                if (string.Equals(Str(one, "dest_cell"), destCell, StringComparison.OrdinalIgnoreCase))
                {
                    return one;
                }
            }

            return null;
        }

        private static bool HasDest(ToolResult r, string destCell)
        {
            return FindDest(r, destCell) != null;
        }

        private static bool HasName(ToolResult r, string name)
        {
            foreach (Dictionary<string, object> one in Pivots(r))
            {
                if (string.Equals(Str(one, "name"), name, StringComparison.Ordinal))
                {
                    return true;
                }
            }

            return false;
        }

        private static bool NearNumber(object raw, double expect)
        {
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
