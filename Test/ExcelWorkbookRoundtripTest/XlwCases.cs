using System;
using System.Collections.Generic;
using System.IO;
using WordAddIn1;
using WordAddIn1.SpreadsheetHost;

namespace ExcelWorkbookRoundtripTest
{
    internal static class XlwCases
    {
        public static void Run(TestRun run, bool et)
        {
            var s = new XlSession(et);
            RunNeg(run, s);
            RunCsv(run, s);
            RunConstraint(run, s);
            RunClear(run, s);
        }

        private static void RunNeg(TestRun run, XlSession s)
        {
            ChannelRegistry.ClearAll();
            ToolResult noCh = s.Write(null, "利润表", "A1", s.WriteCsv("1"));
            ExpectFail(run, "xlw-neg-no-channel", noCh, "打开");

            var req = new SpreadsheetWriteRequest
            {
                SheetName = "利润表",
                RangeA1 = "A1",
                Mode = SpreadsheetWriteMode.Csv,
                CsvGrid = new List<IReadOnlyList<string>> { new List<string> { "1" } }
            };
            if (!SpreadsheetHostAdapter.TryWriteRange(new FakeWordChannel(), req, out _, out ToolResult wrong))
            {
                ExpectFail(run, "xlw-neg-wrong-kind", wrong, "unsupported");
            }
            else
            {
                run.CaseFail("xlw-neg-wrong-kind", "word 渠道应失败");
            }

            WithBook(run, s, "xlw-neg-no-sheet", (id, ch) =>
            {
                ToolResult r = s.Call("F_write_excel_range", new Dictionary<string, object>
                {
                    ["channel_id"] = id,
                    ["range"] = "A1",
                    ["csv_filename"] = s.WriteCsv("1")
                });
                return r.Success ? "缺 sheet 应失败" : (Has(r.Error, "sheet") ? null : r.Error);
            });

            WithBook(run, s, "xlw-neg-no-range", (id, ch) =>
            {
                ToolResult r = s.Call("F_write_excel_range", new Dictionary<string, object>
                {
                    ["channel_id"] = id,
                    ["sheet"] = s.FirstSheetName(ch),
                    ["csv_filename"] = s.WriteCsv("1")
                });
                return r.Success ? "缺 range 应失败" : (Has(r.Error, "range") ? null : r.Error);
            });

            WithBook(run, s, "xlw-neg-missing-sheet", (id, ch) =>
            {
                ToolResult r = s.Write(id, "不存在的表", "A1", s.WriteCsv("1"));
                return r.Success ? "缺表应失败" : (Has(r.Error, "不存在") ? null : r.Error);
            });

            WithBook(run, s, "xlw-neg-bang", (id, ch) =>
            {
                string name = s.FirstSheetName(ch);
                ToolResult r = s.Write(id, name, name + "!A1", s.WriteCsv("1"));
                return r.Success ? "含!应失败" : (Has(r.Error, "纯 A1") || Has(r.Error, "A1") ? null : r.Error);
            });

            WithBook(run, s, "xlw-neg-mutex", (id, ch) =>
            {
                ToolResult r = s.Call("F_write_excel_range", new Dictionary<string, object>
                {
                    ["channel_id"] = id,
                    ["sheet"] = s.FirstSheetName(ch),
                    ["range"] = "A1",
                    ["clear"] = true,
                    ["csv_filename"] = s.WriteCsv("1")
                });
                return r.Success ? "互斥应失败" : (Has(r.Error, "互斥") ? null : r.Error);
            });

            WithBook(run, s, "xlw-neg-no-payload", (id, ch) =>
            {
                ToolResult r = s.Call("F_write_excel_range", new Dictionary<string, object>
                {
                    ["channel_id"] = id,
                    ["sheet"] = s.FirstSheetName(ch),
                    ["range"] = "A1"
                });
                return r.Success ? "无载荷应失败" : (Has(r.Error, "clear") || Has(r.Error, "csv") || Has(r.Error, "path")
                    ? null
                    : r.Error);
            });

            WithBook(run, s, "xlw-neg-missing-csv", (id, ch) =>
            {
                string missing = Path.Combine(Path.GetTempPath(), "ew-xlw-missing-" + Guid.NewGuid().ToString("N") + ".csv");
                ToolResult r = s.Write(id, s.FirstSheetName(ch), "A1", missing);
                return r.Success ? "缺文件应失败" : (Has(r.Error, "不存在") || Has(r.Error, "CSV") ? null : r.Error);
            });

            WithBook(run, s, "xlw-neg-bad-a1", (id, ch) =>
            {
                ToolResult r = s.Write(id, s.FirstSheetName(ch), "ZZ", s.WriteCsv("1"));
                return r.Success ? "非法 range 应失败" : (Has(r.Error, "非法") || Has(r.Error, "range") ? null : r.Error);
            });
        }

        private static void RunCsv(TestRun run, XlSession s)
        {
            WithBook(run, s, "xlw-csv-anchor", (id, ch) =>
            {
                string name = s.FirstSheetName(ch);
                ToolResult w = s.Write(id, name, "C3", s.WriteCsv("地区,销量", "东,10"));
                if (!w.Success)
                {
                    return w.Error;
                }

                if (!string.Equals(Str(w, "mode"), "csv", StringComparison.Ordinal)
                    || !string.Equals(Str(w, "actual_range"), "C3:D4", StringComparison.OrdinalIgnoreCase)
                    || Int(w, "written_count") != 4)
                {
                    return "回包 mode/actual_range/written_count 不符 "
                        + Str(w, "mode") + " " + Str(w, "actual_range") + " " + Int(w, "written_count");
                }

                ToolResult rd = s.Read(id, name, "C3:D4");
                return rd.Success && Has(Str(rd, "table"), "地区") && Has(Str(rd, "table"), "10")
                    ? null
                    : (rd.Error ?? ("回读 " + Str(rd, "table")));
            });

            WithBook(run, s, "xlw-csv-rect", (id, ch) =>
            {
                string name = s.FirstSheetName(ch);
                ToolResult w = s.Write(id, name, "A1:B2", s.WriteCsv("地区,销量", "东,10"));
                if (!w.Success || !string.Equals(Str(w, "actual_range"), "A1:B2", StringComparison.OrdinalIgnoreCase))
                {
                    return w.Success ? "actual_range=" + Str(w, "actual_range") : w.Error;
                }

                ToolResult rd = s.Read(id, name, "A1:B2");
                return rd.Success && Has(Str(rd, "table"), "地区") && Has(Str(rd, "table"), "东")
                    ? null
                    : (rd.Error ?? ("回读 " + Str(rd, "table")));
            });

            WithBook(run, s, "xlw-anchor-expand", (id, ch) =>
            {
                string name = s.FirstSheetName(ch);
                ToolResult w = s.Write(id, name, "C3", s.WriteCsv("上", "下"));
                if (!w.Success)
                {
                    return w.Error;
                }

                return string.Equals(Str(w, "actual_range"), "C3:C4", StringComparison.OrdinalIgnoreCase)
                    ? null
                    : "actual_range=" + Str(w, "actual_range");
            });

            WithBook(run, s, "xlw-csv-formula", (id, ch) =>
            {
                string name = s.FirstSheetName(ch);
                if (!s.SeedCell(ch, name, "A2", 10))
                {
                    return "铺 A2 失败";
                }

                ToolResult w = s.Write(id, name, "B2", s.WriteCsv("=A2+1"));
                if (!w.Success)
                {
                    return w.Error;
                }

                ToolResult fm = s.Read(id, name, "B2", "formula");
                if (!fm.Success || !Has(Str(fm, "table"), "="))
                {
                    return fm.Error ?? ("公式回读 " + Str(fm, "table"));
                }

                ToolResult val = s.Read(id, name, "B2", "value");
                if (!val.Success || !Has(Str(val, "table"), "11"))
                {
                    return val.Error ?? ("值回读 " + Str(val, "table"));
                }

                return Has(Str(val, "table"), "=A2") ? "显示值里不应是公式串" : null;
            });

            WithBook(run, s, "xlw-csv-text-eq", (id, ch) =>
            {
                string name = s.FirstSheetName(ch);
                ToolResult w = s.Write(id, name, "A1", s.WriteCsv("\"'=注意\""));
                if (!w.Success)
                {
                    return w.Error;
                }

                if (s.CellHasFormula(ch, name, "A1"))
                {
                    return "写成了公式";
                }

                ToolResult rd = s.Read(id, name, "A1");
                return rd.Success && Has(Str(rd, "table"), "注意")
                    ? null
                    : (rd.Error ?? ("回读 " + Str(rd, "table")));
            });

            WithBook(run, s, "xlw-csv-blank", (id, ch) =>
            {
                string name = s.FirstSheetName(ch);
                if (!s.SeedCell(ch, name, "A1", "旧"))
                {
                    return "铺底失败";
                }

                ToolResult w = s.Write(id, name, "A1", s.WriteCsv(""));
                if (w.Success)
                {
                    return "空 CSV 行会被跳过，应另写空字段";
                }

                w = s.Write(id, name, "A1:B1", s.WriteCsv(",x"));
                if (!w.Success)
                {
                    return w.Error;
                }

                ToolResult rd = s.Read(id, name, "A1:B1");
                if (!rd.Success)
                {
                    return rd.Error;
                }

                return Has(Str(rd, "table"), "旧") ? "A1 应被清空" : null;
            });

            WithBook(run, s, "xlw-csv-zh", (id, ch) =>
            {
                if (!s.SeedUsed(ch, "利润表"))
                {
                    return "铺底失败";
                }

                ToolResult w = s.Write(id, "利润表", "A1:B1", s.WriteCsv("地区,销量"));
                if (!w.Success)
                {
                    return w.Error;
                }

                ToolResult rd = s.Read(id, "利润表", "A1:B1");
                return rd.Success && Has(Str(rd, "table"), "地区") && Has(Str(rd, "table"), "销量")
                    ? null
                    : (rd.Error ?? ("回读 " + Str(rd, "table")));
            });

            WithBook(run, s, "xlw-csv-hidden", (id, ch) =>
            {
                if (!s.SeedHidden(ch, "可见表", "隐藏底稿"))
                {
                    return "铺隐藏失败";
                }

                ToolResult w = s.Write(id, "隐藏底稿", "A1", s.WriteCsv("隐写"));
                if (!w.Success)
                {
                    return w.Error;
                }

                ToolResult rd = s.Read(id, "隐藏底稿", "A1");
                return rd.Success && Has(Str(rd, "table"), "隐写")
                    ? null
                    : (rd.Error ?? ("回读 " + Str(rd, "table")));
            });

            WithBook(run, s, "xlw-overwrite", (id, ch) =>
            {
                string name = s.FirstSheetName(ch);
                if (!s.SeedCell(ch, name, "A1", "旧字"))
                {
                    return "铺底失败";
                }

                ToolResult w = s.Write(id, name, "A1", s.WriteCsv("新字"));
                if (!w.Success)
                {
                    return w.Error;
                }

                ToolResult rd = s.Read(id, name, "A1");
                if (!rd.Success)
                {
                    return rd.Error;
                }

                return Has(Str(rd, "table"), "新字") && !Has(Str(rd, "table"), "旧字")
                    ? null
                    : "回读 " + Str(rd, "table");
            });

            WithBook(run, s, "xlw-keep-nf", (id, ch) =>
            {
                string name = s.FirstSheetName(ch);
                if (!s.SeedNumberFormat(ch, name, "A1", "0%"))
                {
                    return "铺格式失败";
                }

                ToolResult w = s.Write(id, name, "A1", s.WriteCsv("0.15"));
                if (!w.Success)
                {
                    return w.Error;
                }

                string nf = s.ReadNumberFormat(ch, name, "A1");
                if (nf.IndexOf("%", StringComparison.Ordinal) < 0)
                {
                    return "NumberFormat 被改成 " + nf;
                }

                ToolResult rd = s.Read(id, name, "A1");
                return rd.Success && (Has(Str(rd, "table"), "15") || Has(Str(rd, "table"), "0.15"))
                    ? null
                    : (rd.Error ?? ("回读 " + Str(rd, "table")));
            });

            WithBook(run, s, "xlw-readback", (id, ch) =>
            {
                string name = s.FirstSheetName(ch);
                ToolResult w = s.Write(id, name, "A1:B1", s.WriteCsv("甲,乙"));
                if (!w.Success)
                {
                    return w.Error;
                }

                ToolResult rd = s.Read(id, name, Str(w, "actual_range"));
                return rd.Success
                    && Has(Str(rd, "table"), "addr=")
                    && Has(Str(rd, "table"), "甲")
                    && Has(Str(rd, "table"), "乙")
                    ? null
                    : (rd.Error ?? ("回读 " + Str(rd, "table")));
            });

            WithBook(run, s, "xlw-export-write", (id, ch) =>
            {
                if (!s.SeedUsed(ch, "利润表"))
                {
                    return "铺底失败";
                }

                string export = s.NewCsvPath();
                ToolResult rd = s.Read(id, "利润表", "A1:B2", null, export);
                if (!rd.Success)
                {
                    return rd.Error;
                }

                if (!File.Exists(export))
                {
                    return "export_csv 未落盘";
                }

                string text = File.ReadAllText(export);
                if (text.IndexOf("东", StringComparison.Ordinal) < 0)
                {
                    return "导出没有东: " + text;
                }

                File.WriteAllText(export, text.Replace("东", "西"), new System.Text.UTF8Encoding(true));
                ToolResult w = s.Write(id, "利润表", "A1:B2", export);
                if (!w.Success)
                {
                    return w.Error;
                }

                ToolResult again = s.Read(id, "利润表", "A2");
                return again.Success && Has(Str(again, "table"), "西")
                    ? null
                    : (again.Error ?? ("回读 " + Str(again, "table")));
            });
        }

        private static void RunConstraint(TestRun run, XlSession s)
        {
            WithBook(run, s, "xlw-neg-rect-mismatch", (id, ch) =>
            {
                string name = s.FirstSheetName(ch);
                if (!s.SeedCell(ch, name, "A1", "旧"))
                {
                    return "铺底失败";
                }

                ToolResult w = s.Write(id, name, "A1:B2", s.WriteCsv("a", "b", "c"));
                if (w.Success)
                {
                    return "尺寸不符应失败";
                }

                if (!Has(w.Error, "尺寸") && !Has(w.Error, "不一致"))
                {
                    return w.Error;
                }

                ToolResult rd = s.Read(id, name, "A1");
                return rd.Success && Has(Str(rd, "table"), "旧")
                    ? null
                    : "原格被改 " + Str(rd, "table");
            });

            WithBook(run, s, "xlw-neg-oversize", (id, ch) =>
            {
                string name = s.FirstSheetName(ch);
                if (!s.SeedCell(ch, name, "A1", "旧"))
                {
                    return "铺底失败";
                }

                var lines = new string[101];
                for (int i = 0; i < lines.Length; i++)
                {
                    lines[i] = "v";
                }

                ToolResult w = s.Write(id, name, "A1", s.WriteCsv(lines));
                if (w.Success)
                {
                    return "超限应失败";
                }

                if (!Has(w.Error, "过大"))
                {
                    return w.Error;
                }

                ToolResult rd = s.Read(id, name, "A1");
                return rd.Success && Has(Str(rd, "table"), "旧")
                    ? null
                    : "半截写入 " + Str(rd, "table");
            });

            WithBook(run, s, "xlw-neg-chart", (id, ch) =>
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
                    return "概览未标 chart: " + (chart == null ? "无" : Convert.ToString(chart["sheet_type"]));
                }

                ToolResult w = s.Write(id, chartName, "A1", s.WriteCsv("1"));
                return w.Success ? "图表 sheet 应失败" : (Has(w.Error, "chart") ? null : w.Error);
            });

            WithBook(run, s, "xlw-neg-merge-body", (id, ch) =>
            {
                string name = s.FirstSheetName(ch);
                if (!s.SeedMerge(ch, name))
                {
                    return "铺合并失败";
                }

                ToolResult w = s.Write(id, name, "A1:C1", s.WriteCsv("标题,x,"));
                if (w.Success)
                {
                    return "合并续格非空应失败";
                }

                if (!Has(w.Error, "合并"))
                {
                    return w.Error;
                }

                ToolResult rd = s.Read(id, name, "A1:C1");
                return rd.Success && Has(Str(rd, "table"), "标题")
                    ? null
                    : "合并被硬改 " + Str(rd, "table");
            });
        }

        private static void RunClear(TestRun run, XlSession s)
        {
            WithBook(run, s, "xlw-merge-origin", (id, ch) =>
            {
                string name = s.FirstSheetName(ch);
                if (!s.SeedMerge(ch, name))
                {
                    return "铺合并失败";
                }

                ToolResult w = s.Write(id, name, "A1:C1", s.WriteCsv("新标题,,"));
                if (!w.Success)
                {
                    return w.Error;
                }

                ToolResult rd = s.Read(id, name, "A1:C1");
                if (!rd.Success)
                {
                    return rd.Error;
                }

                string table = Str(rd, "table");
                if (!Has(table, "新标题"))
                {
                    return "回读 " + table;
                }

                return Has(table, "area=") || Has(table, "colspan")
                    ? null
                    : "应有 area/colspan " + table;
            });

            WithBook(run, s, "xlw-clear", (id, ch) =>
            {
                if (!s.SeedUsed(ch, "利润表"))
                {
                    return "铺底失败";
                }

                ToolResult w = s.Write(id, "利润表", "A1:B2", clear: true);
                if (!w.Success || !string.Equals(Str(w, "mode"), "clear", StringComparison.Ordinal))
                {
                    return w.Success ? "mode=" + Str(w, "mode") : w.Error;
                }

                ToolResult rd = s.Read(id, "利润表", "A1:B2");
                return rd.Success && !Has(Str(rd, "table"), "地区") && !Has(Str(rd, "table"), "东")
                    ? null
                    : (rd.Error ?? ("未清空 " + Str(rd, "table")));
            });

            WithBook(run, s, "xlw-clear-keep", (id, ch) =>
            {
                string name = s.FirstSheetName(ch);
                if (!s.SeedCell(ch, name, "A1", "甲") || !s.SeedCell(ch, name, "C1", "丙"))
                {
                    return "铺底失败";
                }

                ToolResult w = s.Write(id, name, "A1", clear: true);
                if (!w.Success)
                {
                    return w.Error;
                }

                ToolResult rd = s.Read(id, name, "A1:C1");
                if (!rd.Success)
                {
                    return rd.Error;
                }

                return !Has(Str(rd, "table"), "甲") && Has(Str(rd, "table"), "丙")
                    ? null
                    : "回读 " + Str(rd, "table");
            });

            WithBook(run, s, "xlw-clear-formula", (id, ch) =>
            {
                if (!s.SeedFormula(ch, "公式表"))
                {
                    return "铺公式失败";
                }

                ToolResult w = s.Write(id, "公式表", "B2", clear: true);
                if (!w.Success)
                {
                    return w.Error;
                }

                if (s.CellHasFormula(ch, "公式表", "B2"))
                {
                    return "公式还在";
                }

                return null;
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

        private static Dictionary<string, object> SheetNamed(ToolResult r, string name)
        {
            Dictionary<string, object> data = Data(r);
            if (data == null || !data.ContainsKey("sheets") || !(data["sheets"] is System.Collections.IEnumerable raw))
            {
                return null;
            }

            foreach (object one in raw)
            {
                if (one is Dictionary<string, object> d
                    && string.Equals(Convert.ToString(d["name"] ?? ""), name, StringComparison.Ordinal))
                {
                    return d;
                }
            }

            return null;
        }
    }
}
