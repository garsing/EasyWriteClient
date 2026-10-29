using System;
using System.Collections.Generic;
using System.IO;
using WordAddIn1;

namespace ExcelWorkbookRoundtripTest
{
    /// <summary>
    /// 写 / 导 CSV / 套格式 / 条件格式 / 透视源 / 图源的 Bulk 上限。不灌满 5 万行。
    /// </summary>
    internal static class XlbCases
    {
        public const int Target = 15;

        public static void Run(TestRun run, bool et)
        {
            var s = new XlSession(et);
            RunRead(run, s);
            RunWrite(run, s);
            RunFormat(run, s);
            RunSource(run, s);
            RunUnchanged(run, s);
        }

        private static void RunRead(TestRun run, XlSession s)
        {
            WithBook(run, s, "xlb-rd-trunc", (id, ch) =>
            {
                ToolResult r = s.Read(id, s.FirstSheetName(ch), "A1:AN120");
                if (!r.Success)
                {
                    return r.Error;
                }

                if (!Bool(r, "truncated"))
                {
                    return "应 truncated";
                }

                string table = Str(r, "table");
                if (string.IsNullOrEmpty(table) || table.IndexOf("<table", StringComparison.OrdinalIgnoreCase) < 0)
                {
                    return "预览应回 table";
                }

                return null;
            });

            WithBook(run, s, "xlb-rd-path-ok", (id, ch) =>
            {
                string name = s.FirstSheetName(ch);
                var lines = new string[120];
                for (int i = 0; i < lines.Length; i++)
                {
                    lines[i] = "r" + i + ",x";
                }

                ToolResult w = s.Write(id, name, "A1", s.WriteCsv(lines));
                if (!w.Success)
                {
                    return w.Error;
                }

                string export = s.NewCsvPath();
                ToolResult r = s.Read(id, name, "A1:B120", null, export);
                if (!r.Success)
                {
                    return r.Error;
                }

                if (Bool(r, "truncated"))
                {
                    return "导出不应 truncated";
                }

                if (!string.IsNullOrEmpty(Str(r, "table")) || Has(Str(r, "display_contents"), "<table"))
                {
                    return "path 不应回 table";
                }

                if (!File.Exists(export))
                {
                    return "未落盘";
                }

                int data = 0;
                foreach (string line in File.ReadAllLines(export))
                {
                    if (!string.IsNullOrWhiteSpace(line))
                    {
                        data++;
                    }
                }

                return data == 120 ? null : "CSV 行数 " + data;
            });

            WithBook(run, s, "xlb-rd-path-oversize", (id, ch) =>
            {
                string export = s.NewCsvPath();
                ToolResult r = s.Read(id, s.FirstSheetName(ch), "A1:A50001", null, export);
                if (r.Success)
                {
                    return "超限应失败";
                }

                if (!Has(r.Error, "50000") || !Has(r.Error, "2500000") || !Has(r.Error, "拆成多次"))
                {
                    return r.Error;
                }

                return File.Exists(export) ? "不应留下半截 CSV" : null;
            });

            WithBook(run, s, "xlb-rd-path-cols", (id, ch) =>
            {
                string export = s.NewCsvPath();
                ToolResult r = s.Read(id, s.FirstSheetName(ch), "A1:AY1", null, export);
                if (r.Success)
                {
                    return "51 列应失败";
                }

                if (!Has(r.Error, "过大") || !Has(r.Error, "拆成多次"))
                {
                    return r.Error;
                }

                return File.Exists(export) ? "不应留下半截 CSV" : null;
            });
        }

        private static void RunWrite(TestRun run, XlSession s)
        {
            WithBook(run, s, "xlb-wr-oversize", (id, ch) =>
            {
                string name = s.FirstSheetName(ch);
                if (!s.SeedCell(ch, name, "A1", "旧"))
                {
                    return "铺底失败";
                }

                var lines = new string[50001];
                for (int i = 0; i < lines.Length; i++)
                {
                    lines[i] = "v";
                }

                ToolResult w = s.Write(id, name, "A1", s.WriteCsv(lines));
                if (w.Success)
                {
                    return "超限应失败";
                }

                if (!Has(w.Error, "过大") || !Has(w.Error, "50000") || !Has(w.Error, "90 行"))
                {
                    return w.Error;
                }

                ToolResult rd = s.Read(id, name, "A1");
                return rd.Success && Has(Str(rd, "table"), "旧")
                    ? null
                    : "半截写入 " + Str(rd, "table");
            });

            WithBook(run, s, "xlb-wr-cols", (id, ch) =>
            {
                string name = s.FirstSheetName(ch);
                if (!s.SeedCell(ch, name, "A1", "旧"))
                {
                    return "铺底失败";
                }

                var cells = new string[51];
                for (int i = 0; i < cells.Length; i++)
                {
                    cells[i] = "c" + i;
                }

                ToolResult w = s.Write(id, name, "A1", s.WriteCsv(string.Join(",", cells)));
                if (w.Success)
                {
                    return "51 列应失败";
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

            WithBook(run, s, "xlb-wr-chunk", (id, ch) =>
            {
                string name = s.FirstSheetName(ch);
                var lines = new string[2100];
                for (int i = 0; i < lines.Length; i++)
                {
                    lines[i] = "v" + i;
                }

                ToolResult w = s.Write(id, name, "A1", s.WriteCsv(lines));
                if (!w.Success)
                {
                    return w.Error;
                }

                ToolResult first = s.Read(id, name, "A1");
                ToolResult mid = s.Read(id, name, "A2000");
                ToolResult last = s.Read(id, name, "A2100");
                if (!first.Success || !mid.Success || !last.Success)
                {
                    return first.Error ?? mid.Error ?? last.Error;
                }

                if (!Has(Str(first, "table"), "v0"))
                {
                    return "首行 " + Str(first, "table");
                }

                if (!Has(Str(mid, "table"), "v1999"))
                {
                    return "第 2000 行 " + Str(mid, "table");
                }

                return Has(Str(last, "table"), "v2099")
                    ? null
                    : "末行 " + Str(last, "table");
            });
        }

        private static void RunFormat(TestRun run, XlSession s)
        {
            WithBook(run, s, "xlb-fmt-oversize", (id, ch) =>
            {
                ToolResult r = s.ApplyFormat(
                    id,
                    s.FirstSheetName(ch),
                    "A1:A50001",
                    new Dictionary<string, object> { ["bold"] = true });
                return r.Success ? "超限应失败" : (Has(r.Error, "过大") ? null : r.Error);
            });

            WithBook(run, s, "xlb-cnd-oversize", (id, ch) =>
            {
                ToolResult r = s.CondFormat(
                    id,
                    s.FirstSheetName(ch),
                    "A1:A50001",
                    "add",
                    new Dictionary<string, object>
                    {
                        ["type"] = "cell_value",
                        ["operator"] = "less",
                        ["formula1"] = "0",
                        ["format"] = new Dictionary<string, object> { ["fill_color"] = "#FFCCCC" }
                    });
                return r.Success ? "超限应失败" : (Has(r.Error, "过大") ? null : r.Error);
            });
        }

        private static void RunSource(TestRun run, XlSession s)
        {
            WithBook(run, s, "xlb-pvt-oversize", (id, ch) =>
            {
                string name = s.FirstSheetName(ch);
                ToolResult r = s.Pivot(
                    id,
                    "create",
                    new Dictionary<string, object> { ["sheet"] = name, ["range"] = "A1:A50001" },
                    new Dictionary<string, object> { ["sheet"] = name, ["cell"] = "H1" },
                    new[] { "地区" },
                    null,
                    new List<Dictionary<string, object>>
                    {
                        new Dictionary<string, object> { ["field"] = "金额", ["agg"] = "sum" }
                    });
                return r.Success ? "超限应失败" : (Has(r.Error, "过大") ? null : r.Error);
            });

            WithBook(run, s, "xlb-cht-src-oversize", (id, ch) =>
            {
                string name = s.FirstSheetName(ch);
                ToolResult r = s.Chart(
                    id,
                    "create",
                    new Dictionary<string, object> { ["sheet"] = name, ["range"] = "E2:K18" },
                    "<ChartConfig><Chart><Type>column</Type><Title>销量</Title>"
                    + "<DataTable><Source sheet=\"" + name + "\" range=\"A1:A50001\""
                    + " plotBy=\"columns\"/></DataTable></Chart></ChartConfig>");
                return r.Success ? "源超限应失败" : (Has(r.Error, "过大") ? null : r.Error);
            });
        }

        private static void RunUnchanged(TestRun run, XlSession s)
        {
            WithBook(run, s, "xlb-get-fmt-still", (id, ch) =>
            {
                ToolResult r = s.GetFormat(id, s.FirstSheetName(ch), "A1:A101");
                return r.Success ? "get 超限应失败" : (Has(r.Error, "过大") || Has(r.Error, "100") ? null : r.Error);
            });

            WithBook(run, s, "xlb-cht-dest-still", (id, ch) =>
            {
                string name = s.FirstSheetName(ch);
                ToolResult r = s.Chart(
                    id,
                    "create",
                    new Dictionary<string, object> { ["sheet"] = name, ["range"] = "A1:A81" },
                    "<ChartConfig><Chart><Type>column</Type><Title>销量</Title>"
                    + "<DataTable><Source sheet=\"" + name + "\" range=\"A1:B4\""
                    + " plotBy=\"columns\"/></DataTable></Chart></ChartConfig>");
                return r.Success
                    ? "dest 超限应失败"
                    : (Has(r.Error, "过大") || Has(r.Error, "80") ? null : r.Error);
            });

            WithBook(run, s, "xlb-str-count-still", (id, ch) =>
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
                return rd.Success && Has(Str(rd, "table"), "旧")
                    ? null
                    : "超限后原格变了 " + Str(rd, "table");
            });

            WithBook(run, s, "xlb-str-merge-still", (id, ch) =>
            {
                ToolResult r = s.Structure(id, s.FirstSheetName(ch), "merge", "A1:A201");
                if (r.Success)
                {
                    return "超限 merge 应失败";
                }

                return Has(r.Error, "过大") || Has(r.Error, "200") ? null : r.Error;
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
            if (err == null)
            {
                run.CaseOk(id);
            }
            else
            {
                run.CaseFail(id, err);
            }
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

        private static bool Has(string text, string part)
        {
            return !string.IsNullOrEmpty(text)
                && text.IndexOf(part, StringComparison.OrdinalIgnoreCase) >= 0;
        }
    }
}
