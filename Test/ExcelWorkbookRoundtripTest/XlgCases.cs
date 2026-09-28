using System;
using System.Collections;
using System.Collections.Generic;
using WordAddIn1;
using WordAddIn1.SpreadsheetHost;

namespace ExcelWorkbookRoundtripTest
{
    internal static class XlgCases
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
            ToolResult noCh = s.Chart(null, "create", Dest("利润表", "E2:K18"), SourceXml("利润表"));
            ExpectFail(run, "xlg-neg-no-channel", noCh, "打开");

            var req = new SpreadsheetChartRequest
            {
                Action = "create",
                DestSheet = "利润表",
                DestRangeA1 = "E2:K18"
            };
            if (!SpreadsheetHostAdapter.TryChart(new FakeWordChannel(), req, out _, out ToolResult wrong))
            {
                ExpectFail(run, "xlg-neg-wrong-kind", wrong, "excel");
            }
            else
            {
                run.CaseFail("xlg-neg-wrong-kind", "word 渠道应失败");
            }

            WithBook(run, s, "xlg-neg-no-action", (id, ch) =>
            {
                ToolResult r = s.Call("F_excel_chart", new Dictionary<string, object>
                {
                    ["channel_id"] = id
                });
                return r.Success ? "缺 action 应失败" : (Has(r.Error, "action") ? null : r.Error);
            });

            WithBook(run, s, "xlg-neg-bad-action", (id, ch) =>
            {
                ToolResult r = s.Chart(id, "refresh");
                return r.Success ? "非法 action 应失败" : (Has(r.Error, "action") ? null : r.Error);
            });

            WithBook(run, s, "xlg-neg-create-no-dest", (id, ch) =>
            {
                ToolResult r = s.Chart(id, "create", xml: SourceXml("Sheet1"));
                return r.Success ? "缺 dest 应失败" : (Has(r.Error, "dest") ? null : r.Error);
            });

            WithBook(run, s, "xlg-neg-create-no-xml", (id, ch) =>
            {
                string name = s.FirstSheetName(ch);
                ToolResult r = s.Chart(id, "create", Dest(name, "E2:K18"));
                return r.Success ? "缺 xml 应失败" : (Has(r.Error, "xml") ? null : r.Error);
            });

            WithBook(run, s, "xlg-neg-bang", (id, ch) =>
            {
                string name = s.FirstSheetName(ch);
                ToolResult r = s.Chart(id, "create", Dest(name, name + "!E2:K18"), SourceXml(name));
                return r.Success ? "含!应失败" : (Has(r.Error, "!") || Has(r.Error, "分开") ? null : r.Error);
            });

            WithBook(run, s, "xlg-neg-dest-cell", (id, ch) =>
            {
                string name = s.FirstSheetName(ch);
                ToolResult r = s.Chart(id, "create", Dest(name, "E2"), SourceXml(name));
                return r.Success ? "单格 dest 应失败" : (Has(r.Error, "2") || Has(r.Error, "矩形") ? null : r.Error);
            });

            WithBook(run, s, "xlg-neg-rows", (id, ch) =>
            {
                string name = s.FirstSheetName(ch);
                const string xml = "<ChartConfig><Chart><Type>column_clustered</Type>"
                    + "<DataTable><Rows><Row>1,2</Row></Rows></DataTable></Chart></ChartConfig>";
                ToolResult r = s.Chart(id, "create", Dest(name, "E2:K18"), xml);
                return r.Success ? "Rows 应失败" : (Has(r.Error, "Rows") || Has(r.Error, "禁止") ? null : r.Error);
            });

            WithBook(run, s, "xlg-neg-missing-sheet", (id, ch) =>
            {
                string name = s.FirstSheetName(ch);
                if (!s.SeedSales(ch, name))
                {
                    return "铺底失败";
                }

                ToolResult r = s.Chart(id, "create", Dest("不存在的表", "E2:K18"), SourceXml("不存在的表"));
                return r.Success
                    ? "缺表应失败"
                    : (Has(r.Error, "不存在") || Has(r.Error, "未找到") ? null : r.Error);
            });
        }

        private static void RunLife(TestRun run, XlSession s)
        {
            WithBook(run, s, "xlg-create-col", (id, ch) =>
            {
                string name = s.FirstSheetName(ch);
                if (!s.SeedSales(ch, name))
                {
                    return "铺底失败";
                }

                ToolResult r = CreateCol(s, id, name);
                if (!r.Success || string.IsNullOrEmpty(Str(r, "chart_name"))
                    || string.IsNullOrEmpty(Str(r, "dest_range")))
                {
                    return r.Success
                        ? "name/dest=" + Str(r, "chart_name") + "/" + Str(r, "dest_range")
                        : r.Error;
                }

                ToolResult g = s.Chart(id, "list");
                if (!g.Success || Int(g, "charts_count") < 1)
                {
                    return g.Error ?? "list 看不到";
                }

                ToolResult rd = s.Read(id, name, "A1:B4");
                return rd.Success && Has(Str(rd, "table"), "10") && Has(Str(rd, "table"), "销量")
                    ? null
                    : (rd.Error ?? "源被改了");
            });

            WithBook(run, s, "xlg-create-named", (id, ch) =>
            {
                string name = s.FirstSheetName(ch);
                if (!s.SeedSales(ch, name))
                {
                    return "铺底失败";
                }

                ToolResult r = CreateCol(s, id, name, "销量图");
                if (!r.Success || !string.Equals(Str(r, "chart_name"), "销量图", StringComparison.Ordinal))
                {
                    return r.Success ? "name=" + Str(r, "chart_name") : r.Error;
                }

                ToolResult g = s.Chart(id, "list");
                return g.Success && HasChartName(g, "销量图")
                    ? null
                    : (g.Error ?? "list 无销量图");
            });

            WithBook(run, s, "xlg-header-cols", (id, ch) =>
            {
                string name = s.FirstSheetName(ch);
                if (!s.SeedSales(ch, name))
                {
                    return "铺底失败";
                }

                ToolResult r = s.Chart(id, "create", Dest(name, "E2:K18"), HeaderXml(name), "分列图");
                if (!r.Success)
                {
                    return r.Error;
                }

                ToolResult ex = s.Chart(id, "extract", chartName: "分列图", scope: "data");
                string xml = Str(ex, "xml_content");
                return ex.Success && !Has(xml, "<Rows")
                    && (Has(xml, "<Source") || Has(xml, "<Column") || Has(xml, "<Header")
                        || Has(xml, "A2") || Has(xml, "B2"))
                    ? null
                    : (ex.Error ?? ("extract data 不对 " + xml));
            });

            WithBook(run, s, "xlg-list-book", (id, ch) =>
            {
                string name = s.FirstSheetName(ch);
                if (!s.SeedSales(ch, name))
                {
                    return "铺底失败";
                }

                ToolResult a = CreateCol(s, id, name, dest: "E2:K12");
                ToolResult b = CreateCol(s, id, name, dest: "M2:S12");
                if (!a.Success || !b.Success)
                {
                    return a.Error ?? b.Error;
                }

                ToolResult g = s.Chart(id, "list");
                return g.Success && Int(g, "charts_count") == 2
                    && (HasDest(g, "E2") || HasDest(g, "E2:K12"))
                    && (HasDest(g, "M2") || HasDest(g, "M2:S12"))
                    ? null
                    : (g.Error ?? "整簿 list 不是 2");
            });

            WithBook(run, s, "xlg-list-sheet", (id, ch) =>
            {
                string src = s.FirstSheetName(ch);
                if (!s.SeedTwoSheets(ch, src, "分析") || !s.SeedSales(ch, src) || !s.SeedSales(ch, "分析"))
                {
                    return "铺两表失败";
                }

                ToolResult a = CreateCol(s, id, src, dest: "E2:K18");
                ToolResult b = CreateCol(s, id, "分析", dest: "E2:K18");
                if (!a.Success || !b.Success)
                {
                    return a.Error ?? b.Error;
                }

                ToolResult g = s.Chart(id, "list", sheet: "分析");
                if (!g.Success || Int(g, "charts_count") != 1)
                {
                    return g.Error ?? ("count=" + Int(g, "charts_count"));
                }

                return string.Equals(Str(FirstChart(g), "sheet"), "分析", StringComparison.Ordinal)
                    ? null
                    : "sheet=" + Str(FirstChart(g), "sheet");
            });

            WithBook(run, s, "xlg-extract", (id, ch) =>
            {
                string name = s.FirstSheetName(ch);
                if (!s.SeedSales(ch, name))
                {
                    return "铺底失败";
                }

                ToolResult created = CreateCol(s, id, name, "销量图");
                if (!created.Success)
                {
                    return created.Error;
                }

                ToolResult ex = s.Chart(id, "extract", chartName: "销量图", scope: "all");
                string xml = Str(ex, "xml_content");
                if (!ex.Success || !Has(xml, "ChartConfig") || !Has(xml, "Type") || Has(xml, "<Rows"))
                {
                    return ex.Error ?? ("xml 不对 " + xml);
                }

                return Has(xml, "<Source") || Has(xml, "<Column") || Has(xml, "<Header")
                    ? null
                    : "没有单元格绑定 " + xml;
            });

            WithBook(run, s, "xlg-apply-format", (id, ch) =>
            {
                string name = s.FirstSheetName(ch);
                if (!s.SeedSales(ch, name))
                {
                    return "铺底失败";
                }

                ToolResult created = CreateCol(s, id, name, "销量图");
                if (!created.Success)
                {
                    return created.Error;
                }

                const string fmt = "<ChartConfig><Chart><Type>column_clustered</Type>"
                    + "<Title>改过的标题</Title></Chart></ChartConfig>";
                ToolResult ap = s.Chart(id, "apply", xml: fmt, chartName: "销量图", scope: "format");
                if (!ap.Success)
                {
                    return ap.Error;
                }

                ToolResult ex = s.Chart(id, "extract", chartName: "销量图", scope: "format");
                if (!ex.Success || !Has(Str(ex, "xml_content"), "改过的标题"))
                {
                    return ex.Error ?? ("Title 没套上 " + Str(ex, "xml_content"));
                }

                ToolResult rd = s.Read(id, name, "B2");
                return rd.Success && Has(Str(rd, "table"), "10")
                    ? null
                    : (rd.Error ?? "源被改了");
            });

            WithBook(run, s, "xlg-apply-data", (id, ch) =>
            {
                string name = s.FirstSheetName(ch);
                if (!s.SeedSales(ch, name)
                    || !s.SeedCell(ch, name, "A5", 2027)
                    || !s.SeedCell(ch, name, "B5", 40))
                {
                    return "铺底失败";
                }

                ToolResult created = CreateCol(s, id, name, "销量图");
                if (!created.Success)
                {
                    return created.Error;
                }

                ToolResult ap = s.Chart(
                    id,
                    "apply",
                    xml: SourceXml(name, "A1:B5"),
                    chartName: "销量图",
                    scope: "data");
                if (!ap.Success)
                {
                    return ap.Error;
                }

                object raw = s.ReadValue2(ch, name, "B5");
                return NearNumber(raw, 40)
                    ? null
                    : "B5 不是 write 的 40: " + raw;
            });

            WithBook(run, s, "xlg-write-follow", (id, ch) =>
            {
                string name = s.FirstSheetName(ch);
                if (!s.SeedSales(ch, name))
                {
                    return "铺底失败";
                }

                ToolResult created = CreateCol(s, id, name, "销量图");
                if (!created.Success)
                {
                    return created.Error;
                }

                string csv = s.WriteCsv("99");
                ToolResult w = s.Write(id, name, "B2", csv);
                if (!w.Success)
                {
                    return w.Error;
                }

                object raw = s.ReadValue2(ch, name, "B2");
                if (!NearNumber(raw, 99))
                {
                    return "源 B2 不是 99: " + raw;
                }

                ToolResult g = s.Chart(id, "list");
                Dictionary<string, object> one = FindName(g, "销量图");
                if (one == null)
                {
                    return g.Error ?? "图丢了";
                }

                string src = Str(one, "source_range");
                return string.IsNullOrEmpty(src)
                    || Has(src, "A1")
                    || Has(src, "A2")
                    || Has(src, "B2")
                    || Has(src, "B")
                    ? null
                    : "source_range=" + src;
            });

            WithBook(run, s, "xlg-delete", (id, ch) =>
            {
                string name = s.FirstSheetName(ch);
                if (!s.SeedSales(ch, name))
                {
                    return "铺底失败";
                }

                ToolResult created = CreateCol(s, id, name, "销量图");
                if (!created.Success)
                {
                    return created.Error;
                }

                ToolResult d = s.Chart(id, "delete", chartName: "销量图");
                if (!d.Success)
                {
                    return d.Error;
                }

                ToolResult g = s.Chart(id, "list");
                if (!g.Success || HasChartName(g, "销量图") || Int(g, "charts_count") != 0)
                {
                    return g.Error ?? "delete 后还在";
                }

                ToolResult rd = s.Read(id, name, "A1:B4");
                return rd.Success && Has(Str(rd, "table"), "10")
                    ? null
                    : (rd.Error ?? "源被删了");
            });
        }

        private static void RunGuard(TestRun run, XlSession s)
        {
            WithBook(run, s, "xlg-neg-conflict", (id, ch) =>
            {
                string name = s.FirstSheetName(ch);
                if (!s.SeedSales(ch, name))
                {
                    return "铺底失败";
                }

                ToolResult first = CreateCol(s, id, name);
                if (!first.Success)
                {
                    return first.Error;
                }

                ToolResult r = CreateCol(s, id, name, dest: "E10:K20");
                return r.Success ? "相交 dest 应失败" : (Has(r.Error, "相交") || Has(r.Error, "已有") ? null : r.Error);
            });

            WithBook(run, s, "xlg-neg-dup-name", (id, ch) =>
            {
                string name = s.FirstSheetName(ch);
                if (!s.SeedSales(ch, name))
                {
                    return "铺底失败";
                }

                ToolResult first = CreateCol(s, id, name, "销量图", "E2:K18");
                if (!first.Success)
                {
                    return first.Error;
                }

                ToolResult r = CreateCol(s, id, name, "销量图", "M2:S18");
                return r.Success ? "同名应失败" : (Has(r.Error, "同名") || Has(r.Error, "已存在") ? null : r.Error);
            });

            WithBook(run, s, "xlg-neg-dest-oversize", (id, ch) =>
            {
                string name = s.FirstSheetName(ch);
                ToolResult r = s.Chart(id, "create", Dest(name, "A1:A81"), SourceXml(name));
                return r.Success ? "dest 超限应失败" : (Has(r.Error, "过大") || Has(r.Error, "80") ? null : r.Error);
            });

            WithBook(run, s, "xlg-neg-source-oversize", (id, ch) =>
            {
                string name = s.FirstSheetName(ch);
                ToolResult r = s.Chart(id, "create", Dest(name, "E2:K18"), SourceXml(name, "A1:A5001"));
                return r.Success ? "源超限应失败" : (Has(r.Error, "过大") || Has(r.Error, "5000") ? null : r.Error);
            });

            WithBook(run, s, "xlg-neg-chart-sheet", (id, ch) =>
            {
                string name = s.FirstSheetName(ch);
                if (!s.SeedSales(ch, name) || !s.SeedChartSheet(ch, out string chartName))
                {
                    return "铺图表 sheet 失败";
                }

                ToolResult r = s.Chart(id, "create", Dest(chartName, "E2:K18"), SourceXml(name));
                return r.Success ? "图表 dest 应失败" : (Has(r.Error, "sheet_type=chart") ? null : r.Error);
            });

            WithBook(run, s, "xlg-neg-apply-no-scope", (id, ch) =>
            {
                string name = s.FirstSheetName(ch);
                ToolResult r = s.Call("F_excel_chart", new Dictionary<string, object>
                {
                    ["channel_id"] = id,
                    ["action"] = "apply",
                    ["chart_name"] = "没有",
                    ["xml_content"] = "<ChartConfig><Chart><Type>column_clustered</Type></Chart></ChartConfig>"
                });
                return r.Success ? "缺 scope 应失败" : (Has(r.Error, "scope") ? null : r.Error);
            });

            WithBook(run, s, "xlg-neg-source-header", (id, ch) =>
            {
                string name = s.FirstSheetName(ch);
                string xml = "<ChartConfig><Chart><Type>column_clustered</Type><DataTable>"
                    + "<Source sheet=\"" + name + "\" range=\"A1:B4\" plotBy=\"columns\"/>"
                    + "<Header><Column type=\"category\" sheet=\"" + name + "\" range=\"A2:A4\">年</Column>"
                    + "<Column type=\"value\" sheet=\"" + name + "\" range=\"B2:B4\">量</Column></Header>"
                    + "</DataTable></Chart></ChartConfig>";
                ToolResult r = s.Chart(id, "create", Dest(name, "E2:K18"), xml);
                return r.Success ? "Source+Header 应失败" : (Has(r.Error, "二选一") || Has(r.Error, "Source") ? null : r.Error);
            });

            WithBook(run, s, "xlg-neg-bad-type", (id, ch) =>
            {
                string name = s.FirstSheetName(ch);
                const string xml = "<ChartConfig><Chart><Type>bubble</Type>"
                    + "<DataTable><Source sheet=\"S\" range=\"A1:B4\" plotBy=\"columns\"/></DataTable>"
                    + "</Chart></ChartConfig>";
                ToolResult r = s.Chart(id, "create", Dest(name, "E2:K18"), xml);
                return r.Success ? "bubble 应失败" : (Has(r.Error, "未知") || Has(r.Error, "类型") ? null : r.Error);
            });

            WithBook(run, s, "xlg-neg-extract-missing", (id, ch) =>
            {
                ToolResult r = s.Chart(id, "extract", chartName: "没有", scope: "all");
                return r.Success ? "缺图 extract 应失败" : (Has(r.Error, "未找到") ? null : r.Error);
            });

            WithBook(run, s, "xlg-zh-hidden", (id, ch) =>
            {
                if (!s.SeedHidden(ch, "可见", "利润表") || !s.SeedSales(ch, "利润表"))
                {
                    return "铺隐藏表失败";
                }

                ToolResult r = CreateCol(s, id, "利润表", "隐图");
                if (!r.Success || !string.Equals(Str(r, "sheet"), "利润表", StringComparison.Ordinal))
                {
                    return r.Success ? "sheet=" + Str(r, "sheet") : r.Error;
                }

                ToolResult g = s.Chart(id, "list", sheet: "利润表");
                return g.Success && HasChartName(g, "隐图")
                    && string.Equals(Str(FindName(g, "隐图"), "sheet"), "利润表", StringComparison.Ordinal)
                    ? null
                    : (g.Error ?? "隐藏表图看不到");
            });
        }

        private static ToolResult CreateCol(
            XlSession s,
            string id,
            string sheet,
            string chartName = null,
            string dest = "E2:K18")
        {
            return s.Chart(id, "create", Dest(sheet, dest), SourceXml(sheet), chartName);
        }

        private static Dictionary<string, object> Dest(string sheet, string range)
        {
            return new Dictionary<string, object>
            {
                ["sheet"] = sheet,
                ["range"] = range
            };
        }

        private static string SourceXml(string sheet, string range = "A1:B4")
        {
            return "<ChartConfig><Chart><Type>column_clustered</Type><Title>销量</Title>"
                + "<DataTable><Source sheet=\"" + sheet + "\" range=\"" + range
                + "\" plotBy=\"columns\"/></DataTable></Chart></ChartConfig>";
        }

        private static string HeaderXml(string sheet)
        {
            return "<ChartConfig><Chart><Type>column_clustered</Type><Title>销量</Title><DataTable><Header>"
                + "<Column type=\"category\" sheet=\"" + sheet + "\" range=\"A2:A4\">年份</Column>"
                + "<Column type=\"value\" sheet=\"" + sheet + "\" range=\"B2:B4\">销量</Column>"
                + "</Header></DataTable></Chart></ChartConfig>";
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

        private static List<Dictionary<string, object>> Charts(ToolResult r)
        {
            var list = new List<Dictionary<string, object>>();
            Dictionary<string, object> data = Data(r);
            if (data == null || !data.ContainsKey("charts") || !(data["charts"] is IEnumerable raw))
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

        private static Dictionary<string, object> FirstChart(ToolResult r)
        {
            List<Dictionary<string, object>> list = Charts(r);
            return list.Count > 0 ? list[0] : null;
        }

        private static Dictionary<string, object> FindName(ToolResult r, string chartName)
        {
            foreach (Dictionary<string, object> one in Charts(r))
            {
                if (string.Equals(Str(one, "chart_name"), chartName, StringComparison.Ordinal))
                {
                    return one;
                }
            }

            return null;
        }

        private static bool HasChartName(ToolResult r, string chartName)
        {
            return FindName(r, chartName) != null;
        }

        private static bool HasDest(ToolResult r, string part)
        {
            foreach (Dictionary<string, object> one in Charts(r))
            {
                if (Has(Str(one, "dest_range"), part))
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
