using System;
using System.Collections;
using System.Collections.Generic;
using WordAddIn1;
using WordAddIn1.SpreadsheetHost;

namespace ExcelWorkbookRoundtripTest
{
    internal static class XlfCases
    {
        public static void Run(TestRun run, bool et)
        {
            var s = new XlSession(et);
            RunNeg(run, s);
            RunApply(run, s);
            RunGet(run, s);
        }

        private static void RunNeg(TestRun run, XlSession s)
        {
            ChannelRegistry.ClearAll();
            ToolResult noCh = s.ApplyFormat(null, "利润表", "A1", Bold());
            ExpectFail(run, "xlf-neg-no-channel", noCh, "打开");

            var req = new SpreadsheetFormatRequest
            {
                SheetName = "利润表",
                RangeA1 = "A1",
                Format = new SpreadsheetFormatPatch { Bold = true }
            };
            if (!SpreadsheetHostAdapter.TryApplyFormat(new FakeWordChannel(), req, out _, out ToolResult wrong))
            {
                ExpectFail(run, "xlf-neg-wrong-kind", wrong, "F_apply_document_format");
            }
            else
            {
                run.CaseFail("xlf-neg-wrong-kind", "word 渠道应失败");
            }

            WithBook(run, s, "xlf-neg-no-sheet", (id, ch) =>
            {
                ToolResult r = s.Call("F_apply_excel_format", new Dictionary<string, object>
                {
                    ["channel_id"] = id,
                    ["range"] = "A1",
                    ["format"] = Bold()
                });
                return r.Success ? "缺 sheet 应失败" : (Has(r.Error, "sheet") ? null : r.Error);
            });

            WithBook(run, s, "xlf-neg-no-range", (id, ch) =>
            {
                ToolResult r = s.Call("F_apply_excel_format", new Dictionary<string, object>
                {
                    ["channel_id"] = id,
                    ["sheet"] = s.FirstSheetName(ch),
                    ["format"] = Bold()
                });
                return r.Success ? "缺 range 应失败" : (Has(r.Error, "range") ? null : r.Error);
            });

            WithBook(run, s, "xlf-neg-no-format", (id, ch) =>
            {
                ToolResult r = s.Call("F_apply_excel_format", new Dictionary<string, object>
                {
                    ["channel_id"] = id,
                    ["sheet"] = s.FirstSheetName(ch),
                    ["range"] = "A1"
                });
                return r.Success ? "缺 format 应失败" : (Has(r.Error, "format") ? null : r.Error);
            });

            WithBook(run, s, "xlf-neg-empty-format", (id, ch) =>
            {
                ToolResult r = s.ApplyFormat(id, s.FirstSheetName(ch), "A1", new Dictionary<string, object>
                {
                    ["wrap_text"] = true
                });
                return r.Success ? "空有效字段应失败" : (Has(r.Error, "未指定") ? null : r.Error);
            });

            WithBook(run, s, "xlf-neg-bang", (id, ch) =>
            {
                string name = s.FirstSheetName(ch);
                ToolResult r = s.ApplyFormat(id, name, name + "!A1", Bold());
                return r.Success ? "含!应失败" : (Has(r.Error, "纯 A1") || Has(r.Error, "!") ? null : r.Error);
            });

            WithBook(run, s, "xlf-neg-missing-sheet", (id, ch) =>
            {
                ToolResult r = s.ApplyFormat(id, "不存在的表", "A1", Bold());
                return r.Success
                    ? "缺表应失败"
                    : (Has(r.Error, "不存在") || Has(r.Error, "未找到") ? null : r.Error);
            });

            WithBook(run, s, "xlf-neg-chart", (id, ch) =>
            {
                if (!s.SeedChartSheet(ch, out string chartName))
                {
                    return "铺图表 sheet 失败";
                }

                ToolResult r = s.ApplyFormat(id, chartName, "A1", Bold());
                return r.Success ? "图表 sheet 应失败" : (Has(r.Error, "sheet_type=chart") ? null : r.Error);
            });

            WithBook(run, s, "xlf-neg-apply-oversize", (id, ch) =>
            {
                ToolResult r = s.ApplyFormat(id, s.FirstSheetName(ch), "A1:A501", Bold());
                return r.Success ? "超限应失败" : (Has(r.Error, "过大") || Has(r.Error, "500") ? null : r.Error);
            });
        }

        private static void RunApply(TestRun run, XlSession s)
        {
            WithBook(run, s, "xlf-bold-keep", (id, ch) =>
            {
                string name = s.FirstSheetName(ch);
                if (!s.SeedCell(ch, name, "A1", "甲"))
                {
                    return "铺底失败";
                }

                ToolResult r = s.ApplyFormat(id, name, "A1", Bold());
                if (!r.Success || !AppliedHas(r, "bold"))
                {
                    return r.Success ? "applied_fields 无 bold" : r.Error;
                }

                ToolResult g = s.GetFormat(id, name, "A1");
                if (!g.Success)
                {
                    return g.Error;
                }

                if (!IsTrue(Fmt(ItemAt(g, 0)), "bold"))
                {
                    return "get 不是 bold";
                }

                ToolResult rd = s.Read(id, name, "A1");
                return rd.Success && Has(Str(rd, "table"), "甲")
                    ? null
                    : (rd.Error ?? "值丢了");
            });

            WithBook(run, s, "xlf-number", (id, ch) =>
            {
                string name = s.FirstSheetName(ch);
                if (!s.SeedCell(ch, name, "A1", 0.15))
                {
                    return "铺底失败";
                }

                ToolResult r = s.ApplyFormat(id, name, "A1", new Dictionary<string, object>
                {
                    ["number_format"] = "0.00%"
                });
                if (!r.Success)
                {
                    return r.Error;
                }

                ToolResult g = s.GetFormat(id, name, "A1");
                if (!g.Success || !Has(Str(Fmt(ItemAt(g, 0)), "number_format"), "%"))
                {
                    return g.Success ? "数字格式无 %" : g.Error;
                }

                object raw = s.ReadValue2(ch, name, "A1");
                double n;
                try
                {
                    n = Convert.ToDouble(raw);
                }
                catch
                {
                    return "Value2 不是数 " + raw;
                }

                return Math.Abs(n - 0.15) < 0.0001
                    ? null
                    : "值被改成 " + raw;
            });

            WithBook(run, s, "xlf-font", (id, ch) =>
            {
                string name = s.FirstSheetName(ch);
                ToolResult r = s.ApplyFormat(id, name, "A1", new Dictionary<string, object>
                {
                    ["font_name"] = "宋体",
                    ["font_size"] = 14
                });
                if (!r.Success)
                {
                    return r.Error;
                }

                Dictionary<string, object> fmt = Fmt(ItemAt(s.GetFormat(id, name, "A1"), 0));
                if (!Near(Num(fmt, "font_size"), 14, 0.01))
                {
                    return "字号 " + Num(fmt, "font_size");
                }

                return Has(Str(fmt, "font_name"), "宋体")
                    ? null
                    : "字体 " + Str(fmt, "font_name");
            });

            WithBook(run, s, "xlf-fill", (id, ch) =>
            {
                string name = s.FirstSheetName(ch);
                ToolResult r = s.ApplyFormat(id, name, "A1", new Dictionary<string, object>
                {
                    ["fill_color"] = "#FFF2CC"
                });
                if (!r.Success)
                {
                    return r.Error;
                }

                string fill = Str(Fmt(ItemAt(s.GetFormat(id, name, "A1"), 0)), "fill_color");
                return HexEq(fill, "FFF2CC") ? null : "fill=" + fill;
            });

            WithBook(run, s, "xlf-align", (id, ch) =>
            {
                string name = s.FirstSheetName(ch);
                ToolResult r = s.ApplyFormat(id, name, "A1", new Dictionary<string, object>
                {
                    ["horizontal_align"] = "center",
                    ["vertical_align"] = "center"
                });
                if (!r.Success)
                {
                    return r.Error;
                }

                Dictionary<string, object> fmt = Fmt(ItemAt(s.GetFormat(id, name, "A1"), 0));
                return StrEq(fmt, "horizontal_align", "center") && StrEq(fmt, "vertical_align", "center")
                    ? null
                    : "对齐 " + Str(fmt, "horizontal_align") + "/" + Str(fmt, "vertical_align");
            });

            WithBook(run, s, "xlf-border", (id, ch) =>
            {
                string name = s.FirstSheetName(ch);
                ToolResult r = s.ApplyFormat(id, name, "A1", new Dictionary<string, object>
                {
                    ["border"] = new Dictionary<string, object>
                    {
                        ["bottom"] = new Dictionary<string, object>
                        {
                            ["style"] = "thin",
                            ["color"] = "#000000"
                        }
                    }
                });
                if (!r.Success)
                {
                    return r.Error;
                }

                Dictionary<string, object> fmt = Fmt(ItemAt(s.GetFormat(id, name, "A1"), 0));
                if (!fmt.ContainsKey("border") || !(fmt["border"] is Dictionary<string, object> border))
                {
                    return "无 border";
                }

                foreach (string edge in new[] { "top", "bottom", "left", "right" })
                {
                    if (!border.ContainsKey(edge))
                    {
                        return "缺边 " + edge;
                    }
                }

                Dictionary<string, object> bottom = border["bottom"] as Dictionary<string, object>;
                string style = bottom == null ? "" : Convert.ToString(bottom.ContainsKey("style") ? bottom["style"] : "");
                return string.Equals(style, "thin", StringComparison.OrdinalIgnoreCase)
                    ? null
                    : "底边 " + style;
            });

            WithBook(run, s, "xlf-col-width", (id, ch) =>
            {
                string name = s.FirstSheetName(ch);
                ToolResult r = s.ApplyFormat(id, name, "A1:D1", new Dictionary<string, object>
                {
                    ["column_width"] = 14
                });
                if (!r.Success)
                {
                    return r.Error;
                }

                Dictionary<string, object> fmt = Fmt(ItemAt(s.GetFormat(id, name, "A1"), 0));
                return Near(Num(fmt, "column_width"), 14, 0.3)
                    ? null
                    : "列宽 " + Num(fmt, "column_width");
            });

            WithBook(run, s, "xlf-row-height", (id, ch) =>
            {
                string name = s.FirstSheetName(ch);
                ToolResult r = s.ApplyFormat(id, name, "A1:A3", new Dictionary<string, object>
                {
                    ["row_height"] = 22
                });
                if (!r.Success)
                {
                    return r.Error;
                }

                Dictionary<string, object> fmt = Fmt(ItemAt(s.GetFormat(id, name, "A1"), 0));
                return Near(Num(fmt, "row_height"), 22, 0.3)
                    ? null
                    : "行高 " + Num(fmt, "row_height");
            });

            WithBook(run, s, "xlf-partial", (id, ch) =>
            {
                string name = s.FirstSheetName(ch);
                ToolResult first = s.ApplyFormat(id, name, "A1", Bold());
                if (!first.Success)
                {
                    return first.Error;
                }

                ToolResult second = s.ApplyFormat(id, name, "A1", new Dictionary<string, object>
                {
                    ["fill_color"] = "#FFF2CC"
                });
                if (!second.Success)
                {
                    return second.Error;
                }

                if (AppliedHas(second, "bold") || !AppliedHas(second, "fill_color"))
                {
                    return "第二次 applied_fields 不对";
                }

                Dictionary<string, object> fmt = Fmt(ItemAt(s.GetFormat(id, name, "A1"), 0));
                return IsTrue(fmt, "bold") && HexEq(Str(fmt, "fill_color"), "FFF2CC")
                    ? null
                    : "partial 丢了字段";
            });

            WithBook(run, s, "xlf-hidden", (id, ch) =>
            {
                if (!s.SeedHidden(ch, "可见表", "隐藏底稿"))
                {
                    return "铺隐藏表失败";
                }

                ToolResult r = s.ApplyFormat(id, "隐藏底稿", "A1", Bold());
                if (!r.Success)
                {
                    return r.Error;
                }

                ToolResult g = s.GetFormat(id, "隐藏底稿", "A1");
                return g.Success && IsTrue(Fmt(ItemAt(g, 0)), "bold")
                    ? null
                    : (g.Error ?? "隐藏表没套上");
            });
        }

        private static void RunGet(TestRun run, XlSession s)
        {
            WithBook(run, s, "xlf-get-neg-no-range", (id, ch) =>
            {
                ToolResult r = s.Call("F_get_excel_format", new Dictionary<string, object>
                {
                    ["channel_id"] = id,
                    ["sheet"] = s.FirstSheetName(ch)
                });
                return r.Success ? "缺 range 应失败" : (Has(r.Error, "range") ? null : r.Error);
            });

            WithBook(run, s, "xlf-get-oversize", (id, ch) =>
            {
                ToolResult r = s.GetFormat(id, s.FirstSheetName(ch), "A1:A101");
                return r.Success ? "get 超限应失败" : (Has(r.Error, "过大") || Has(r.Error, "100") ? null : r.Error);
            });

            WithBook(run, s, "xlf-get-cells", (id, ch) =>
            {
                ToolResult r = s.GetFormat(id, s.FirstSheetName(ch), "A1:B2");
                if (!r.Success)
                {
                    return r.Error;
                }

                if (!string.Equals(Str(r, "mode"), "cells", StringComparison.Ordinal))
                {
                    return "mode=" + Str(r, "mode");
                }

                List<Dictionary<string, object>> items = Items(r);
                if (items.Count != 4)
                {
                    return "items=" + items.Count;
                }

                foreach (Dictionary<string, object> one in items)
                {
                    if (string.IsNullOrEmpty(Str(one, "addr")))
                    {
                        return "缺 addr";
                    }
                }

                return null;
            });

            WithBook(run, s, "xlf-get-column", (id, ch) =>
            {
                ToolResult r = s.GetFormat(id, s.FirstSheetName(ch), "A1:A50");
                if (!r.Success)
                {
                    return r.Error;
                }

                if (!string.Equals(Str(r, "mode"), "column", StringComparison.Ordinal))
                {
                    return "mode=" + Str(r, "mode");
                }

                Dictionary<string, object> one = ItemAt(r, 0);
                if (one == null || !string.Equals(Str(one, "col"), "A", StringComparison.Ordinal)
                    || string.IsNullOrEmpty(Str(one, "range")))
                {
                    return "缺 col/range";
                }

                Dictionary<string, object> fmt = Fmt(one);
                return !fmt.ContainsKey("row_height")
                    ? null
                    : "列汇总不该带 row_height";
            });

            WithBook(run, s, "xlf-get-columns", (id, ch) =>
            {
                ToolResult r = s.GetFormat(id, s.FirstSheetName(ch), "A1:C20");
                if (!r.Success)
                {
                    return r.Error;
                }

                if (!string.Equals(Str(r, "mode"), "columns", StringComparison.Ordinal))
                {
                    return "mode=" + Str(r, "mode");
                }

                List<Dictionary<string, object>> items = Items(r);
                if (items.Count != 3)
                {
                    return "items=" + items.Count;
                }

                return !string.IsNullOrEmpty(Str(items[0], "col")) && !string.IsNullOrEmpty(Str(items[1], "col"))
                    ? null
                    : "缺 col";
            });

            WithBook(run, s, "xlf-get-mixed", (id, ch) =>
            {
                string name = s.FirstSheetName(ch);
                if (!s.SeedCell(ch, name, "A1", "上") || !s.SeedCell(ch, name, "A2", "下"))
                {
                    return "铺底失败";
                }

                ToolResult ap = s.ApplyFormat(id, name, "A1", Bold());
                if (!ap.Success)
                {
                    return ap.Error;
                }

                ToolResult g = s.GetFormat(id, name, "A1:A2");
                if (!g.Success)
                {
                    return g.Error;
                }

                if (!string.Equals(Str(g, "mode"), "cells", StringComparison.Ordinal) || Items(g).Count != 2)
                {
                    return "mode/items " + Str(g, "mode") + "/" + Items(g).Count;
                }

                Dictionary<string, object> a1 = ItemNamed(g, "A1");
                Dictionary<string, object> a2 = ItemNamed(g, "A2");
                if (a1 == null || a2 == null)
                {
                    return "缺 A1/A2";
                }

                return IsTrue(Fmt(a1), "bold") && !IsTrue(Fmt(a2), "bold")
                    ? null
                    : "两格皮相同";
            });

            WithBook(run, s, "xlf-roundtrip", (id, ch) =>
            {
                string name = s.FirstSheetName(ch);
                if (!s.SeedCell(ch, name, "A1", 12.5) || !s.SeedCell(ch, name, "B1", 3))
                {
                    return "铺底失败";
                }

                ToolResult ap = s.ApplyFormat(id, name, "A1", new Dictionary<string, object>
                {
                    ["number_format"] = "0.00"
                });
                if (!ap.Success)
                {
                    return ap.Error;
                }

                ToolResult g = s.GetFormat(id, name, "A1");
                if (!g.Success)
                {
                    return g.Error;
                }

                Dictionary<string, object> copied = Fmt(ItemAt(g, 0));
                if (copied.Count == 0)
                {
                    return "get format 空";
                }

                ToolResult again = s.ApplyFormat(id, name, "B1", copied);
                if (!again.Success)
                {
                    return again.Error;
                }

                string a = Str(Fmt(ItemAt(s.GetFormat(id, name, "A1"), 0)), "number_format");
                string b = Str(Fmt(ItemAt(s.GetFormat(id, name, "B1"), 0)), "number_format");
                return !string.IsNullOrEmpty(a) && string.Equals(a, b, StringComparison.Ordinal)
                    ? null
                    : "回抄数字格式 " + a + " / " + b;
            });

            WithBook(run, s, "xlf-get-chart", (id, ch) =>
            {
                if (!s.SeedChartSheet(ch, out string chartName))
                {
                    return "铺图表 sheet 失败";
                }

                ToolResult r = s.GetFormat(id, chartName, "A1");
                return r.Success ? "图表 get 应失败" : (Has(r.Error, "sheet_type=chart") ? null : r.Error);
            });

            WithBook(run, s, "xlf-zh-sheet", (id, ch) =>
            {
                if (!s.SeedUsed(ch, "利润表"))
                {
                    return "铺底失败";
                }

                ToolResult ap = s.ApplyFormat(id, "利润表", "A1", Bold());
                if (!ap.Success || !string.Equals(Str(ap, "sheet"), "利润表", StringComparison.Ordinal))
                {
                    return ap.Success ? "apply sheet=" + Str(ap, "sheet") : ap.Error;
                }

                ToolResult g = s.GetFormat(id, "利润表", "A1");
                return g.Success && string.Equals(Str(g, "sheet"), "利润表", StringComparison.Ordinal)
                    ? null
                    : (g.Success ? "get sheet=" + Str(g, "sheet") : g.Error);
            });

            WithBook(run, s, "xlf-neg-border-all", (id, ch) =>
            {
                ToolResult r = s.ApplyFormat(id, s.FirstSheetName(ch), "A1", new Dictionary<string, object>
                {
                    ["border"] = "all"
                });
                return r.Success ? "border=all 应失败" : (Has(r.Error, "对象") ? null : r.Error);
            });
        }

        private static Dictionary<string, object> Bold()
        {
            return new Dictionary<string, object> { ["bold"] = true };
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

        private static bool AppliedHas(ToolResult r, string field)
        {
            Dictionary<string, object> data = Data(r);
            if (data == null || !data.ContainsKey("applied_fields") || !(data["applied_fields"] is IEnumerable raw))
            {
                return false;
            }

            foreach (object one in raw)
            {
                if (string.Equals(Convert.ToString(one), field, StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
            }

            return false;
        }

        private static List<Dictionary<string, object>> Items(ToolResult r)
        {
            var list = new List<Dictionary<string, object>>();
            Dictionary<string, object> data = Data(r);
            if (data == null || !data.ContainsKey("items") || !(data["items"] is IEnumerable raw))
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

        private static Dictionary<string, object> ItemAt(ToolResult r, int index)
        {
            List<Dictionary<string, object>> items = Items(r);
            return index >= 0 && index < items.Count ? items[index] : null;
        }

        private static Dictionary<string, object> ItemNamed(ToolResult r, string addr)
        {
            foreach (Dictionary<string, object> one in Items(r))
            {
                if (string.Equals(Str(one, "addr"), addr, StringComparison.Ordinal))
                {
                    return one;
                }
            }

            return null;
        }

        private static Dictionary<string, object> Fmt(Dictionary<string, object> item)
        {
            if (item == null || !item.ContainsKey("format"))
            {
                return new Dictionary<string, object>();
            }

            return item["format"] as Dictionary<string, object> ?? new Dictionary<string, object>();
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

        private static bool StrEq(Dictionary<string, object> d, string key, string expect)
        {
            return string.Equals(Str(d, key), expect, StringComparison.OrdinalIgnoreCase);
        }

        private static double Num(Dictionary<string, object> d, string key)
        {
            if (d == null || !d.ContainsKey(key) || d[key] == null)
            {
                return 0;
            }

            try
            {
                return Convert.ToDouble(d[key]);
            }
            catch
            {
                return 0;
            }
        }

        private static bool Near(double a, double b, double tol)
        {
            return Math.Abs(a - b) <= tol;
        }

        private static bool HexEq(string raw, string rgb)
        {
            string a = (raw ?? "").Replace("#", "").Trim();
            return string.Equals(a, rgb, StringComparison.OrdinalIgnoreCase);
        }
    }
}
