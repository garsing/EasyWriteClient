using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using WordAddIn1;
using WordAddIn1.OpenFiles;
using WordAddIn1.SpreadsheetHost;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelWorkbookRoundtripTest
{
    internal static class XiCases
    {
        private static readonly Regex ImageIdRx = new Regex(@"^I_[A-Za-z0-9]{8}$", RegexOptions.Compiled);

        private const string TinyPngB64 =
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==";

        public static void Run(TestRun run, bool et)
        {
            var s = new XlSession(et);
            RunNeg(run, s);
            RunLife(run, s);
        }

        private static void RunNeg(TestRun run, XlSession s)
        {
            ChannelRegistry.ClearAll();
            ToolResult noCh = s.Image(null, "insert", "Sheet1", "B2", WriteTinyPng());
            ExpectFail(run, "xi-neg-no-channel", noCh, "打开");

            var req = new SpreadsheetImageRequest
            {
                Action = "insert",
                Sheet = "Sheet1",
                Anchor = "B2",
                LocalPath = WriteTinyPng()
            };
            if (!SpreadsheetHostAdapter.TryImage(new FakeWordChannel(), req, out _, out ToolResult wrong))
            {
                ExpectFail(run, "xi-neg-wrong-kind", wrong, "excel");
            }
            else
            {
                run.CaseFail("xi-neg-wrong-kind", "word 渠道应失败");
            }

            ExpectFail(
                run,
                "xi-neg-svg-none",
                s.Call("F_render_svg", new Dictionary<string, object>()),
                "二选一");
            ExpectFail(
                run,
                "xi-neg-svg-both",
                s.Call("F_render_svg", new Dictionary<string, object>
                {
                    ["svg"] = "<svg xmlns=\"http://www.w3.org/2000/svg\"></svg>",
                    ["path"] = "a.svg"
                }),
                "二选一");

            if (!DocumentCheckpointMutatingTools.ShouldCheckpoint(
                    "F_excel_image",
                    new Dictionary<string, object> { ["action"] = "list" })
                && DocumentCheckpointMutatingTools.ShouldCheckpoint(
                    "F_excel_image",
                    new Dictionary<string, object> { ["action"] = "insert" })
                && !DocumentCheckpointMutatingTools.ShouldCheckpoint("F_render_svg", null))
            {
                run.CaseOk("xi-ckpt-action");
            }
            else
            {
                run.CaseFail("xi-ckpt-action", "list/render 不应建点，insert 应建");
            }

            WithBook(run, s, "xi-neg-no-action", (id, ch) =>
            {
                ToolResult r = s.Call("F_excel_image", new Dictionary<string, object>
                {
                    ["channel_id"] = id
                });
                return r.Success ? "缺 action 应失败" : (Has(r.Error, "action") ? null : r.Error);
            });

            WithBook(run, s, "xi-neg-bad-action", (id, ch) =>
            {
                ToolResult r = s.Image(id, "refresh");
                return r.Success ? "非法 action 应失败" : (Has(r.Error, "action") ? null : r.Error);
            });

            WithBook(run, s, "xi-neg-insert-no-sheet", (id, ch) =>
            {
                ToolResult r = s.Image(id, "insert", anchor: "B2", path: WriteTinyPng());
                return r.Success ? "缺 sheet 应失败" : (Has(r.Error, "sheet") ? null : r.Error);
            });

            WithBook(run, s, "xi-neg-insert-no-anchor", (id, ch) =>
            {
                string name = s.FirstSheetName(ch);
                ToolResult r = s.Image(id, "insert", name, path: WriteTinyPng());
                return r.Success ? "缺 anchor 应失败" : (Has(r.Error, "anchor") ? null : r.Error);
            });

            WithBook(run, s, "xi-neg-insert-range", (id, ch) =>
            {
                string name = s.FirstSheetName(ch);
                ToolResult r = s.Image(id, "insert", name, "B2:C3", WriteTinyPng());
                return r.Success ? "区域 anchor 应失败" : (Has(r.Error, "A1") || Has(r.Error, "anchor") ? null : r.Error);
            });

            WithBook(run, s, "xi-neg-insert-missing", (id, ch) =>
            {
                string name = s.FirstSheetName(ch);
                string missing = Path.Combine(Path.GetTempPath(), "ew-xi-missing-" + Guid.NewGuid().ToString("N") + ".png");
                ToolResult r = s.Image(id, "insert", name, "B2", missing);
                return r.Success ? "缺文件应失败" : (Has(r.Error, "不存在") || Has(r.Error, "文件") ? null : r.Error);
            });

            WithBook(run, s, "xi-neg-insert-svg", (id, ch) =>
            {
                string name = s.FirstSheetName(ch);
                string svg = Path.Combine(Path.GetTempPath(), "ew-xi-" + Guid.NewGuid().ToString("N") + ".svg");
                File.WriteAllText(svg, "<svg xmlns=\"http://www.w3.org/2000/svg\"></svg>");
                ToolResult r = s.Image(id, "insert", name, "B2", svg);
                return r.Success ? "svg 应失败" : (Has(r.Error, "png") || Has(r.Error, "后缀") ? null : r.Error);
            });

            WithBook(run, s, "xi-neg-update-bad-id", (id, ch) =>
            {
                ToolResult r = s.Image(id, "update", imageId: "I_xxxxxxxx", widthCm: 10);
                return r.Success ? "假码应失败" : (Has(r.Error, "未找到") || Has(r.Error, "image_id") ? null : r.Error);
            });
        }

        private static void RunLife(TestRun run, XlSession s)
        {
            WithBook(run, s, "xi-insert-list", (id, ch) =>
            {
                string name = s.FirstSheetName(ch);
                string png = WriteTinyPng();
                ToolResult ins = s.Image(id, "insert", name, "B2", png);
                string imageId = Str(ins, "image_id");
                if (!ins.Success || !ImageIdRx.IsMatch(imageId))
                {
                    return ins.Success ? "image_id=" + imageId : ins.Error;
                }

                if (!NearNumber(Data(ins) != null && Data(ins).ContainsKey("width_cm") ? Data(ins)["width_cm"] : null, 8))
                {
                    return "默认宽=" + Str(ins, "width_cm");
                }

                ToolResult listed = s.Image(id, "list");
                if (!listed.Success || Int(listed, "images_count") != 1)
                {
                    return listed.Error ?? "list count=" + Int(listed, "images_count");
                }

                Dictionary<string, object> one = FirstImage(listed);
                return one != null && string.Equals(Str(one, "image_id"), imageId, StringComparison.Ordinal)
                    && Has(Str(one, "anchor"), "B2")
                    ? null
                    : "list 对不上";
            });

            WithBook(run, s, "xi-list-stamp", (id, ch) =>
            {
                string name = s.FirstSheetName(ch);
                ToolResult ins = s.Image(id, "insert", name, "B2", WriteTinyPng());
                if (!ins.Success)
                {
                    return ins.Error;
                }

                if (!ClearPictureAlts(ch))
                {
                    return "清戳失败";
                }

                ToolResult listed = s.Image(id, "list");
                string stamped = listed.Success ? Str(FirstImage(listed), "image_id") : "";
                if (!listed.Success || !ImageIdRx.IsMatch(stamped))
                {
                    return listed.Error ?? "补戳失败 " + stamped;
                }

                ToolResult again = s.Image(id, "list");
                return again.Success && string.Equals(Str(FirstImage(again), "image_id"), stamped, StringComparison.Ordinal)
                    ? null
                    : (again.Error ?? "二次 list 码变了");
            });

            WithBook(run, s, "xi-update-width-move", (id, ch) =>
            {
                string name = s.FirstSheetName(ch);
                ToolResult ins = s.Image(id, "insert", name, "B2", WriteTinyPng());
                string imageId = Str(ins, "image_id");
                if (!ins.Success)
                {
                    return ins.Error;
                }

                ToolResult wide = s.Image(id, "update", imageId: imageId, widthCm: 10);
                if (!wide.Success || !string.Equals(Str(wide, "image_id"), imageId, StringComparison.Ordinal)
                    || !NearNumber(Data(wide) != null && Data(wide).ContainsKey("width_cm") ? Data(wide)["width_cm"] : null, 10))
                {
                    return wide.Success ? "宽=" + Str(wide, "width_cm") : wide.Error;
                }

                ToolResult moved = s.Image(id, "update", imageId: imageId, anchor: "D4");
                if (!moved.Success || !Has(Str(moved, "anchor"), "D4")
                    || !string.Equals(Str(moved, "image_id"), imageId, StringComparison.Ordinal))
                {
                    return moved.Success ? "anchor=" + Str(moved, "anchor") : moved.Error;
                }

                ToolResult swapped = s.Image(id, "update", imageId: imageId, path: WriteTinyPng());
                return swapped.Success && string.Equals(Str(swapped, "image_id"), imageId, StringComparison.Ordinal)
                    ? null
                    : (swapped.Error ?? "换图码变了");
            });

            WithBook(run, s, "xi-delete", (id, ch) =>
            {
                string name = s.FirstSheetName(ch);
                ToolResult ins = s.Image(id, "insert", name, "B2", WriteTinyPng());
                string imageId = Str(ins, "image_id");
                if (!ins.Success)
                {
                    return ins.Error;
                }

                ToolResult del = s.Image(id, "delete", imageId: imageId);
                if (!del.Success)
                {
                    return del.Error;
                }

                ToolResult listed = s.Image(id, "list");
                return listed.Success && Int(listed, "images_count") == 0
                    ? null
                    : (listed.Error ?? "删后仍有图");
            });

            WithBook(run, s, "xi-neg-update-no-fields", (id, ch) =>
            {
                string name = s.FirstSheetName(ch);
                ToolResult ins = s.Image(id, "insert", name, "B2", WriteTinyPng());
                string imageId = Str(ins, "image_id");
                if (!ins.Success)
                {
                    return ins.Error;
                }

                ToolResult r = s.Image(id, "update", imageId: imageId);
                return r.Success ? "空 update 应失败" : (Has(r.Error, "update") || Has(r.Error, "至少") ? null : r.Error);
            });

            WithBook(run, s, "xi-render-svg", (id, ch) =>
            {
                ToolResult render = s.Call("F_render_svg", new Dictionary<string, object>
                {
                    ["svg"] = "<svg xmlns=\"http://www.w3.org/2000/svg\" viewBox=\"0 0 10 10\"><rect width=\"10\" height=\"10\" fill=\"#09c\"/></svg>"
                });
                if (!render.Success)
                {
                    return Has(render.Error, "工作区") || Has(render.Error, "登录")
                        ? "SKIP:" + render.Error
                        : render.Error;
                }

                string filename = Str(render, "filename");
                if (string.IsNullOrEmpty(filename) || !filename.EndsWith(".png", StringComparison.OrdinalIgnoreCase))
                {
                    return "filename=" + filename;
                }

                string name = s.FirstSheetName(ch);
                ToolResult ins = s.Image(id, "insert", name, "B2", filename);
                return ins.Success && ImageIdRx.IsMatch(Str(ins, "image_id"))
                    ? null
                    : (ins.Error ?? "render 后 insert 失败");
            });
        }

        private static string WriteTinyPng()
        {
            string path = Path.Combine(Path.GetTempPath(), "ew-xi-" + Guid.NewGuid().ToString("N") + ".png");
            File.WriteAllBytes(path, Convert.FromBase64String(TinyPngB64));
            return path;
        }

        private static bool ClearPictureAlts(IOperationChannel ch)
        {
            try
            {
                if (ch is ExcelChannel excel && excel.TryGetLiveWorkbook(out Excel.Workbook book))
                {
                    foreach (Excel.Worksheet ws in book.Worksheets)
                    {
                        foreach (Excel.Shape sh in ws.Shapes)
                        {
                            int t = Convert.ToInt32(sh.Type);
                            if (t == 13 || t == 11)
                            {
                                sh.AlternativeText = "";
                            }
                        }
                    }

                    return true;
                }

                if (ch is EtChannel et && et.TryGetLiveWorkbook(out object etBook))
                {
                    foreach (object sheet in EtCom.EnumerateSheets(etBook))
                    {
                        object shapes = EtCom.GetProperty(sheet, "Shapes");
                        if (shapes == null)
                        {
                            continue;
                        }

                        int count = Convert.ToInt32(EtCom.GetProperty(shapes, "Count"));
                        for (int i = 1; i <= count; i++)
                        {
                            object shape = EtCom.GetIndexed(shapes, i);
                            EtCom.TrySetProperty(shape, "AlternativeText", "");
                        }
                    }

                    return true;
                }
            }
            catch
            {
                return false;
            }

            return false;
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

        private static List<Dictionary<string, object>> Images(ToolResult r)
        {
            var list = new List<Dictionary<string, object>>();
            Dictionary<string, object> data = Data(r);
            if (data == null || !data.ContainsKey("images") || !(data["images"] is IEnumerable raw))
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

        private static Dictionary<string, object> FirstImage(ToolResult r)
        {
            List<Dictionary<string, object>> list = Images(r);
            return list.Count > 0 ? list[0] : null;
        }

        private static bool NearNumber(object raw, double expect)
        {
            try
            {
                return Math.Abs(Convert.ToDouble(raw) - expect) < 0.15;
            }
            catch
            {
                return false;
            }
        }
    }
}
