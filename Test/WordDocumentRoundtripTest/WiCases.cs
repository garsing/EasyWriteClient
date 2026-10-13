using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using Newtonsoft.Json;
using WordAddIn1;

namespace WordDocumentRoundtripTest
{
    /// <summary>
    /// 第 8 波：插图（位图 / SVG）+ 页截图。
    /// Word 与 WPS 同一套 ID，不 SKIP（无默认渠条现场被占时记环境 SKIP）。
    /// 用例清单见 docs/测试/word测试/插图与页图/用例清单.md。
    /// </summary>
    internal static class WiCases
    {
        public const int Target = 38;

        private const string BeforeSent = "易写图前句。";
        private const string AnchorSent = "易写句后锚。";

        private static readonly string[] BitmapIds =
        {
            "wi-ins-ok", "wi-ins-fields", "wi-ins-id", "wi-ins-text", "wi-ins-width"
        };

        private static readonly string[] KeepSentIds = { "wi-ins-keep-sent" };

        private static readonly string[] AnchorIds = { "wi-ins-after-codes" };

        private static readonly string[] StartIds = { "wi-ins-start" };

        private static readonly string[] SvgTextIds = { "wi-ins-svg-text" };

        private static readonly string[] SvgFileIds = { "wi-ins-svg-file" };

        public static void Run(
            TestRun run,
            string host,
            int wordBefore,
            int wpsBefore,
            bool wordWinBefore,
            bool wpsWinBefore)
        {
            var s = new WdSession(host);
            string png = FixturePng.Ensure();
            string svgPath = WriteSvgFile();
            string svgText = LegalSvg();

            RunImageNeg(run, s, png, svgPath, svgText);

            RunBitmapInsert(run, s, png);
            RunKeepSent(run, s, png);
            RunAnchorInsert(run, s, png);
            RunStartInsert(run, s, png);
            RunSvgTextInsert(run, s, svgText);
            RunSvgFileInsert(run, s, svgPath);

            RunCaptureNeg(run, s);
            RunCapturePos(run, s);

            RunNoQuit(run, wordBefore, wpsBefore, wordWinBefore, wpsWinBefore);
        }

        // ---------- A. 插图负例 ----------

        private static void RunImageNeg(TestRun run, WdSession s, string png, string svgPath, string svgText)
        {
            ToolResult noCh = s.InsertImage(new Dictionary<string, object> { ["path"] = png });
            if (!noCh.Success
                && (Contains(noCh.Error, "打开")
                    || Contains(noCh.Error, "unsupported")
                    || Contains(noCh.Error, "browser")
                    || Contains(noCh.Error, "渠道")))
            {
                run.CaseOk("wi-neg-no-channel");
            }
            else if (noCh.Success)
            {
                run.CaseSkip("wi-neg-no-channel", "现场已有默认文字渠，不关用户稿");
            }
            else
            {
                run.CaseFail("wi-neg-no-channel", noCh.Error ?? "应失败或跳过");
            }

            ExpectFail(run, "wi-neg-unknown-id",
                s.InsertImage(new Dictionary<string, object> { ["channel_id"] = "w999", ["path"] = png }),
                "未知");
            ExpectFail(run, "wi-neg-svg-both",
                s.InsertSvgImage(new Dictionary<string, object>
                {
                    ["path"] = svgPath,
                    ["svg"] = svgText
                }),
                "只能传");
            ExpectFail(run, "wi-neg-svg-none",
                s.InsertSvgImage(new Dictionary<string, object>()),
                "必须提供");
            string[] needChannel =
            {
                "wi-neg-no-path", "wi-neg-missing", "wi-neg-bad-ext",
                "wi-neg-target-two", "wi-neg-target-intable-only",
                "wi-neg-target-pos-end", "wi-neg-target-pos-num",
                "wi-neg-target-unknown-code", "wi-neg-target-unknown-tbl",
                "wi-neg-target-unknown-chart", "wi-neg-target-unknown-img"
            };
            string path, id;
            if (!TryOpen(run, s, out path, out id, needChannel))
            {
                return;
            }

            ExpectFail(run, "wi-neg-no-path",
                s.InsertImage(new Dictionary<string, object> { ["channel_id"] = id }),
                "path");
            ExpectFail(run, "wi-neg-missing",
                s.InsertImage(new Dictionary<string, object>
                {
                    ["channel_id"] = id,
                    ["path"] = Path.Combine(Path.GetTempPath(), "ew-wi-missing-" + Guid.NewGuid().ToString("N") + ".png")
                }),
                "不存在");
            ExpectFail(run, "wi-neg-bad-ext",
                s.InsertImage(new Dictionary<string, object>
                {
                    ["channel_id"] = id,
                    ["path"] = WdSession.DummyPath(".txt")
                }),
                "unsupported_format");
            ExpectFail(run, "wi-neg-target-two",
                s.InsertImage(TargetArgs(id, png, new Dictionary<string, object>
                {
                    ["codes"] = "S_00001",
                    ["table"] = "T_00001"
                })),
                "定位对象");
            ExpectFail(run, "wi-neg-target-intable-only",
                s.InsertImage(TargetArgs(id, png, new Dictionary<string, object>
                {
                    ["in_table"] = "T_00001"
                })),
                "定位对象");
            ExpectFail(run, "wi-neg-target-pos-end",
                s.InsertImage(TargetArgs(id, png, new Dictionary<string, object> { ["position"] = "end" })),
                "position");
            ExpectFail(run, "wi-neg-target-pos-num",
                s.InsertImage(TargetArgs(id, png, new Dictionary<string, object> { ["position"] = "3" })),
                "position");
            ExpectFail(run, "wi-neg-target-unknown-code",
                s.InsertImage(TargetArgs(id, png, new Dictionary<string, object> { ["codes"] = "S_99999" })),
                "无法定位");
            ExpectFail(run, "wi-neg-target-unknown-tbl",
                s.InsertImage(TargetArgs(id, png, new Dictionary<string, object> { ["table"] = "T_99999" })),
                "找不到");
            ExpectFail(run, "wi-neg-target-unknown-chart",
                s.InsertImage(TargetArgs(id, png, new Dictionary<string, object> { ["chart"] = "C_99999" })),
                "找不到");
            ExpectFail(run, "wi-neg-target-unknown-img",
                s.InsertImage(TargetArgs(id, png, new Dictionary<string, object> { ["image"] = "I_99999" })),
                "找不到");
            s.CloseDocByPath(path);
        }

        // ---------- B. 插图正例 ----------

        private static void RunBitmapInsert(TestRun run, WdSession s, string png)
        {
            string path, id;
            if (!TryOpen(run, s, out path, out id, BitmapIds))
            {
                return;
            }

            ToolResult ins = s.InsertImage(new Dictionary<string, object>
            {
                ["channel_id"] = id,
                ["path"] = png
            });
            string imageId = WdSession.Str(ins, "image_id");
            Ok(run, "wi-ins-ok", ins.Success, ins.Error ?? "插图失败");
            Ok(run, "wi-ins-fields",
                ins.Success
                    && WdSession.Bool(ins, "inserted")
                    && Num(ins, "image_width_cm") > 0
                    && Num(ins, "image_height_cm") > 0,
                ins.Success ? "回包缺 inserted / 宽高" : (ins.Error ?? ""));
            Ok(run, "wi-ins-id",
                ins.Success && imageId.StartsWith("I_", StringComparison.Ordinal),
                ins.Success ? "无 I_（image_id=" + imageId + "）" : (ins.Error ?? ""));
            string body = ReadBody(s, id);
            Ok(run, "wi-ins-text",
                ins.Success && body.IndexOf("<image", StringComparison.Ordinal) >= 0,
                ins.Success ? "正文无 <image" : (ins.Error ?? ""));

            ToolResult wide = s.InsertImage(new Dictionary<string, object>
            {
                ["channel_id"] = id,
                ["path"] = png,
                ["width_cm"] = 4.0
            });
            double ratio = FixturePng.WidthPx / (double)FixturePng.HeightPx;
            double w = Num(wide, "image_width_cm");
            double h = Num(wide, "image_height_cm");
            Ok(run, "wi-ins-width",
                wide.Success
                    && Math.Abs(w - 4.0) < 0.2
                    && h > 0
                    && Math.Abs(w / h - ratio) < 0.15,
                wide.Success
                    ? "宽高不符：w=" + w + " h=" + h + " 期望比 " + ratio.ToString("F2")
                    : (wide.Error ?? ""));

            s.CloseDocByPath(path);
        }

        private static void RunKeepSent(TestRun run, WdSession s, string png)
        {
            string path, id;
            if (!TryOpen(run, s, out path, out id, KeepSentIds))
            {
                return;
            }

            if (!s.TypePlain(path, BeforeSent))
            {
                FailMany(run, "COM 写图前句失败", KeepSentIds);
                s.CloseDocByPath(path);
                return;
            }

            s.ActivateDoc(path);
            s.MoveSelectionToEnd(path);
            ToolResult ins = s.InsertImage(new Dictionary<string, object>
            {
                ["channel_id"] = id,
                ["path"] = png
            });
            string body = ReadBody(s, id);
            Ok(run, "wi-ins-keep-sent",
                ins.Success && body.IndexOf(BeforeSent, StringComparison.Ordinal) >= 0,
                ins.Success ? "前句丢" : (ins.Error ?? ""));
            s.CloseDocByPath(path);
        }

        private static void RunAnchorInsert(TestRun run, WdSession s, string png)
        {
            string path, id;
            if (!TryOpen(run, s, out path, out id, AnchorIds))
            {
                return;
            }

            if (!s.TypePlain(path, AnchorSent))
            {
                FailMany(run, "COM 写锚句失败", AnchorIds);
                s.CloseDocByPath(path);
                return;
            }

            s.ActivateDoc(path);
            string body = ReadBody(s, id);
            string code = SentCodeOf(body, AnchorSent);
            if (string.IsNullOrEmpty(code))
            {
                FailMany(run, "回读无锚句编码（body=" + Snip(body) + "）", AnchorIds);
                s.CloseDocByPath(path);
                return;
            }

            ToolResult ins = s.InsertImage(new Dictionary<string, object>
            {
                ["channel_id"] = id,
                ["path"] = png,
                ["target"] = new Dictionary<string, object> { ["codes"] = code }
            });
            string after = ReadBody(s, id);
            Ok(run, "wi-ins-after-codes",
                ins.Success
                    && after.IndexOf(AnchorSent, StringComparison.Ordinal) >= 0
                    && after.IndexOf("<image", StringComparison.Ordinal) >= 0,
                ins.Success ? "句后插图后正文缺句或缺图" : (ins.Error ?? ""));
            s.CloseDocByPath(path);
        }

        private static void RunStartInsert(TestRun run, WdSession s, string png)
        {
            string path, id;
            if (!TryOpen(run, s, out path, out id, StartIds))
            {
                return;
            }

            ToolResult ins = s.InsertImage(new Dictionary<string, object>
            {
                ["channel_id"] = id,
                ["path"] = png,
                ["target"] = new Dictionary<string, object> { ["position"] = "start" }
            });
            string body = ReadBody(s, id);
            Ok(run, "wi-ins-start",
                ins.Success && body.IndexOf("<image", StringComparison.Ordinal) >= 0,
                ins.Success ? "文首插图后正文无 <image" : (ins.Error ?? ""));
            s.CloseDocByPath(path);
        }

        private static void RunSvgTextInsert(TestRun run, WdSession s, string svgText)
        {
            string path, id;
            if (!TryOpen(run, s, out path, out id, SvgTextIds))
            {
                return;
            }

            ToolResult ins = s.InsertSvgImage(new Dictionary<string, object>
            {
                ["channel_id"] = id,
                ["svg"] = svgText
            });
            string imageId = WdSession.Str(ins, "image_id");
            Ok(run, "wi-ins-svg-text",
                ins.Success
                    && WdSession.Bool(ins, "inserted")
                    && imageId.StartsWith("I_", StringComparison.Ordinal),
                ins.Success ? "回包缺 inserted / I_（image_id=" + imageId + "）" : (ins.Error ?? ""));
            s.CloseDocByPath(path);
        }

        private static void RunSvgFileInsert(TestRun run, WdSession s, string svgPath)
        {
            string path, id;
            if (!TryOpen(run, s, out path, out id, SvgFileIds))
            {
                return;
            }

            ToolResult ins = s.InsertSvgImage(new Dictionary<string, object>
            {
                ["channel_id"] = id,
                ["path"] = svgPath
            });
            Ok(run, "wi-ins-svg-file", ins.Success, ins.Error ?? "SVG 文件插图失败");
            s.CloseDocByPath(path);
        }

        // ---------- D. 页截图负例 ----------

        private static void RunCaptureNeg(TestRun run, WdSession s)
        {
            ToolResult noCh = s.CapturePageImage(new Dictionary<string, object>());
            if (!noCh.Success
                && (Contains(noCh.Error, "打开")
                    || Contains(noCh.Error, "unsupported")
                    || Contains(noCh.Error, "browser")
                    || Contains(noCh.Error, "渠道")))
            {
                run.CaseOk("wp-neg-no-channel");
            }
            else if (noCh.Success)
            {
                run.CaseSkip("wp-neg-no-channel", "现场已有默认文字渠，不关用户稿");
            }
            else
            {
                run.CaseFail("wp-neg-no-channel", noCh.Error ?? "应失败或跳过");
            }

            ExpectFail(run, "wp-neg-unknown-id",
                s.CapturePageImage(new Dictionary<string, object> { ["channel_id"] = "w999" }),
                "未知");
            RunSheetRangeNeg(run, s);

            string[] capNeedChannel =
            {
                "wp-neg-span", "wp-neg-storage", "wp-neg-shape",
                "wp-neg-page-zero", "wp-neg-page-over"
            };
            string capPath, capId;
            if (!TryOpen(run, s, out capPath, out capId, capNeedChannel))
            {
                return;
            }

            ExpectFail(run, "wp-neg-span",
                s.CapturePageImage(new Dictionary<string, object>
                {
                    ["channel_id"] = capId,
                    ["page_from"] = 1,
                    ["page_to"] = 2
                }),
                "page_from");
            ExpectFail(run, "wp-neg-storage",
                s.CapturePageImage(new Dictionary<string, object>
                {
                    ["channel_id"] = capId,
                    ["storage_doc_uuid"] = "u-1"
                }),
                "storage_doc_uuid");
            ExpectFail(run, "wp-neg-shape",
                s.CapturePageImage(new Dictionary<string, object>
                {
                    ["channel_id"] = capId,
                    ["shape_id"] = "P_1"
                }),
                "shape");
            ExpectFail(run, "wp-neg-page-zero",
                s.CapturePageImage(new Dictionary<string, object>
                {
                    ["channel_id"] = capId,
                    ["page_number"] = 0
                }),
                "page_number");
            ExpectFail(run, "wp-neg-page-over",
                s.CapturePageImage(new Dictionary<string, object>
                {
                    ["channel_id"] = capId,
                    ["page_number"] = 999
                }),
                "页");
            s.CloseDocByPath(capPath);
        }

        // ---------- E. 页截图正例 ----------

        private static void RunCapturePos(TestRun run, WdSession s)
        {
            string[] ids = { "wp-cap-omit", "wp-cap-fields", "wp-cap-p1" };
            string path, id;
            if (!TryOpen(run, s, out path, out id, ids))
            {
                return;
            }

            if (!s.TypePlain(path, "易写截图第一页。"))
            {
                FailMany(run, "COM 写页失败", ids);
                s.CloseDocByPath(path);
                return;
            }

            s.ActivateDoc(path);

            ToolResult omit = s.CapturePageImage(new Dictionary<string, object> { ["channel_id"] = id });
            Ok(run, "wp-cap-omit",
                omit.Success && WdSession.Str(omit, "image_base64").Length > 0,
                omit.Success ? "image_base64 空" : (omit.Error ?? ""));
            Ok(run, "wp-cap-fields",
                omit.Success
                    && Num(omit, "width_px") > 0
                    && Num(omit, "height_px") > 0
                    && Num(omit, "total_pages") >= 1
                    && WdSession.Str(omit, "format").Length > 0
                    && Num(omit, "page_number") >= 1,
                omit.Success
                    ? "字段缺：w=" + Num(omit, "width_px") + " h=" + Num(omit, "height_px")
                        + " pages=" + Num(omit, "total_pages") + " fmt=" + WdSession.Str(omit, "format")
                        + " page=" + Num(omit, "page_number")
                    : (omit.Error ?? ""));

            ToolResult p1 = s.CapturePageImage(new Dictionary<string, object>
            {
                ["channel_id"] = id,
                ["page_number"] = 1
            });
            Ok(run, "wp-cap-p1",
                p1.Success && Num(p1, "page_number") == 1,
                p1.Success ? "page_number=" + Num(p1, "page_number") : (p1.Error ?? ""));

            s.CloseDocByPath(path);
        }

        // ---------- C/F. 哨兵 ----------

        private static void RunNoQuit(
            TestRun run,
            int wordBefore,
            int wpsBefore,
            bool wordWinBefore,
            bool wpsWinBefore)
        {
            bool ok = true;
            string why = "";
            if (wordWinBefore && !OfficeAlive.HasWindow("WINWORD"))
            {
                ok = false;
                why = "用户 Word 窗没了";
            }
            else if (wpsWinBefore && !OfficeAlive.HasWindow("wps"))
            {
                ok = false;
                why = "用户 WPS 窗没了";
            }

            if (ok)
            {
                run.CaseOk("wi-no-quit");
                run.CaseOk("wp-no-quit");
            }
            else
            {
                run.CaseFail("wi-no-quit", why);
                run.CaseFail("wp-no-quit", why);
            }
        }

        // ---------- 帮手 ----------

        private static Dictionary<string, object> TargetArgs(
            string channelId,
            string png,
            Dictionary<string, object> target)
        {
            return new Dictionary<string, object>
            {
                ["channel_id"] = channelId,
                ["path"] = png,
                ["target"] = target
            };
        }

        private static void RunSheetRangeNeg(TestRun run, WdSession s)
        {
            string path = s.NewDocxPath();
            ToolResult opened = s.Open(path, true, s.App);
            if (!opened.Success)
            {
                run.CaseFail("wp-neg-sheet-range", "开稿失败: " + (opened.Error ?? ""));
                s.CloseDocByPath(path);
                return;
            }

            string id = WdSession.Str(opened, "channel_id");
            s.ActivateDoc(path);
            ExpectFail(run, "wp-neg-sheet-range",
                s.CapturePageImage(new Dictionary<string, object>
                {
                    ["channel_id"] = id,
                    ["sheet"] = "Sheet1",
                    ["range"] = "A1:B2"
                }),
                "sheet");
            s.CloseDocByPath(path);
        }

        private static bool TryOpen(
            TestRun run,
            WdSession s,
            out string path,
            out string id,
            params string[] ids)
        {
            path = s.NewDocxPath();
            id = null;
            ToolResult opened = s.Open(path, true, s.App);
            if (!opened.Success)
            {
                FailMany(run, "开稿失败: " + (opened.Error ?? ""), ids);
                s.CloseDocByPath(path);
                return false;
            }

            id = WdSession.Str(opened, "channel_id");
            s.ActivateDoc(path);
            return true;
        }

        private static string ReadBody(WdSession s, string id)
        {
            ToolResult got = s.GetContent(id);
            return got.Success ? WdSession.Str(got, "document_content") : "";
        }

        private static string SentCodeOf(string body, string sent)
        {
            if (string.IsNullOrEmpty(body) || string.IsNullOrEmpty(sent))
            {
                return "";
            }

            int idx = body.IndexOf(sent, StringComparison.Ordinal);
            if (idx < 0)
            {
                return "";
            }

            var m = Regex.Match(body.Substring(0, idx), @"\[S_([A-Za-z0-9]+)-start\][^\[]*$");
            if (m.Success)
            {
                return "S_" + m.Groups[1].Value;
            }

            var m2 = Regex.Match(body.Substring(0, idx), @"S_[A-Za-z0-9]+");
            return m2.Success ? m2.Value : "";
        }

        private static string WriteSvgFile()
        {
            string path = Path.Combine(
                Path.GetTempPath(),
                "ew-wi-" + Guid.NewGuid().ToString("N") + ".svg");
            File.WriteAllText(path, LegalSvg());
            return path;
        }

        private static string LegalSvg()
        {
            return "<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"24\" height=\"24\">"
                + "<rect width=\"24\" height=\"24\" fill=\"#4472C4\"/></svg>";
        }

        private static string Snip(string body)
        {
            if (string.IsNullOrEmpty(body))
            {
                return "";
            }

            return body.Length <= 80 ? body : body.Substring(0, 80);
        }

        private static double Num(ToolResult r, string key)
        {
            Dictionary<string, object> d = Data(r);
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

        private static Dictionary<string, object> Data(ToolResult r)
        {
            return r == null ? null : r.Data as Dictionary<string, object>;
        }

        private static void ExpectFail(TestRun run, string id, ToolResult r, string needle)
        {
            if (r != null && !r.Success && (Contains(r.Error, needle) || BlobHas(r, needle)))
            {
                run.CaseOk(id);
            }
            else if (r != null && !r.Success)
            {
                run.CaseFail(id, "失败文案不含「" + needle + "」: " + (r.Error ?? ""));
            }
            else
            {
                run.CaseFail(id, "应失败");
            }
        }

        private static bool BlobHas(ToolResult r, string needle)
        {
            if (r == null || r.Data == null || string.IsNullOrEmpty(needle))
            {
                return false;
            }

            try
            {
                return JsonConvert.SerializeObject(r.Data).IndexOf(needle, StringComparison.OrdinalIgnoreCase) >= 0;
            }
            catch
            {
                return false;
            }
        }

        private static bool Contains(string text, string needle)
        {
            return !string.IsNullOrEmpty(text)
                && text.IndexOf(needle, StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private static void Ok(TestRun run, string id, bool cond, string reason)
        {
            if (cond)
            {
                run.CaseOk(id);
            }
            else
            {
                run.CaseFail(id, reason ?? "");
            }
        }

        private static void FailMany(TestRun run, string reason, params string[] ids)
        {
            foreach (string id in ids)
            {
                run.CaseFail(id, reason);
            }
        }

        /// <summary>本机夹具 PNG：用 GDI+ 画一张小图，不依赖网络。</summary>
        private static class FixturePng
        {
            public const int WidthPx = 96;
            public const int HeightPx = 48;

            private static string _path;

            public static string Ensure()
            {
                if (_path != null && File.Exists(_path))
                {
                    return _path;
                }

                string path = Path.Combine(
                    Path.GetTempPath(),
                    "ew-wi-" + Guid.NewGuid().ToString("N") + ".png");
                using (var bmp = new System.Drawing.Bitmap(WidthPx, HeightPx))
                using (var g = System.Drawing.Graphics.FromImage(bmp))
                {
                    g.Clear(System.Drawing.Color.FromArgb(0x44, 0x72, 0xC4));
                    bmp.Save(path, System.Drawing.Imaging.ImageFormat.Png);
                }

                _path = path;
                return path;
            }
        }
    }
}
