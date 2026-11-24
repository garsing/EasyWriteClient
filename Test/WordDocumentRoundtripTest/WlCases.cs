using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using WordAddIn1;

namespace WordDocumentRoundtripTest
{
    /// <summary>
    /// 一批定位：Range 重叠双方不写（同码不拦）、相邻句从后往前、整批一次 PD。前缀 wl-。
    /// </summary>
    internal static class WlCases
    {
        public const int Target = 20;

        private const string Title = "实验操作方法和步骤";
        private const string Jia = "甲句。";
        private const string Yi = "乙句。";
        private const string JiaNew = "甲句已改。";
        private const string YiNew = "乙句已改。";
        private const string TitleNew = "题目已改。";
        private const string GeoFill = "空格已填。";
        private const string InsText = "插在甲后。";

        public static void Run(
            TestRun run,
            string host,
            int wordBefore,
            int wpsBefore,
            bool wordWinBefore,
            bool wpsWinBefore)
        {
            var s = new WdSession(host);
            string warmPath = WarmWps(s);
            try
            {
                RunRejects(run, s);
                RunAbut(run, s);
                RunGeoOther(run, s);
                RunPdOnce(run, s);
                RunPartial(run, s);
                RunNoStale(run, s);
                RunInsertAbut(run, s);
                RunTwoGeo(run, s);
                RunNoQuit(run, wordBefore, wpsBefore, wordWinBefore, wpsWinBefore);
            }
            finally
            {
                if (!string.IsNullOrEmpty(warmPath))
                {
                    s.CloseDocByPath(warmPath);
                }
            }
        }

        private static string WarmWps(WdSession s)
        {
            if (!s.IsWps)
            {
                return null;
            }

            string path = s.NewDocxPath();
            ToolResult opened = s.Open(path, true, s.App);
            if (!opened.Success)
            {
                return null;
            }

            string id = WdSession.Str(opened, "channel_id");
            s.ActivateDoc(path);
            s.GetContent(id);
            return path;
        }

        private static void RunRejects(TestRun run, WdSession s)
        {
            Plant p;
            if (!TryPlant(run, s, out p,
                "wl-dup-s", "wl-dup-geo", "wl-overlap-geo-s", "wl-cover",
                "wl-msg-dup", "wl-msg-ov", "wl-all-failed", "wl-del-dup", "wl-geo-del-s"))
            {
                return;
            }

            try
            {
                ToolResult dupS = s.ProcessActions(p.Id, new List<Dictionary<string, object>>
                {
                    WmCases.SentReplace(p.Jia, JiaNew),
                    WmCases.SentReplace(p.Jia, "另一改。")
                });
                string dupBlob = WmCases.Blob(dupS);
                WmCases.Ok(run, "wl-dup-s",
                    !dupS.Success
                    && WdSession.IntOf(dupS, "executed_actions") == 0
                    && StillHas(s, p.Id, Jia)
                    && !StillHas(s, p.Id, JiaNew),
                    Reason(dupS, "同一 Range 两条仍写了"));
                WmCases.Ok(run, "wl-msg-dup",
                    BlobHas(dupS, "Range 重叠") && BlobHas(dupS, p.Jia),
                    dupS.Success ? "未失败" : "回包未点名重叠码 blob=" + WmCases.TrimBlob(dupBlob));
                WmCases.Ok(run, "wl-all-failed",
                    WdSession.Str(dupS, "batch_outcome") == "all_failed",
                    "outcome=" + WdSession.Str(dupS, "batch_outcome"));

                ToolResult dupGeo = s.ProcessActions(p.Id, new List<Dictionary<string, object>>
                {
                    WmCases.GeoReplace(p.EmptyA, GeoFill),
                    WmCases.GeoReplace(p.EmptyA, "另一格。")
                });
                WmCases.Ok(run, "wl-dup-geo",
                    !dupGeo.Success
                    && WdSession.IntOf(dupGeo, "executed_actions") == 0
                    && !StillHas(s, p.Id, GeoFill),
                    Reason(dupGeo, "同格几何码仍写了"));

                ToolResult ov = s.ProcessActions(p.Id, new List<Dictionary<string, object>>
                {
                    WmCases.GeoReplace(p.FilledGeo, GeoFill),
                    WmCases.SentReplace(p.Jia, JiaNew)
                });
                WmCases.Ok(run, "wl-overlap-geo-s",
                    !ov.Success
                    && WdSession.IntOf(ov, "executed_actions") == 0
                    && StillHas(s, p.Id, Jia)
                    && !StillHas(s, p.Id, JiaNew)
                    && !StillHas(s, p.Id, GeoFill),
                    Reason(ov, "同格几何码和句码仍写了"));
                WmCases.Ok(run, "wl-msg-ov",
                    BlobHas(ov, "Range 重叠") && BlobHas(ov, p.Jia) && BlobHas(ov, p.FilledGeo),
                    ov.Success ? "未失败" : "回包未点名重叠码 blob=" + WmCases.TrimBlob(WmCases.Blob(ov)));

                ToolResult cover = s.ProcessActions(p.Id, new List<Dictionary<string, object>>
                {
                    WmCases.SentReplace(p.Jia + "," + p.Yi, JiaNew + YiNew),
                    WmCases.SentReplace(p.Yi, YiNew)
                });
                WmCases.Ok(run, "wl-cover",
                    !cover.Success
                    && WdSession.IntOf(cover, "executed_actions") == 0
                    && StillHas(s, p.Id, Jia)
                    && StillHas(s, p.Id, Yi),
                    Reason(cover, "跨句盖住仍写了"));

                ToolResult delDup = s.ProcessActions(p.Id, new List<Dictionary<string, object>>
                {
                    SentDelete(p.Jia),
                    SentDelete(p.Jia)
                });
                WmCases.Ok(run, "wl-del-dup",
                    !delDup.Success
                    && WdSession.IntOf(delDup, "executed_actions") == 0
                    && StillHas(s, p.Id, Jia),
                    Reason(delDup, "重复删除仍删了"));

                ToolResult geoDelS = s.ProcessActions(p.Id, new List<Dictionary<string, object>>
                {
                    WmCases.GeoDelete(p.FilledGeo),
                    WmCases.SentReplace(p.Jia, JiaNew)
                });
                WmCases.Ok(run, "wl-geo-del-s",
                    !geoDelS.Success
                    && StillHas(s, p.Id, Jia)
                    && !StillHas(s, p.Id, JiaNew),
                    Reason(geoDelS, "整格删+句改仍写了"));
            }
            finally
            {
                s.CloseDocByPath(p.Path);
            }
        }

        private static void RunAbut(TestRun run, WdSession s)
        {
            Plant p;
            if (!TryPlant(run, s, out p, "wl-abut-ok"))
            {
                return;
            }

            ToolResult wr = s.ProcessActions(p.Id, new List<Dictionary<string, object>>
            {
                WmCases.SentReplace(p.Jia, JiaNew),
                WmCases.SentReplace(p.Yi, YiNew)
            });
            WmCases.Ok(run, "wl-abut-ok",
                wr.Success
                && WdSession.Str(wr, "batch_outcome") == "all_succeeded"
                && StillHas(s, p.Id, JiaNew)
                && StillHas(s, p.Id, YiNew)
                && !StillHas(s, p.Id, Jia)
                && !StillHas(s, p.Id, Yi),
                Reason(wr, "相邻句应对都写成"));
            s.CloseDocByPath(p.Path);
        }

        private static void RunGeoOther(TestRun run, WdSession s)
        {
            Plant p;
            if (!TryPlant(run, s, out p, "wl-geo-other"))
            {
                return;
            }

            ToolResult wr = s.ProcessActions(p.Id, new List<Dictionary<string, object>>
            {
                WmCases.GeoReplace(p.EmptyA, GeoFill),
                WmCases.SentReplace(p.Title, TitleNew)
            });
            WmCases.Ok(run, "wl-geo-other",
                wr.Success
                && StillHas(s, p.Id, GeoFill)
                && StillHas(s, p.Id, TitleNew)
                && StillHas(s, p.Id, Jia),
                Reason(wr, "空格+题目应都能写"));
            s.CloseDocByPath(p.Path);
        }

        private static void RunPdOnce(TestRun run, WdSession s)
        {
            Plant p;
            if (!TryPlant(run, s, out p, "wl-pd-once"))
            {
                return;
            }

            ToolResult wr = s.ProcessActions(p.Id, new List<Dictionary<string, object>>
            {
                WmCases.SentReplace(p.Jia, JiaNew),
                WmCases.SentReplace(p.Yi, YiNew),
                WmCases.SentReplace(p.Title, TitleNew)
            });
            string blob = WmCases.Blob(wr);
            int pd = CountNeedle(blob, "post_pd mode=batch");
            WmCases.Ok(run, "wl-pd-once",
                wr.Success
                && pd == 1
                && blob.IndexOf("post_pd mode=single", StringComparison.Ordinal) < 0
                && blob.IndexOf("post_pd mode=geo_run", StringComparison.Ordinal) < 0,
                wr.Success ? "post_pd batch=" + pd + " blob=" + WmCases.TrimBlob(blob) : Reason(wr, ""));
            s.CloseDocByPath(p.Path);
        }

        private static void RunPartial(TestRun run, WdSession s)
        {
            Plant p;
            if (!TryPlant(run, s, out p, "wl-partial-dup", "wl-partial-out", "wl-executed", "wl-partial-ov"))
            {
                return;
            }

            try
            {
                ToolResult wr = s.ProcessActions(p.Id, new List<Dictionary<string, object>>
                {
                    WmCases.SentReplace(p.Jia, JiaNew),
                    WmCases.SentReplace(p.Jia, "另一改。"),
                    WmCases.SentReplace(p.Title, TitleNew)
                });
                WmCases.Ok(run, "wl-partial-dup",
                    !wr.Success
                    && StillHas(s, p.Id, Jia)
                    && !StillHas(s, p.Id, JiaNew)
                    && StillHas(s, p.Id, TitleNew),
                    Reason(wr, "重复两条应跳过、题目应写上"));
                WmCases.Ok(run, "wl-partial-out",
                    WdSession.Str(wr, "batch_outcome") == "partial_succeeded",
                    "outcome=" + WdSession.Str(wr, "batch_outcome"));
                WmCases.Ok(run, "wl-executed",
                    WdSession.IntOf(wr, "executed_actions") == 1
                    && WdSession.IntOf(wr, "failed_actions") == 2,
                    "executed=" + WdSession.IntOf(wr, "executed_actions")
                    + " failed=" + WdSession.IntOf(wr, "failed_actions"));

                ToolResult ov = s.ProcessActions(p.Id, new List<Dictionary<string, object>>
                {
                    WmCases.GeoReplace(p.FilledGeo, GeoFill),
                    WmCases.SentReplace(p.Yi, YiNew),
                    WmCases.GeoReplace(p.EmptyA, "旁格。")
                });
                WmCases.Ok(run, "wl-partial-ov",
                    !ov.Success
                    && StillHas(s, p.Id, Yi)
                    && !StillHas(s, p.Id, YiNew)
                    && !StillHas(s, p.Id, GeoFill)
                    && StillHas(s, p.Id, "旁格。"),
                    Reason(ov, "冲突格应跳过、另一空格应写上"));
            }
            finally
            {
                s.CloseDocByPath(p.Path);
            }
        }

        private static void RunNoStale(TestRun run, WdSession s)
        {
            Plant p;
            if (!TryPlant(run, s, out p, "wl-no-stale"))
            {
                return;
            }

            ToolResult wr = s.ProcessActions(p.Id, new List<Dictionary<string, object>>
            {
                WmCases.SentReplace(p.Jia, JiaNew),
                WmCases.SentReplace(p.Yi, YiNew)
            });
            string blob = WmCases.Blob(wr);
            WmCases.Ok(run, "wl-no-stale",
                wr.Success && blob.IndexOf("当前定位列表里没有这颗码", StringComparison.Ordinal) < 0,
                wr.Success ? "blob=" + WmCases.TrimBlob(blob) : Reason(wr, ""));
            s.CloseDocByPath(p.Path);
        }

        private static void RunInsertAbut(TestRun run, WdSession s)
        {
            Plant p;
            if (!TryPlant(run, s, out p, "wl-insert-abut"))
            {
                return;
            }

            ToolResult wr = s.ProcessActions(p.Id, new List<Dictionary<string, object>>
            {
                WmCases.SentReplace(p.Jia, JiaNew),
                SentInsert(p.Yi, InsText)
            });
            WmCases.Ok(run, "wl-insert-abut",
                wr.Success && StillHas(s, p.Id, JiaNew) && StillHas(s, p.Id, InsText) && StillHas(s, p.Id, Yi),
                Reason(wr, "改甲句并插在乙句后（不同码、首尾相接）应都能做"));
            s.CloseDocByPath(p.Path);
        }

        private static void RunTwoGeo(TestRun run, WdSession s)
        {
            Plant p;
            if (!TryPlant(run, s, out p, "wl-two-geo"))
            {
                return;
            }

            ToolResult wr = s.ProcessActions(p.Id, new List<Dictionary<string, object>>
            {
                WmCases.GeoReplace(p.EmptyA, "格甲。"),
                WmCases.GeoReplace(p.EmptyB, "格乙。")
            });
            string blob = WmCases.Blob(wr);
            WmCases.Ok(run, "wl-two-geo",
                wr.Success
                && StillHas(s, p.Id, "格甲。")
                && StillHas(s, p.Id, "格乙。")
                && CountNeedle(blob, "post_pd mode=batch") == 1,
                Reason(wr, "两空格应写成且只 PD 一次"));
            s.CloseDocByPath(p.Path);
        }

        private static void RunNoQuit(
            TestRun run,
            int wordBefore,
            int wpsBefore,
            bool wordWinBefore,
            bool wpsWinBefore)
        {
            string wordDetail;
            string wpsDetail;
            bool wordOk = OfficeAlive.NotQuit("WINWORD", wordBefore, wordWinBefore, out wordDetail);
            bool wpsOk = OfficeAlive.NotQuit("wps", wpsBefore, wpsWinBefore, out wpsDetail);
            WmCases.Ok(run, "wl-no-quit", wordOk && wpsOk, wordDetail + " " + wpsDetail);
        }

        private struct Plant
        {
            public string Path;
            public string Id;
            public string Title;
            public string Jia;
            public string Yi;
            public string EmptyA;
            public string EmptyB;
            public string FilledGeo;
        }

        private static bool TryPlant(TestRun run, WdSession s, out Plant p, params string[] failIds)
        {
            p = new Plant();
            p.Path = s.NewDocxPath();
            ToolResult opened = s.Open(p.Path, true, s.App);
            if (!opened.Success)
            {
                WmCases.FailMany(run, opened.Error ?? "新建失败", failIds);
                return false;
            }

            p.Id = WdSession.Str(opened, "channel_id");
            s.ActivateDoc(p.Path);
            ToolResult created = s.CreateTableFromXml(
                WmCases.CreateArgs(p.Id, WmCases.WriteXml(GridXml(Title, Jia + Yi, "", ""))));
            if (!created.Success)
            {
                WmCases.FailMany(run, created.Error ?? "建表失败", failIds);
                s.CloseDocByPath(p.Path);
                return false;
            }

            string body = ReadBody(s, p.Id);
            p.Title = CodeOf(body, Title);
            p.Jia = CodeOf(body, Jia);
            p.Yi = CodeOf(body, Yi);
            var geos = new List<string>();
            foreach (Match m in Regex.Matches(body ?? "", @"\[(T_[0-9A-Za-z]+_r[0-9]+_c[0-9]+)-start\]"))
            {
                string code = m.Groups[1].Value;
                if (!geos.Contains(code))
                {
                    geos.Add(code);
                }
            }

            p.EmptyA = geos.Count > 0 ? geos[0] : "";
            p.EmptyB = geos.Count > 1 ? geos[1] : "";
            EmptyCellCode.Parsed parsed;
            if (EmptyCellCode.TryParse(p.EmptyA, out parsed) && !string.IsNullOrEmpty(parsed.TableId))
            {
                p.FilledGeo = EmptyCellCode.Format(parsed.TableId, 1, 2);
            }

            if (string.IsNullOrEmpty(p.Jia)
                || string.IsNullOrEmpty(p.Yi)
                || string.IsNullOrEmpty(p.Title)
                || string.IsNullOrEmpty(p.EmptyA)
                || string.IsNullOrEmpty(p.EmptyB)
                || string.IsNullOrEmpty(p.FilledGeo))
            {
                WmCases.FailMany(run, "缺码 body=" + WmCases.TrimBlob(body), failIds);
                s.CloseDocByPath(p.Path);
                return false;
            }

            return true;
        }

        private static string GridXml(string a, string b, string c, string d)
        {
            return
                "<TableConfig><General><TableFontColor>000000</TableFontColor></General><Table><Properties><ColWidths><float>120</float><float>180</float></ColWidths></Properties><Data>"
                + "<tr><td>" + a + "</td><td>" + b + "</td></tr>"
                + "<tr><td>" + c + "</td><td>" + d + "</td></tr>"
                + "</Data></Table></TableConfig>";
        }

        private static Dictionary<string, object> SentDelete(string code)
        {
            return new Dictionary<string, object>
            {
                ["type"] = "delete",
                ["detail"] = new Dictionary<string, object>
                {
                    ["deletion"] = new Dictionary<string, object> { ["codes"] = code }
                }
            };
        }

        private static Dictionary<string, object> SentInsert(string code, string text)
        {
            return new Dictionary<string, object>
            {
                ["type"] = "insert",
                ["detail"] = new Dictionary<string, object>
                {
                    ["target"] = new Dictionary<string, object> { ["codes"] = code },
                    ["insert"] = text,
                    ["format"] = WmCases.DefaultFormat()
                }
            };
        }

        private static string CodeOf(string body, string needle)
        {
            Match m = Regex.Match(
                body ?? "",
                @"\[(S_[0-9a-zA-Z]{5})-start(?: [^\]]*)?\](?:\[in_table=[^\]]+\])?" + Regex.Escape(needle));
            return m.Success ? m.Groups[1].Value : "";
        }

        private static string ReadBody(WdSession s, string id)
        {
            ToolResult got = s.GetContent(id);
            return got.Success ? WdSession.Str(got, "document_content") : "";
        }

        private static bool StillHas(WdSession s, string id, string needle)
        {
            return ReadBody(s, id).IndexOf(needle, StringComparison.Ordinal) >= 0;
        }

        private static bool BlobHas(ToolResult r, string needle)
        {
            if (r != null && !string.IsNullOrEmpty(r.Error)
                && r.Error.IndexOf(needle, StringComparison.Ordinal) >= 0)
            {
                return true;
            }

            return WmCases.Blob(r).IndexOf(needle, StringComparison.Ordinal) >= 0;
        }

        private static string Reason(ToolResult wr, string whenOk)
        {
            if (wr == null)
            {
                return "null";
            }

            if (wr.Success)
            {
                return whenOk + " blob=" + WmCases.TrimBlob(WmCases.Blob(wr));
            }

            return wr.Error ?? WdSession.Str(wr, "message") ?? "失败";
        }

        private static int CountNeedle(string blob, string needle)
        {
            if (string.IsNullOrEmpty(blob) || string.IsNullOrEmpty(needle))
            {
                return 0;
            }

            int n = 0;
            int i = 0;
            while (true)
            {
                int at = blob.IndexOf(needle, i, StringComparison.Ordinal);
                if (at < 0)
                {
                    return n;
                }

                n++;
                i = at + needle.Length;
            }
        }
    }
}
