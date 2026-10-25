using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using Newtonsoft.Json;
using WordAddIn1;

namespace WordDocumentRoundtripTest
{
    /// <summary>
    /// 多格几何码路径：对 process_actions 回包 <c>geo_plan_log</c> 里的 <c>[几何码路径]</c> 断言走了集合还是单步。
    /// </summary>
    internal static class WmCases
    {
        public const int Target = 8;

        private const string Title = "实验操作方法和步骤";
        private const string Note = "备注";
        private const string CellA = "格甲路径。";
        private const string CellB = "格乙路径。";

        public static void Run(
            TestRun run,
            string host,
            int wordBefore,
            int wpsBefore,
            bool wordWinBefore,
            bool wpsWinBefore)
        {
            var s = new WdSession(host);
            string warmPath = WarmWpsChannel(s);
            try
            {
                RunTwo(run, s);
                RunOne(run, s);
                RunSplit(run, s);
                RunMix(run, s);
                RunDelete(run, s);
                RunOrder(run, s);
                RunEcho(run, s);
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

        private static string WarmWpsChannel(WdSession s)
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

        private static void RunTwo(TestRun run, WdSession s)
        {
            if (!TryPlant(run, s, out string path, out string id, out Codes codes, "wm-two"))
            {
                return;
            }

            ToolResult wr = s.ProcessActions(id, new List<Dictionary<string, object>>
            {
                GeoReplace(codes.A, CellA),
                GeoReplace(codes.B, CellB)
            });
            ExpectPath(
                run,
                "wm-two",
                wr,
                have: new[]
                {
                    "plan geo_runs=1",
                    "singles=0",
                    "step=geo_run cells=2",
                    "locate via=geo_run n=1",
                    "map=once",
                    "write via=map",
                    "post_pd mode=geo_run",
                    "flash row="
                },
                forbid: new[]
                {
                    "locate via=single",
                    "step=single",
                    "post_pd mode=single"
                },
                Count("write via=map", 2),
                Count("locate via=geo_run", 1),
                Count("post_pd mode=geo_run", 1),
                Count("map=once", 1));
            s.CloseDocByPath(path);
        }

        private static void RunOne(TestRun run, WdSession s)
        {
            if (!TryPlant(run, s, out string path, out string id, out Codes codes, "wm-one"))
            {
                return;
            }

            ToolResult wr = s.ProcessActions(id, new List<Dictionary<string, object>>
            {
                GeoReplace(codes.A, CellA)
            });
            ExpectPath(
                run,
                "wm-one",
                wr,
                have: new[]
                {
                    "plan geo_runs=0",
                    "singles=1",
                    "step=single",
                    "geo=true",
                    "locate via=single n=1",
                    "post_pd mode=single"
                },
                forbid: new[]
                {
                    "step=geo_run",
                    "locate via=geo_run",
                    "map=once",
                    "write via=map",
                    "post_pd mode=geo_run"
                },
                Count("locate via=single", 1),
                Count("post_pd mode=single", 1));
            s.CloseDocByPath(path);
        }

        private static void RunSplit(TestRun run, WdSession s)
        {
            if (!TryPlant(run, s, out string path, out string id, out Codes codes, "wm-split"))
            {
                return;
            }

            if (string.IsNullOrEmpty(codes.Note))
            {
                ToolResult peek = s.GetContent(id);
                string peekBody = peek.Success ? WdSession.Str(peek, "document_content") : (peek.Error ?? "");
                run.CaseFail("wm-split", "无备注 S_ body=" + TrimBlob(peekBody));
                s.CloseDocByPath(path);
                return;
            }

            ToolResult wr = s.ProcessActions(id, new List<Dictionary<string, object>>
            {
                GeoReplace(codes.A, CellA),
                SentReplace(codes.Note, "备注改。"),
                GeoReplace(codes.B, CellB)
            });
            ExpectPath(
                run,
                "wm-split",
                wr,
                have: new[]
                {
                    "plan geo_runs=0",
                    "singles=3",
                    "step=single",
                    "locate via=single",
                    "post_pd mode=single"
                },
                forbid: new[]
                {
                    "step=geo_run",
                    "locate via=geo_run",
                    "map=once",
                    "write via=map",
                    "post_pd mode=geo_run"
                },
                Count("step=single", 3),
                Count("post_pd mode=single", 3));
            s.CloseDocByPath(path);
        }

        private static void RunMix(TestRun run, WdSession s)
        {
            if (!TryPlant(run, s, out string path, out string id, out Codes codes, "wm-mix"))
            {
                return;
            }

            ToolResult wr = s.ProcessActions(id, new List<Dictionary<string, object>>
            {
                GeoInsert(codes.A, CellA),
                GeoReplace(codes.B, CellB)
            });
            ExpectPath(
                run,
                "wm-mix",
                wr,
                have: new[]
                {
                    "plan geo_runs=1",
                    "step=geo_run cells=2",
                    "locate via=geo_run n=1",
                    "write via=map type=insert",
                    "write via=map type=replace",
                    "post_pd mode=geo_run"
                },
                forbid: new[]
                {
                    "locate via=single",
                    "post_pd mode=single"
                },
                Count("write via=map", 2),
                Count("locate via=geo_run", 1));
            s.CloseDocByPath(path);
        }

        private static void RunDelete(TestRun run, WdSession s)
        {
            if (!TryPlant(run, s, out string path, out string id, out Codes codes, "wm-delete"))
            {
                return;
            }

            ToolResult wr = s.ProcessActions(id, new List<Dictionary<string, object>>
            {
                GeoDelete(codes.A),
                GeoDelete(codes.B)
            });
            ExpectPath(
                run,
                "wm-delete",
                wr,
                have: new[]
                {
                    "plan geo_runs=1",
                    "step=geo_run cells=2",
                    "locate via=geo_run n=1",
                    "write via=map type=delete",
                    "post_pd mode=geo_run"
                },
                forbid: new[]
                {
                    "locate via=single",
                    "post_pd mode=single"
                },
                Count("write via=map type=delete", 2),
                Count("locate via=geo_run", 1),
                Count("post_pd mode=geo_run", 1));
            s.CloseDocByPath(path);
        }

        private static void RunOrder(TestRun run, WdSession s)
        {
            if (!TryPlant(run, s, out string path, out string id, out Codes codes, "wm-order"))
            {
                return;
            }

            EmptyCellCode.Parsed pa;
            EmptyCellCode.Parsed pb;
            if (!EmptyCellCode.TryParse(codes.A, out pa) || !EmptyCellCode.TryParse(codes.B, out pb))
            {
                run.CaseFail("wm-order", "空格码解析失败");
                s.CloseDocByPath(path);
                return;
            }

            int later = pa.Row > pb.Row || (pa.Row == pb.Row && pa.Col > pb.Col) ? pa.Row : pb.Row;
            int laterCol = pa.Row > pb.Row || (pa.Row == pb.Row && pa.Col > pb.Col) ? pa.Col : pb.Col;
            int earlier = pa.Row < pb.Row || (pa.Row == pb.Row && pa.Col < pb.Col) ? pa.Row : pb.Row;
            int earlierCol = pa.Row < pb.Row || (pa.Row == pb.Row && pa.Col < pb.Col) ? pa.Col : pb.Col;

            ToolResult wr = s.ProcessActions(id, new List<Dictionary<string, object>>
            {
                GeoReplace(codes.A, CellA),
                GeoReplace(codes.B, CellB)
            });
            string blob = Blob(wr);
            string laterTok = "write via=map type=replace r=" + later + " c=" + laterCol;
            string earlierTok = "write via=map type=replace r=" + earlier + " c=" + earlierCol;
            int iLater = blob.IndexOf(laterTok, StringComparison.Ordinal);
            int iEarlier = blob.IndexOf(earlierTok, StringComparison.Ordinal);
            bool geo = blob.IndexOf("step=geo_run cells=2", StringComparison.Ordinal) >= 0
                && blob.IndexOf("locate via=single", StringComparison.Ordinal) < 0;
            Ok(run, "wm-order",
                wr != null && wr.Success && geo && iLater >= 0 && iEarlier >= 0 && iLater < iEarlier,
                wr == null || !wr.Success
                    ? (wr == null ? "null" : wr.Error)
                    : ("later@" + iLater + " earlier@" + iEarlier + " blob=" + TrimBlob(blob)));
            s.CloseDocByPath(path);
        }

        private static void RunEcho(TestRun run, WdSession s)
        {
            if (!TryPlant(run, s, out string path, out string id, out Codes codes, "wm-echo"))
            {
                return;
            }

            ToolResult wr = s.ProcessActions(id, new List<Dictionary<string, object>>
            {
                GeoReplace(codes.A, CellA)
            });
            bool hasKey = WdSession.HasKey(wr, "geo_plan_log");
            bool inBlob = Blob(wr).IndexOf(DocumentActionPathLog.LogTag, StringComparison.Ordinal) >= 0;
            Ok(run, "wm-echo", wr != null && wr.Success && hasKey && inBlob,
                wr == null || !wr.Success
                    ? (wr == null ? "null" : wr.Error)
                    : ("hasKey=" + hasKey + " inBlob=" + inBlob));
            s.CloseDocByPath(path);
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
            Ok(run, "wm-no-quit", wordOk && wpsOk, wordDetail + " " + wpsDetail);
        }

        private struct Codes
        {
            public string A;
            public string B;
            public string Title;
            public string Note;
        }

        private static bool TryPlant(
            TestRun run,
            WdSession s,
            out string path,
            out string id,
            out Codes codes,
            params string[] failIds)
        {
            codes = new Codes();
            path = s.NewDocxPath();
            id = null;
            ToolResult opened = s.Open(path, true, s.App);
            if (!opened.Success)
            {
                FailMany(run, opened.Error ?? "新建失败", failIds);
                return false;
            }

            id = WdSession.Str(opened, "channel_id");
            s.ActivateDoc(path);
            ToolResult created = s.CreateTableFromXml(CreateArgs(id, WriteXml(GridXml(Title, "", Note, ""))));
            if (!created.Success)
            {
                FailMany(run, created.Error ?? "建表失败", failIds);
                s.CloseDocByPath(path);
                return false;
            }

            ToolResult got = s.GetContent(id);
            string body = got.Success ? WdSession.Str(got, "document_content") : "";
            var emptyCodes = new List<string>();
            foreach (Match m in Regex.Matches(body ?? "", @"\[(T_[0-9A-Za-z]+_r[0-9]+_c[0-9]+)-start\]"))
            {
                string code = m.Groups[1].Value;
                if (!emptyCodes.Contains(code))
                {
                    emptyCodes.Add(code);
                }
            }

            codes.A = emptyCodes.Count > 0 ? emptyCodes[0] : "";
            codes.B = emptyCodes.Count > 1 ? emptyCodes[1] : "";
            Match title = Regex.Match(
                body ?? "",
                @"\[(S_[0-9a-zA-Z]{5})-start(?: [^\]]*)?\](?:\[in_table=[^\]]+\])?" + Regex.Escape(Title));
            codes.Title = title.Success ? title.Groups[1].Value : "";
            Match note = Regex.Match(
                body ?? "",
                @"\[(S_[0-9a-zA-Z]{5})-start(?: [^\]]*)?\](?:\[in_table=[^\]]+\])?" + Regex.Escape(Note));
            codes.Note = note.Success ? note.Groups[1].Value : "";
            if (string.IsNullOrEmpty(codes.A) || string.IsNullOrEmpty(codes.B))
            {
                FailMany(run, "缺两个空格码 body=" + TrimBlob(body), failIds);
                s.CloseDocByPath(path);
                return false;
            }

            return true;
        }

        private static void ExpectPath(
            TestRun run,
            string id,
            ToolResult wr,
            string[] have,
            string[] forbid,
            params KeyValuePair<string, int>[] counts)
        {
            if (wr == null)
            {
                run.CaseFail(id, "null");
                return;
            }

            string blob = Blob(wr);
            if (blob.IndexOf(DocumentActionPathLog.LogTag, StringComparison.Ordinal) < 0)
            {
                run.CaseFail(id, "回包无 " + DocumentActionPathLog.LogTag + " blob=" + TrimBlob(blob));
                return;
            }

            if (have != null)
            {
                foreach (string n in have)
                {
                    if (blob.IndexOf(n, StringComparison.Ordinal) < 0)
                    {
                        run.CaseFail(id, "日志缺「" + n + "」 blob=" + TrimBlob(blob));
                        return;
                    }
                }
            }

            if (forbid != null)
            {
                foreach (string n in forbid)
                {
                    if (blob.IndexOf(n, StringComparison.Ordinal) >= 0)
                    {
                        run.CaseFail(id, "日志不该有「" + n + "」 blob=" + TrimBlob(blob));
                        return;
                    }
                }
            }

            if (counts != null)
            {
                foreach (KeyValuePair<string, int> c in counts)
                {
                    int got = CountTag(blob, c.Key);
                    if (got != c.Value)
                    {
                        run.CaseFail(
                            id,
                            "「" + c.Key + "」次数=" + got + " 期望=" + c.Value + " blob=" + TrimBlob(blob));
                        return;
                    }
                }
            }

            run.CaseOk(id);
        }

        private static KeyValuePair<string, int> Count(string needle, int n)
        {
            return new KeyValuePair<string, int>(needle, n);
        }

        private static int CountTag(string blob, string needle)
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

        private static string Blob(ToolResult r)
        {
            if (r == null || r.Data == null)
            {
                return r == null ? "" : (r.Error ?? "");
            }

            try
            {
                return JsonConvert.SerializeObject(r.Data) + " " + (r.Error ?? "");
            }
            catch
            {
                return Convert.ToString(r.Data) ?? "";
            }
        }

        private static string TrimBlob(string blob)
        {
            if (string.IsNullOrEmpty(blob))
            {
                return "";
            }

            return blob.Length <= 500 ? blob : blob.Substring(0, 500);
        }

        private static Dictionary<string, object> GeoReplace(string code, string text)
        {
            return new Dictionary<string, object>
            {
                ["type"] = "replace",
                ["detail"] = new Dictionary<string, object>
                {
                    ["original"] = new Dictionary<string, object> { ["codes"] = code },
                    ["new"] = text,
                    ["format"] = DefaultFormat()
                }
            };
        }

        private static Dictionary<string, object> GeoInsert(string code, string text)
        {
            return new Dictionary<string, object>
            {
                ["type"] = "insert",
                ["detail"] = new Dictionary<string, object>
                {
                    ["target"] = new Dictionary<string, object> { ["codes"] = code },
                    ["insert"] = text,
                    ["format"] = DefaultFormat()
                }
            };
        }

        private static Dictionary<string, object> GeoDelete(string code)
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

        private static Dictionary<string, object> SentReplace(string code, string text)
        {
            return new Dictionary<string, object>
            {
                ["type"] = "replace",
                ["detail"] = new Dictionary<string, object>
                {
                    ["original"] = new Dictionary<string, object> { ["codes"] = code },
                    ["new"] = text,
                    ["format"] = DefaultFormat()
                }
            };
        }

        private static Dictionary<string, object> DefaultFormat()
        {
            return new Dictionary<string, object>
            {
                ["char_format"] = new Dictionary<string, object>
                {
                    ["font_name"] = "宋体",
                    ["font_size"] = 12,
                    ["is_bold"] = false
                }
            };
        }

        private static Dictionary<string, object> CreateArgs(string channelId, string path)
        {
            var args = new Dictionary<string, object> { ["path"] = path ?? "" };
            if (channelId != null)
            {
                args["channel_id"] = channelId;
            }

            return args;
        }

        private static string GridXml(string a, string b, string c, string d)
        {
            return
                "<TableConfig><General><TableFontColor>000000</TableFontColor></General><Table><Properties><ColWidths><float>120</float><float>180</float></ColWidths></Properties><Data>"
                + "<tr><td>" + a + "</td><td>" + b + "</td></tr>"
                + "<tr><td>" + c + "</td><td>" + d + "</td></tr>"
                + "</Data></Table></TableConfig>";
        }

        private static string WriteXml(string xml)
        {
            string path = Path.Combine(Path.GetTempPath(), "ew-wm-" + Guid.NewGuid().ToString("N") + ".xml");
            File.WriteAllText(path, xml ?? "");
            return path;
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
    }
}
