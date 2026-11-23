using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using WordAddIn1;

namespace WordDocumentRoundtripTest
{
    /// <summary>
    /// 格末段 Range 削 \a 后 replace / insert / delete。铺稿 COM，抄当次 P_，写走 process_actions。
    /// </summary>
    internal static class PcCases
    {
        public const int Target = 20;

        private static HashSet<string> Only;

        public static void Run(
            TestRun run,
            string host,
            string filter,
            int wordBefore,
            int wpsBefore,
            bool wordWinBefore,
            bool wpsWinBefore)
        {
            Only = null;
            if (!string.IsNullOrWhiteSpace(filter)
                && !filter.Equals("pc", StringComparison.OrdinalIgnoreCase)
                && !filter.Equals("pc-", StringComparison.OrdinalIgnoreCase))
            {
                var picked = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                foreach (string part in filter.Split(','))
                {
                    string id = part.Trim();
                    if (id.Length > 0)
                    {
                        picked.Add(id);
                    }
                }

                if (picked.Count > 0)
                {
                    Only = picked;
                    run.TotalTarget = picked.Count;
                }
            }

            var s = new WdSession(host);
            RunBody(run, s);
            RunTwoParaCell(run, s);
            RunOnlyCell(run, s);
            RunGrid(run, s);
            RunThree(run, s);
            RunRefs(run, s);
            RunAroundTable(run, s);
            if (Want("pc-no-quit"))
            {
                string wordDetail;
                string wpsDetail;
                bool wordOk = OfficeAlive.NotQuit("WINWORD", wordBefore, wordWinBefore, out wordDetail);
                bool wpsOk = OfficeAlive.NotQuit("wps", wpsBefore, wpsWinBefore, out wpsDetail);
                Ok(run, "pc-no-quit", wordOk && wpsOk, wordDetail + " " + wpsDetail);
            }
        }

        private static bool Want(string id)
        {
            return Only == null || Only.Contains(id);
        }

        private static void RunBody(TestRun run, WdSession s)
        {
            if (!Want("pc-body-replace"))
            {
                return;
            }

            using (Planted doc = PlantBody(s, "正文寅。"))
            {
                if (!Ready(run, doc, "pc-body-replace"))
                {
                    return;
                }

                string code = doc.Paragraph("正文寅");
                ToolResult wrote = s.ProcessActions(doc.Id, ReplaceWhole(code, "正文寅已改。"));
                Ok(run, "pc-body-replace",
                    wrote.Success && doc.Has("正文寅已改") && !doc.Has("正文寅。"),
                    wrote.Success ? "正文整段没有换成新句" : (wrote.Error ?? ""));
            }
        }

        private static void RunTwoParaCell(TestRun run, WdSession s)
        {
            if (!Want("pc-cell-first-replace")
                && !Want("pc-cell-last-replace")
                && !Want("pc-cell-last-insert")
                && !Want("pc-cell-last-delete")
                && !Want("pc-last-text-slice")
                && !Want("pc-first-keep-last"))
            {
                return;
            }

            using (Planted doc = PlantTable(s, 1, 1, "格首甲。\r格末乙。"))
            {
                if (!Ready(run, doc,
                    "pc-cell-first-replace",
                    "pc-cell-last-replace",
                    "pc-cell-last-insert",
                    "pc-cell-last-delete",
                    "pc-last-text-slice",
                    "pc-first-keep-last"))
                {
                    return;
                }

                string first = doc.Paragraph("格首甲");
                string last = doc.Paragraph("格末乙");
                if (string.IsNullOrEmpty(first) || string.IsNullOrEmpty(last))
                {
                    FailWanted(run, "抄不到格内两段",
                        "pc-cell-first-replace",
                        "pc-cell-last-replace",
                        "pc-cell-last-insert",
                        "pc-cell-last-delete",
                        "pc-last-text-slice",
                        "pc-first-keep-last");
                    return;
                }

                if (Want("pc-cell-first-replace"))
                {
                    ToolResult wrote = s.ProcessActions(doc.Id, ReplaceWhole(first, "格首甲已改。"));
                    Ok(run, "pc-cell-first-replace",
                        wrote.Success && doc.Has("格首甲已改") && doc.Has("格末乙"),
                        wrote.Success ? "格内前一段没有改到" : (wrote.Error ?? ""));
                    doc.Reread(s);
                    first = doc.Paragraph("格首甲已改");
                    last = doc.Paragraph("格末乙");
                }

                if (Want("pc-first-keep-last"))
                {
                    Ok(run, "pc-first-keep-last",
                        doc.Has("格末乙") && !string.IsNullOrEmpty(last),
                        "改前一段后格末段丢了");
                }

                if (Want("pc-last-text-slice") && !string.IsNullOrEmpty(last))
                {
                    ToolResult slice = s.ProcessActions(doc.Id, ReplaceSlice(last, "格末乙", "格末乙片"));
                    Ok(run, "pc-last-text-slice",
                        slice.Success && doc.Has("格末乙片"),
                        slice.Success ? "格末段 text 切片没写上" : (slice.Error ?? ""));
                    doc.Reread(s);
                    last = doc.Paragraph("格末乙片");
                }

                if (Want("pc-cell-last-replace") && !string.IsNullOrEmpty(last))
                {
                    ToolResult wrote = s.ProcessActions(doc.Id, ReplaceWhole(last, "格末乙整段。"));
                    Ok(run, "pc-cell-last-replace",
                        wrote.Success && doc.Has("格末乙整段"),
                        wrote.Success ? "格末整段没有换成新句" : (wrote.Error ?? ""));
                    doc.Reread(s);
                    last = doc.Paragraph("格末乙整段");
                }

                if (Want("pc-cell-last-insert") && !string.IsNullOrEmpty(last))
                {
                    ToolResult inserted = s.ProcessActions(doc.Id, InsertAfter(last, "插在格末。"));
                    Ok(run, "pc-cell-last-insert",
                        inserted.Success && doc.Has("插在格末"),
                        inserted.Success ? "格末后插没有进格" : (inserted.Error ?? ""));
                    doc.Reread(s);
                }

                if (Want("pc-cell-last-delete"))
                {
                    string toDelete = doc.Paragraph("插在格末");
                    if (string.IsNullOrEmpty(toDelete))
                    {
                        toDelete = doc.Paragraph("格末乙");
                    }

                    ToolResult deleted = s.ProcessActions(doc.Id, DeletePara(toDelete));
                    Ok(run, "pc-cell-last-delete",
                        deleted.Success && doc.Has("格首甲") && !doc.Has("插在格末"),
                        deleted.Success ? "删格末段失败或把前一段带走了" : (deleted.Error ?? ""));
                }
            }
        }

        private static void RunOnlyCell(TestRun run, WdSession s)
        {
            if (Want("pc-cell-only-replace"))
            {
                using (Planted doc = PlantTable(s, 1, 1, "独格丙。"))
                {
                    if (!Ready(run, doc, "pc-cell-only-replace"))
                    {
                        return;
                    }

                    string code = doc.Paragraph("独格丙");
                    ToolResult wrote = s.ProcessActions(doc.Id, ReplaceWhole(code, "独格丙已改。"));
                    Ok(run, "pc-cell-only-replace",
                        wrote.Success && doc.Has("独格丙已改"),
                        wrote.Success ? "单段格整段没改上" : (wrote.Error ?? ""));
                }
            }

            if (Want("pc-cell-only-insert"))
            {
                using (Planted doc = PlantTable(s, 1, 1, "独格丙。"))
                {
                    if (!Ready(run, doc, "pc-cell-only-insert"))
                    {
                        return;
                    }

                    string code = doc.Paragraph("独格丙");
                    ToolResult inserted = s.ProcessActions(doc.Id, InsertAfter(code, "独格后续。"));
                    Ok(run, "pc-cell-only-insert",
                        inserted.Success && doc.Has("独格丙") && doc.Has("独格后续"),
                        inserted.Success ? "单段格后插失败" : (inserted.Error ?? ""));
                }
            }

            if (Want("pc-cell-only-delete"))
            {
                using (Planted doc = PlantTable(s, 1, 1, "独格丙。"))
                {
                    if (!Ready(run, doc, "pc-cell-only-delete"))
                    {
                        return;
                    }

                    string code = doc.Paragraph("独格丙");
                    ToolResult deleted = s.ProcessActions(doc.Id, DeletePara(code));
                    Ok(run, "pc-cell-only-delete",
                        deleted.Success && !doc.Has("独格丙") && doc.HasTable(),
                        deleted.Success ? "单段格删除后表没了或字还在" : (deleted.Error ?? ""));
                }
            }
        }

        private static void RunGrid(TestRun run, WdSession s)
        {
            if (Want("pc-row-first-cell") || Want("pc-row-last-cell"))
            {
                using (Planted doc = PlantTable(s, 1, 2, "左格丁。", "右格戊。"))
                {
                    if (!Ready(run, doc, "pc-row-first-cell", "pc-row-last-cell"))
                    {
                        return;
                    }

                    string left = doc.Paragraph("左格丁");
                    string right = doc.Paragraph("右格戊");
                    if (Want("pc-row-first-cell"))
                    {
                        ToolResult wrote = s.ProcessActions(doc.Id, ReplaceWhole(left, "左格丁已改。"));
                        Ok(run, "pc-row-first-cell",
                            wrote.Success && doc.Has("左格丁已改") && doc.Has("右格戊"),
                            wrote.Success ? "行首格没改上" : (wrote.Error ?? ""));
                    }

                    if (Want("pc-row-last-cell"))
                    {
                        doc.Reread(s);
                        right = doc.Paragraph("右格戊");
                        ToolResult wrote = s.ProcessActions(doc.Id, ReplaceWhole(right, "右格戊已改。"));
                        Ok(run, "pc-row-last-cell",
                            wrote.Success && doc.Has("右格戊已改"),
                            wrote.Success ? "行末格整段没改上" : (wrote.Error ?? ""));
                    }
                }
            }

            if (Want("pc-grid-br"))
            {
                using (Planted doc = PlantTable(s, 2, 2, "甲格。", "乙格。", "丙格。", "右下己。"))
                {
                    if (!Ready(run, doc, "pc-grid-br"))
                    {
                        return;
                    }

                    string br = doc.Paragraph("右下己");
                    ToolResult wrote = s.ProcessActions(doc.Id, ReplaceWhole(br, "右下己已改。"));
                    Ok(run, "pc-grid-br",
                        wrote.Success && doc.Has("右下己已改") && doc.Has("甲格"),
                        wrote.Success ? "2x2 右下格没改上" : (wrote.Error ?? ""));
                }
            }
        }

        private static void RunThree(TestRun run, WdSession s)
        {
            if (!Want("pc-three-mid-replace") && !Want("pc-three-last-replace"))
            {
                return;
            }

            using (Planted doc = PlantTable(s, 1, 1, "三格己。\r三中庚。\r三末辛。"))
            {
                if (!Ready(run, doc, "pc-three-mid-replace", "pc-three-last-replace"))
                {
                    return;
                }

                if (Want("pc-three-mid-replace"))
                {
                    string mid = doc.Paragraph("三中庚");
                    ToolResult wrote = s.ProcessActions(doc.Id, ReplaceWhole(mid, "三中庚已改。"));
                    Ok(run, "pc-three-mid-replace",
                        wrote.Success && doc.Has("三中庚已改") && doc.Has("三格己") && doc.Has("三末辛"),
                        wrote.Success ? "格内中段没改上" : (wrote.Error ?? ""));
                    doc.Reread(s);
                }

                if (Want("pc-three-last-replace"))
                {
                    string last = doc.Paragraph("三末辛");
                    ToolResult wrote = s.ProcessActions(doc.Id, ReplaceWhole(last, "三末辛已改。"));
                    Ok(run, "pc-three-last-replace",
                        wrote.Success && doc.Has("三末辛已改"),
                        wrote.Success ? "三格格末没改上" : (wrote.Error ?? ""));
                }
            }
        }

        private static void RunRefs(TestRun run, WdSession s)
        {
            const string Refs = "[1] 参一壬。\r[2] 参二癸。\r[3] 参三子。";
            if (Want("pc-refs-last-replace"))
            {
                using (Planted doc = PlantTable(s, 1, 1, Refs))
                {
                    if (!Ready(run, doc, "pc-refs-last-replace"))
                    {
                        return;
                    }

                    string last = doc.Paragraph("参三子");
                    ToolResult wrote = s.ProcessActions(doc.Id, ReplaceWhole(last, "[3] 参三子已改。"));
                    Ok(run, "pc-refs-last-replace",
                        wrote.Success && doc.Has("参三子已改") && doc.Has("参一壬"),
                        wrote.Success ? "参考文献末条没改上" : (wrote.Error ?? ""));
                }
            }

            if (Want("pc-refs-last-insert"))
            {
                using (Planted doc = PlantTable(s, 1, 1, Refs))
                {
                    if (!Ready(run, doc, "pc-refs-last-insert"))
                    {
                        return;
                    }

                    string last = doc.Paragraph("参三子");
                    ToolResult inserted = s.ProcessActions(doc.Id, InsertAfter(last, "[4] 参四新。"));
                    Ok(run, "pc-refs-last-insert",
                        inserted.Success && doc.Has("参四新") && doc.Has("参三子"),
                        inserted.Success ? "参考文献末条后插失败" : (inserted.Error ?? ""));
                }
            }
        }

        private static void RunAroundTable(TestRun run, WdSession s)
        {
            if (Want("pc-after-table-body"))
            {
                using (Planted doc = PlantTableThenBody(s, "格内。", "表后丑。"))
                {
                    if (!Ready(run, doc, "pc-after-table-body"))
                    {
                        return;
                    }

                    string code = doc.Paragraph("表后丑");
                    ToolResult wrote = s.ProcessActions(doc.Id, ReplaceWhole(code, "表后丑已改。"));
                    Ok(run, "pc-after-table-body",
                        wrote.Success && doc.Has("表后丑已改") && doc.Has("格内"),
                        wrote.Success ? "表后正文没改上" : (wrote.Error ?? ""));
                }
            }

            if (Want("pc-last-keep-table"))
            {
                using (Planted doc = PlantTable(s, 1, 1, "格首甲。\r格末乙。"))
                {
                    if (!Ready(run, doc, "pc-last-keep-table"))
                    {
                        return;
                    }

                    string last = doc.Paragraph("格末乙");
                    ToolResult wrote = s.ProcessActions(doc.Id, ReplaceWhole(last, "格末乙留表。"));
                    Ok(run, "pc-last-keep-table",
                        wrote.Success && doc.Has("格末乙留表") && doc.HasTable(),
                        wrote.Success ? "改格末后表结构没了" : (wrote.Error ?? ""));
                }
            }
        }

        private static List<Dictionary<string, object>> ReplaceWhole(string code, string neu)
        {
            return new List<Dictionary<string, object>>
            {
                new Dictionary<string, object>
                {
                    ["type"] = "replace",
                    ["detail"] = new Dictionary<string, object>
                    {
                        ["original"] = new Dictionary<string, object> { ["paragraph_codes"] = code },
                        ["new"] = neu,
                        ["format"] = DefaultFormat()
                    }
                }
            };
        }

        private static List<Dictionary<string, object>> ReplaceSlice(string code, string text, string neu)
        {
            return new List<Dictionary<string, object>>
            {
                new Dictionary<string, object>
                {
                    ["type"] = "replace",
                    ["detail"] = new Dictionary<string, object>
                    {
                        ["original"] = new Dictionary<string, object>
                        {
                            ["paragraph_codes"] = code,
                            ["text"] = text
                        },
                        ["new"] = neu,
                        ["format"] = DefaultFormat()
                    }
                }
            };
        }

        private static List<Dictionary<string, object>> InsertAfter(string code, string insert)
        {
            return new List<Dictionary<string, object>>
            {
                new Dictionary<string, object>
                {
                    ["type"] = "insert",
                    ["detail"] = new Dictionary<string, object>
                    {
                        ["target"] = new Dictionary<string, object> { ["paragraph_codes"] = code },
                        ["insert"] = insert,
                        ["format"] = DefaultFormat()
                    }
                }
            };
        }

        private static List<Dictionary<string, object>> DeletePara(string code)
        {
            return new List<Dictionary<string, object>>
            {
                new Dictionary<string, object>
                {
                    ["type"] = "delete",
                    ["detail"] = new Dictionary<string, object>
                    {
                        ["deletion"] = new Dictionary<string, object> { ["paragraph_codes"] = code }
                    }
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

        private static Planted PlantBody(WdSession s, string plain)
        {
            string path = s.NewDocxPath();
            ToolResult opened = s.Open(path, true, s.App);
            if (!opened.Success)
            {
                return new Planted { Error = opened.Error ?? "新建失败" };
            }

            string id = WdSession.Str(opened, "channel_id");
            if (!s.TypePlain(path, plain))
            {
                s.CloseDocByPath(path);
                return new Planted { Error = "COM 写句失败" };
            }

            var doc = new Planted { Path = path, Id = id, Session = s };
            doc.Reread(s);
            return doc;
        }

        private static Planted PlantTable(WdSession s, int rows, int cols, params string[] cells)
        {
            string path = s.NewDocxPath();
            ToolResult opened = s.Open(path, true, s.App);
            if (!opened.Success)
            {
                return new Planted { Error = opened.Error ?? "新建失败" };
            }

            string id = WdSession.Str(opened, "channel_id");
            if (!s.InsertTable(path, rows, cols, cells))
            {
                s.CloseDocByPath(path);
                return new Planted { Error = "COM 建表失败" };
            }

            var doc = new Planted { Path = path, Id = id, Session = s };
            doc.Reread(s);
            return doc;
        }

        private static Planted PlantTableThenBody(WdSession s, string cell, string after)
        {
            string path = s.NewDocxPath();
            ToolResult opened = s.Open(path, true, s.App);
            if (!opened.Success)
            {
                return new Planted { Error = opened.Error ?? "新建失败" };
            }

            string id = WdSession.Str(opened, "channel_id");
            if (!s.TypePlain(path, after) || !s.InsertTableBeforePara(path, 1, cell))
            {
                s.CloseDocByPath(path);
                return new Planted { Error = "COM 表后正文铺稿失败" };
            }

            var doc = new Planted { Path = path, Id = id, Session = s };
            doc.Reread(s);
            return doc;
        }

        private static bool Ready(TestRun run, Planted doc, params string[] ids)
        {
            if (doc != null && string.IsNullOrEmpty(doc.Error))
            {
                return true;
            }

            FailWanted(run, doc == null ? "铺稿失败" : (doc.Error ?? "铺稿失败"), ids);
            return false;
        }

        private static void FailWanted(TestRun run, string reason, params string[] ids)
        {
            foreach (string id in ids)
            {
                if (Want(id))
                {
                    run.CaseFail(id, reason);
                }
            }
        }

        private static void Ok(TestRun run, string id, bool pass, string reason)
        {
            if (pass)
            {
                run.CaseOk(id);
            }
            else
            {
                run.CaseFail(id, reason ?? "");
            }
        }

        private sealed class Planted : IDisposable
        {
            public string Path;
            public string Id;
            public WdSession Session;
            public string Error;
            public string Body;
            public List<Marker> Paragraphs = new List<Marker>();

            public string Paragraph(string needle)
            {
                foreach (Marker marker in Paragraphs)
                {
                    string body = Strip(marker.Text);
                    if (body.IndexOf(needle, StringComparison.Ordinal) >= 0)
                    {
                        return marker.Code;
                    }
                }

                return "";
            }

            public bool Has(string needle)
            {
                ToolResult got = Session.GetContent(Id);
                Body = got.Success ? WdSession.Str(got, "document_content") : "";
                return (Body ?? "").IndexOf(needle, StringComparison.Ordinal) >= 0;
            }

            public bool HasTable()
            {
                Has("");
                return (Body ?? "").IndexOf("<table", StringComparison.OrdinalIgnoreCase) >= 0
                    || (Body ?? "").IndexOf("[in_table=", StringComparison.Ordinal) >= 0;
            }

            public void Reread(WdSession s)
            {
                ToolResult got = s.GetContent(Id);
                Body = got.Success ? WdSession.Str(got, "document_content") : "";
                Paragraphs = Parse(Body);
            }

            public void Dispose()
            {
                if (Session != null && !string.IsNullOrEmpty(Path))
                {
                    Session.CloseDocByPath(Path);
                }
            }

            private static string Strip(string text)
            {
                return Regex.Replace(text ?? "", @"\[in_table=[^\]]+\]", "");
            }
        }

        private sealed class Marker
        {
            public string Code;
            public string Text;
        }

        private static readonly Regex ParagraphPattern = new Regex(
            @"\[(P_[0-9A-Za-z]+(?:_[1-9][0-9]*)?)(?:-start| start)\]([\s\S]*?)\[\1(?:-end| end)\]",
            RegexOptions.Compiled);

        private static List<Marker> Parse(string text)
        {
            var list = new List<Marker>();
            if (string.IsNullOrEmpty(text))
            {
                return list;
            }

            foreach (Match match in ParagraphPattern.Matches(text))
            {
                list.Add(new Marker
                {
                    Code = match.Groups[1].Value,
                    Text = match.Groups[2].Value ?? ""
                });
            }

            return list;
        }
    }
}
