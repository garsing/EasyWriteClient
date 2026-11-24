using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using WordAddIn1;

namespace WordDocumentRoundtripTest
{
    /// <summary>
    /// 同一批同一 P_：不相交 text 照写，Range 相交双方不写。前缀 pw-。
    /// </summary>
    internal static class PwCases
    {
        public const int Target = 17;

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
                && !filter.Equals("pw", StringComparison.OrdinalIgnoreCase)
                && !filter.Equals("pw-", StringComparison.OrdinalIgnoreCase))
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
            RunTwoText(run, s);
            RunThreeText(run, s);
            RunIndexTwo(run, s);
            RunAbut(run, s);
            RunReplaceInsert(run, s);
            RunDupWhole(run, s);
            RunDupDelete(run, s);
            RunMsgOv(run, s);
            RunWholeAndText(run, s);
            RunCover(run, s);
            RunCellTwo(run, s);
            RunCellLast(run, s);
            RunCellDupWhole(run, s);
            RunPartial(run, s);
            RunInsertSame(run, s);
            RunTwoPara(run, s);
            if (Want("pw-no-quit"))
            {
                string wordDetail;
                string wpsDetail;
                bool wordOk = OfficeAlive.NotQuit("WINWORD", wordBefore, wordWinBefore, out wordDetail);
                bool wpsOk = OfficeAlive.NotQuit("wps", wpsBefore, wpsWinBefore, out wpsDetail);
                Ok(run, "pw-no-quit", wordOk && wpsOk, wordDetail + " " + wpsDetail);
            }
        }

        private static bool Want(string id)
        {
            return Only == null || Only.Contains(id);
        }

        private static void RunTwoText(TestRun run, WdSession s)
        {
            if (!Want("pw-two-text"))
            {
                return;
            }

            using (Planted doc = PlantBody(s, "见[3]和[4]。"))
            {
                if (!Ready(run, doc, "pw-two-text"))
                {
                    return;
                }

                string code = doc.Paragraph("见[3]和[4]");
                ToolResult wrote = s.ProcessActions(doc.Id, Batch(
                    Slice(code, "[3]", "[3改]"),
                    Slice(code, "[4]", "[4改]")));
                Ok(run, "pw-two-text",
                    wrote.Success
                    && doc.Has("[3改]") && doc.Has("[4改]")
                    && !doc.Has("[3]和") && !doc.Has("和[4]")
                    && !BlobHas(wrote, "重复出现"),
                    wrote.Success ? "同一段两处 text 没有都写上" : (wrote.Error ?? ""));
            }
        }

        private static void RunThreeText(TestRun run, WdSession s)
        {
            if (!Want("pw-three-text"))
            {
                return;
            }

            using (Planted doc = PlantBody(s, "见[1]、[2]、[3]。"))
            {
                if (!Ready(run, doc, "pw-three-text"))
                {
                    return;
                }

                string code = doc.Paragraph("见[1]、[2]、[3]");
                ToolResult wrote = s.ProcessActions(doc.Id, Batch(
                    Slice(code, "[1]", "[1改]"),
                    Slice(code, "[2]", "[2改]"),
                    Slice(code, "[3]", "[3改]")));
                Ok(run, "pw-three-text",
                    wrote.Success && doc.Has("[1改]") && doc.Has("[2改]") && doc.Has("[3改]"),
                    wrote.Success ? "三处 text 没有都写上" : (wrote.Error ?? ""));
            }
        }

        private static void RunIndexTwo(TestRun run, WdSession s)
        {
            if (!Want("pw-index-two"))
            {
                return;
            }

            using (Planted doc = PlantBody(s, "见[1]甲[1]乙。"))
            {
                if (!Ready(run, doc, "pw-index-two"))
                {
                    return;
                }

                string code = doc.Paragraph("见[1]甲[1]乙");
                ToolResult wrote = s.ProcessActions(doc.Id, Batch(
                    SliceIndex(code, "[1]", 1, "甲标"),
                    SliceIndex(code, "[1]", 2, "乙标")));
                Ok(run, "pw-index-two",
                    wrote.Success && doc.Has("甲标") && doc.Has("乙标") && !doc.Has("[1]"),
                    wrote.Success ? "两处 index 没有都写上" : (wrote.Error ?? ""));
            }
        }

        private static void RunAbut(TestRun run, WdSession s)
        {
            if (!Want("pw-abut"))
            {
                return;
            }

            using (Planted doc = PlantBody(s, "见[3]和[4]。"))
            {
                if (!Ready(run, doc, "pw-abut"))
                {
                    return;
                }

                string code = doc.Paragraph("见[3]和[4]");
                ToolResult wrote = s.ProcessActions(doc.Id, Batch(
                    Slice(code, "见[3]", "见甲"),
                    Slice(code, "和[4]。", "和乙。")));
                Ok(run, "pw-abut",
                    wrote.Success && doc.Has("见甲和乙。"),
                    wrote.Success ? "首尾相接两处没有都写上" : (wrote.Error ?? ""));
            }
        }

        private static void RunReplaceInsert(TestRun run, WdSession s)
        {
            if (!Want("pw-replace-insert"))
            {
                return;
            }

            using (Planted doc = PlantBody(s, "见[3]和[4]。"))
            {
                if (!Ready(run, doc, "pw-replace-insert"))
                {
                    return;
                }

                string code = doc.Paragraph("见[3]和[4]");
                ToolResult wrote = s.ProcessActions(doc.Id, Batch(
                    Slice(code, "[3]", "[3改]"),
                    InsertAfterText(code, "[4]", "插")));
                Ok(run, "pw-replace-insert",
                    wrote.Success && doc.Has("[3改]") && doc.Has("插") && doc.Has("[4]"),
                    wrote.Success ? "replace+insert 没有都写上" : (wrote.Error ?? ""));
            }
        }

        private static void RunDupWhole(TestRun run, WdSession s)
        {
            if (!Want("pw-dup-whole"))
            {
                return;
            }

            using (Planted doc = PlantBody(s, "甲段。\r乙段。"))
            {
                if (!Ready(run, doc, "pw-dup-whole"))
                {
                    return;
                }

                string a = doc.Paragraph("甲段");
                ToolResult wrote = s.ProcessActions(doc.Id, Batch(Whole(a, "甲一。"), Whole(a, "甲二。")));
                Ok(run, "pw-dup-whole",
                    !wrote.Success && BlobHas(wrote, "Range 重叠")
                    && doc.Has("甲段。") && doc.Has("乙段。")
                    && !doc.Has("甲一。") && !doc.Has("甲二。"),
                    wrote.Success ? "整段同码两次仍写了" : (wrote.Error ?? ""));
            }
        }

        private static void RunDupDelete(TestRun run, WdSession s)
        {
            if (!Want("pw-dup-delete"))
            {
                return;
            }

            using (Planted doc = PlantBody(s, "甲段。\r乙段。"))
            {
                if (!Ready(run, doc, "pw-dup-delete"))
                {
                    return;
                }

                string a = doc.Paragraph("甲段");
                ToolResult wrote = s.ProcessActions(doc.Id, Batch(Delete(a), Delete(a)));
                Ok(run, "pw-dup-delete",
                    !wrote.Success && BlobHas(wrote, "Range 重叠") && doc.Has("甲段。"),
                    wrote.Success ? "整段删除两次仍删了" : (wrote.Error ?? ""));
            }
        }

        private static void RunMsgOv(TestRun run, WdSession s)
        {
            if (!Want("pw-msg-ov"))
            {
                return;
            }

            using (Planted doc = PlantBody(s, "甲段。\r乙段。"))
            {
                if (!Ready(run, doc, "pw-msg-ov"))
                {
                    return;
                }

                string a = doc.Paragraph("甲段");
                ToolResult wrote = s.ProcessActions(doc.Id, Batch(Whole(a, "甲一。"), Whole(a, "甲二。")));
                Ok(run, "pw-msg-ov",
                    !wrote.Success
                    && BlobHas(wrote, "Range 重叠")
                    && BlobHas(wrote, "action[0]")
                    && BlobHas(wrote, "action[1]")
                    && BlobHas(wrote, a),
                    wrote.Success ? "未失败" : "回包未点名重叠 action blob=" + Trim(wrote));
            }
        }

        private static void RunWholeAndText(TestRun run, WdSession s)
        {
            if (!Want("pw-whole-and-text"))
            {
                return;
            }

            using (Planted doc = PlantBody(s, "见[3]和[4]。"))
            {
                if (!Ready(run, doc, "pw-whole-and-text"))
                {
                    return;
                }

                string code = doc.Paragraph("见[3]和[4]");
                ToolResult wrote = s.ProcessActions(doc.Id, Batch(
                    Whole(code, "整段。"),
                    Slice(code, "[3]", "[3改]")));
                Ok(run, "pw-whole-and-text",
                    !wrote.Success && BlobHas(wrote, "Range 重叠") && doc.Has("见[3]和[4]。"),
                    wrote.Success ? "整段加切片仍写了" : (wrote.Error ?? ""));
            }
        }

        private static void RunCover(TestRun run, WdSession s)
        {
            if (!Want("pw-cover"))
            {
                return;
            }

            using (Planted doc = PlantBody(s, "见[3]和[4]。"))
            {
                if (!Ready(run, doc, "pw-cover"))
                {
                    return;
                }

                string code = doc.Paragraph("见[3]和[4]");
                ToolResult wrote = s.ProcessActions(doc.Id, Batch(
                    Slice(code, "见[3]和[4]", "整段切片"),
                    Slice(code, "[3]", "[3改]")));
                Ok(run, "pw-cover",
                    !wrote.Success && BlobHas(wrote, "Range 重叠") && doc.Has("见[3]和[4]。"),
                    wrote.Success ? "大切片盖住仍写了" : (wrote.Error ?? ""));
            }
        }

        private static void RunCellTwo(TestRun run, WdSession s)
        {
            if (!Want("pw-cell-two"))
            {
                return;
            }

            using (Planted doc = PlantTable(s, 1, 1, "格内[3]和[4]。"))
            {
                if (!Ready(run, doc, "pw-cell-two"))
                {
                    return;
                }

                string code = doc.Paragraph("格内[3]和[4]");
                ToolResult wrote = s.ProcessActions(doc.Id, Batch(
                    Slice(code, "[3]", "[3改]"),
                    Slice(code, "[4]", "[4改]")));
                Ok(run, "pw-cell-two",
                    wrote.Success && doc.Has("[3改]") && doc.Has("[4改]") && doc.HasTable(),
                    wrote.Success ? "格内两处 text 没有都写上" : (wrote.Error ?? ""));
            }
        }

        private static void RunCellLast(TestRun run, WdSession s)
        {
            if (!Want("pw-cell-last"))
            {
                return;
            }

            using (Planted doc = PlantTable(s, 1, 1, "格首甲。\r格末[3]和[4]。"))
            {
                if (!Ready(run, doc, "pw-cell-last"))
                {
                    return;
                }

                string last = doc.Paragraph("格末[3]和[4]");
                ToolResult wrote = s.ProcessActions(doc.Id, Batch(
                    Slice(last, "[3]", "[3改]"),
                    Slice(last, "[4]", "[4改]")));
                Ok(run, "pw-cell-last",
                    wrote.Success && doc.Has("格首甲") && doc.Has("[3改]") && doc.Has("[4改]"),
                    wrote.Success ? "格末两处 text 没有都写上或前一段没了" : (wrote.Error ?? ""));
            }
        }

        private static void RunCellDupWhole(TestRun run, WdSession s)
        {
            if (!Want("pw-cell-dup-whole"))
            {
                return;
            }

            using (Planted doc = PlantTable(s, 1, 1, "格内[3]和[4]。"))
            {
                if (!Ready(run, doc, "pw-cell-dup-whole"))
                {
                    return;
                }

                string code = doc.Paragraph("格内[3]和[4]");
                ToolResult wrote = s.ProcessActions(doc.Id, Batch(Whole(code, "格一。"), Whole(code, "格二。")));
                Ok(run, "pw-cell-dup-whole",
                    !wrote.Success && BlobHas(wrote, "Range 重叠") && doc.Has("格内[3]和[4]"),
                    wrote.Success ? "格内整段同码两次仍写了" : (wrote.Error ?? ""));
            }
        }

        private static void RunPartial(TestRun run, WdSession s)
        {
            if (!Want("pw-partial"))
            {
                return;
            }

            using (Planted doc = PlantBody(s, "甲段。\r乙段。"))
            {
                if (!Ready(run, doc, "pw-partial"))
                {
                    return;
                }

                string a = doc.Paragraph("甲段");
                string b = doc.Paragraph("乙段");
                ToolResult wrote = s.ProcessActions(doc.Id, Batch(
                    Whole(a, "甲一。"),
                    Whole(a, "甲二。"),
                    Whole(b, "乙已改。")));
                bool partial = WdSession.Str(wrote, "batch_outcome") == "partial_succeeded";
                Ok(run, "pw-partial",
                    partial && BlobHas(wrote, "Range 重叠")
                    && doc.Has("甲段。") && doc.Has("乙已改。") && !doc.Has("甲一。"),
                    wrote.Success ? "重叠对没有挡住甲、乙却写了" : (wrote.Error ?? ""));
            }
        }

        private static void RunInsertSame(TestRun run, WdSession s)
        {
            if (!Want("pw-insert-same"))
            {
                return;
            }

            using (Planted doc = PlantBody(s, "甲段。\r乙段。"))
            {
                if (!Ready(run, doc, "pw-insert-same"))
                {
                    return;
                }

                string a = doc.Paragraph("甲段");
                ToolResult wrote = s.ProcessActions(doc.Id, Batch(
                    InsertAfter(a, "插一。"),
                    InsertAfter(a, "插二。")));
                Ok(run, "pw-insert-same",
                    !wrote.Success && BlobHas(wrote, "Range 重叠")
                    && doc.Has("甲段。") && doc.Has("乙段。")
                    && !doc.Has("插一。") && !doc.Has("插二。"),
                    wrote.Success ? "无 text 两次后插仍写了（定位是整段 Range）" : (wrote.Error ?? ""));
            }
        }

        private static void RunTwoPara(TestRun run, WdSession s)
        {
            if (!Want("pw-two-para"))
            {
                return;
            }

            using (Planted doc = PlantBody(s, "甲段。\r乙段。"))
            {
                if (!Ready(run, doc, "pw-two-para"))
                {
                    return;
                }

                string a = doc.Paragraph("甲段");
                string b = doc.Paragraph("乙段");
                ToolResult wrote = s.ProcessActions(doc.Id, Batch(Whole(a, "甲已改。"), Whole(b, "乙已改。")));
                Ok(run, "pw-two-para",
                    wrote.Success && doc.Has("甲已改。") && doc.Has("乙已改。"),
                    wrote.Success ? "两段各改没有都写上" : (wrote.Error ?? ""));
            }
        }

        private static List<Dictionary<string, object>> Batch(params Dictionary<string, object>[] items)
        {
            return new List<Dictionary<string, object>>(items);
        }

        private static Dictionary<string, object> Whole(string code, string neu)
        {
            return new Dictionary<string, object>
            {
                ["type"] = "replace",
                ["detail"] = new Dictionary<string, object>
                {
                    ["original"] = new Dictionary<string, object> { ["paragraph_codes"] = code },
                    ["new"] = neu,
                    ["format"] = DefaultFormat()
                }
            };
        }

        private static Dictionary<string, object> Slice(string code, string text, string neu)
        {
            return new Dictionary<string, object>
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
            };
        }

        private static Dictionary<string, object> SliceIndex(string code, string text, int index, string neu)
        {
            return new Dictionary<string, object>
            {
                ["type"] = "replace",
                ["detail"] = new Dictionary<string, object>
                {
                    ["original"] = new Dictionary<string, object>
                    {
                        ["paragraph_codes"] = code,
                        ["text"] = new Dictionary<string, object>
                        {
                            ["value"] = text,
                            ["index"] = index
                        }
                    },
                    ["new"] = neu,
                    ["format"] = DefaultFormat()
                }
            };
        }

        private static Dictionary<string, object> Delete(string code)
        {
            return new Dictionary<string, object>
            {
                ["type"] = "delete",
                ["detail"] = new Dictionary<string, object>
                {
                    ["deletion"] = new Dictionary<string, object> { ["paragraph_codes"] = code }
                }
            };
        }

        private static Dictionary<string, object> InsertAfter(string code, string insert)
        {
            return new Dictionary<string, object>
            {
                ["type"] = "insert",
                ["detail"] = new Dictionary<string, object>
                {
                    ["target"] = new Dictionary<string, object> { ["paragraph_codes"] = code },
                    ["insert"] = insert,
                    ["format"] = DefaultFormat()
                }
            };
        }

        private static Dictionary<string, object> InsertAfterText(string code, string text, string insert)
        {
            return new Dictionary<string, object>
            {
                ["type"] = "insert",
                ["detail"] = new Dictionary<string, object>
                {
                    ["target"] = new Dictionary<string, object>
                    {
                        ["paragraph_codes"] = code,
                        ["text"] = text
                    },
                    ["insert"] = insert,
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

        private static bool Ready(TestRun run, Planted doc, params string[] ids)
        {
            if (doc != null && string.IsNullOrEmpty(doc.Error))
            {
                return true;
            }

            foreach (string id in ids)
            {
                if (Want(id))
                {
                    run.CaseFail(id, doc == null ? "铺稿失败" : (doc.Error ?? "铺稿失败"));
                }
            }

            return false;
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

        private static bool BlobHas(ToolResult r, string needle)
        {
            if (r != null && !string.IsNullOrEmpty(r.Error)
                && r.Error.IndexOf(needle, StringComparison.Ordinal) >= 0)
            {
                return true;
            }

            return WmCases.Blob(r).IndexOf(needle, StringComparison.Ordinal) >= 0;
        }

        private static string Trim(ToolResult r)
        {
            return WmCases.TrimBlob(WmCases.Blob(r));
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
                    string body = Regex.Replace(marker.Text ?? "", @"\[in_table=[^\]]+\]", "");
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
