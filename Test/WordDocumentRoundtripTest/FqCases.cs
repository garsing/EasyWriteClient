using System;
using System.Collections.Generic;
using System.IO;
using Newtonsoft.Json.Linq;
using WordAddIn1;

namespace WordDocumentRoundtripTest
{
    /// <summary>格式搜索。对照 docs/测试/word测试/格式搜索。前缀 fq-。</summary>
    internal static class FqCases
    {
        public const int Target = 35;

        private static HashSet<string> Only;
        private static string SessionDir;

        public static void Run(TestRun run, string host, string filter)
        {
            Only = Pick(filter, "fq");
            if (Only != null)
            {
                run.TotalTarget = Only.Count;
            }

            var session = new WdSession(host);
            ToolResult info = DesktopToolPipe.Call("F_workspace_info", new Dictionary<string, object>());
            SessionDir = WdSession.Str(info, "session_dir");
            if (!info.Success || string.IsNullOrWhiteSpace(SessionDir))
            {
                FailWanted(run, info.Error ?? "工作区不可用");
                return;
            }

            RunArgs(run, session);
            RunChar(run, session);
            RunPara(run, session);
            RunFile(run, session);
            RunPath(run, session);
        }

        private static void RunArgs(TestRun run, WdSession session)
        {
            Case(run, "fq-arg-kind-miss", () =>
            {
                using (FormatProbe doc = Fresh(session, "入参甲。"))
                {
                    if (!doc.Ready)
                    {
                        return doc.Id;
                    }

                    FormatProbe.DeleteSearchFiles(SessionDir);
                    ToolResult result = Search(session, doc.Id, null, Q("is_bold", "true"), null);
                    return FailedWithoutFile(result, "format_search_character.json") ? null : Dump(result);
                }
            });
            Case(run, "fq-arg-kind-bad", () =>
            {
                using (FormatProbe doc = Fresh(session, "入参乙。"))
                {
                    if (!doc.Ready)
                    {
                        return doc.Id;
                    }

                    FormatProbe.DeleteSearchFiles(SessionDir);
                    ToolResult result = Search(session, doc.Id, "both", Q("is_bold", "true"), null);
                    return FailedWithoutFile(result, "format_search_character.json") ? null : Dump(result);
                }
            });
            Case(run, "fq-arg-query-empty", () => ArgFail(session, "character", new List<Dictionary<string, object>>(), null));
            Case(run, "fq-arg-match-miss", () => ArgFail(session, "character", new List<Dictionary<string, object>>
            {
                Clause("is_bold", "true"),
                Clause("is_italic", "true")
            }, null));
            Case(run, "fq-arg-match-bad", () => ArgFail(session, "character", Q("is_bold", "true"), "not"));
            Case(run, "fq-arg-attr-bad", () => ArgFail(session, "character", Q("outline_level", "wdOutlineLevel1"), null));
            Case(run, "fq-arg-pattern-bad", () => ArgFail(session, "character", Q("is_bold", "("), null));
            Case(run, "fq-arg-no-display", () =>
            {
                using (FormatProbe doc = Fresh(session, "未读文甲。"))
                {
                    if (!doc.Ready)
                    {
                        return doc.Id;
                    }

                    FormatProbe.DeleteSearchFiles(SessionDir);
                    ToolResult result = Search(session, doc.Id, "character", Q("is_bold", "true"), null);
                    string error = result == null ? "" : (result.Error ?? "");
                    return result != null && !result.Success
                        && error.IndexOf("unknown tool", StringComparison.OrdinalIgnoreCase) < 0
                        ? null
                        : Dump(result);
                }
            });
        }

        private static void RunChar(TestRun run, WdSession session)
        {
            Case(run, "fq-char-bold", () =>
            {
                using (FormatProbe doc = Fresh(session, "甲句加粗乙丙。"))
                {
                    string err = Prepare(doc, true, false);
                    if (err != null)
                    {
                        return err;
                    }

                    if (!session.PaintNeedle(doc.Path, "加粗", 2, Font("Bold", 1), null))
                    {
                        return "铺加粗失败";
                    }

                    JObject data = Hit(session, doc, "character", Q("is_bold", "true"), null, out err);
                    if (err != null)
                    {
                        return err;
                    }

                    JArray items = Arr(data, "items");
                    bool one = items.Count == 1 && Txt(items[0], "text") == "加粗"
                        && Flag(Obj(items[0], "char_format"), "is_bold")
                        && Txt(items[0], "target_codes").StartsWith("S_", StringComparison.Ordinal)
                        && data.ToString().IndexOf("P_", StringComparison.Ordinal) < 0;
                    return one ? null : data.ToString();
                }
            });
            Case(run, "fq-char-whole", () =>
            {
                using (FormatProbe doc = Fresh(session, "整句加粗。"))
                {
                    string err = Prepare(doc, true, false);
                    if (err != null)
                    {
                        return err;
                    }

                    if (!session.SetTextChar(doc.Path, "整句加粗", "", 0, true, false))
                    {
                        return "铺整句加粗失败";
                    }

                    doc.ReadSentences();
                    JObject data = Hit(session, doc, "character", Q("is_bold", "true"), null, out err);
                    if (err != null)
                    {
                        return err;
                    }

                    JObject item = OneItem(data);
                    return item != null && Txt(item, "target_codes") == doc.Sentence("整句加粗") && item["text"] == null
                        ? null
                        : data.ToString();
                }
            });
            Case(run, "fq-char-slice", () =>
            {
                using (FormatProbe doc = Fresh(session, "句首加粗后面不是。"))
                {
                    string err = Prepare(doc, true, false);
                    if (err != null)
                    {
                        return err;
                    }

                    if (!session.PaintNeedle(doc.Path, "句首", 2, Font("Bold", 1), null))
                    {
                        return "铺句首失败";
                    }

                    JObject data = Hit(session, doc, "character", Q("is_bold", "true"), null, out err);
                    if (err != null)
                    {
                        return err;
                    }

                    JObject item = OneItem(data);
                    doc.ReadSentences();
                    return item != null && Txt(item, "target_codes") == doc.Sentence("句首") && Txt(item, "text") == "句首"
                        ? null
                        : data.ToString();
                }
            });
            Case(run, "fq-char-cross", () =>
            {
                using (FormatProbe doc = Fresh(session, "连续甲句。连续乙句。"))
                {
                    string err = Prepare(doc, true, false);
                    if (err != null)
                    {
                        return err;
                    }

                    if (!session.SetTextChar(doc.Path, "连续甲句", "", 0, true, false)
                        || !session.SetTextChar(doc.Path, "连续乙句", "", 0, true, false))
                    {
                        return "铺连续加粗失败";
                    }

                    doc.ReadSentences();
                    JObject data = Hit(session, doc, "character", Q("is_bold", "true"), null, out err);
                    if (err != null)
                    {
                        return err;
                    }

                    JObject item = OneItem(data);
                    string codes = doc.Sentence("连续甲句") + "," + doc.Sentence("连续乙句");
                    return item != null && Txt(item, "target_codes") == codes && item["text"] == null
                        ? null
                        : data.ToString();
                }
            });
            Case(run, "fq-char-link", () =>
            {
                using (FormatProbe doc = Fresh(session, "标题链接句。"))
                {
                    string err = Prepare(doc, true, false);
                    if (err != null)
                    {
                        return err;
                    }

                    if (!StyleParagraph(doc, "标题链接句", "标题 1", "Heading 1", "heading 1"))
                    {
                        return "套标题 1 失败";
                    }

                    ToolResult ctx = session.GetFormatContext(Map(doc.Id, "target_codes", doc.Sentence("标题链接句")));
                    JObject final = Obj(RunOf(FormatProbe.Data(ctx)), "char_format");
                    string size = Txt(final, "font_size");
                    if (string.IsNullOrEmpty(size))
                    {
                        return "标题没有最终字号 " + Dump(ctx);
                    }

                    JObject data = Hit(session, doc, "character", Q("font_size", size), null, out err);
                    if (err != null)
                    {
                        return err;
                    }

                    JObject item = OneItem(data);
                    return item != null && Txt(Obj(item, "char_format"), "font_size") == size ? null : data.ToString();
                }
            });
            Case(run, "fq-char-cover", () =>
            {
                using (FormatProbe doc = Fresh(session, "样式加粗直接关掉旁边仍粗。"))
                {
                    string err = Prepare(doc, true, false);
                    if (err != null)
                    {
                        return err;
                    }

                    if (!session.AddParagraphStyle(doc.Path, "测加粗段", null, true, null)
                        || !StyleParagraph(doc, "样式加粗", "测加粗段")
                        || !session.PaintNeedle(doc.Path, "关掉", 2, Font("Bold", 0), null)
                        || !session.PaintNeedle(doc.Path, "仍粗", 2, Font("Bold", 1), null))
                    {
                        return "铺覆盖失败";
                    }

                    JObject data = Hit(session, doc, "character", Q("is_bold", "true"), null, out err);
                    if (err != null)
                    {
                        return err;
                    }

                    string blob = data.ToString();
                    bool off = blob.IndexOf("关掉", StringComparison.Ordinal) < 0;
                    bool on = blob.IndexOf("仍粗", StringComparison.Ordinal) >= 0;
                    return off && on ? null : blob;
                }
            });
            Case(run, "fq-char-header", () =>
            {
                using (FormatProbe doc = Fresh(session, "正文没有加粗。"))
                {
                    string err = Prepare(doc, true, false);
                    if (err != null)
                    {
                        return err;
                    }

                    if (!doc.Touch(BoldHeader))
                    {
                        return "铺页眉失败";
                    }

                    JObject data = Hit(session, doc, "character", Q("is_bold", "true"), null, out err);
                    if (err != null)
                    {
                        return err;
                    }

                    return Num(data, "count") == 0 ? null : data.ToString();
                }
            });
            Case(run, "fq-char-table-style", () =>
            {
                using (FormatProbe doc = Fresh(session, "表外正文。"))
                {
                    string err = Prepare(doc, true, false);
                    if (err != null)
                    {
                        return err;
                    }

                    if (!session.InsertTable(doc.Path, 1, 1, new[] { "格内不加粗。" }))
                    {
                        return "插表失败";
                    }

                    doc.Touch(StyleTable);
                    JObject data = Hit(session, doc, "character", Q("is_bold", "true"), null, out err);
                    if (err != null)
                    {
                        return err;
                    }

                    return data.ToString().IndexOf("格内不加粗", StringComparison.Ordinal) < 0 ? null : data.ToString();
                }
            });
            Case(run, "fq-char-readonly", () =>
            {
                using (FormatProbe doc = Fresh(session, "只读甲句。"))
                {
                    string err = Prepare(doc, true, false);
                    if (err != null)
                    {
                        return err;
                    }

                    if (!session.SetTextChar(doc.Path, "只读甲句", "", 0, true, false))
                    {
                        return "铺只读加粗失败";
                    }

                    string before = doc.Snapshot();
                    JObject data = Hit(session, doc, "character", Q("is_bold", "true"), null, out err);
                    if (err != null)
                    {
                        return err;
                    }

                    string after = doc.Snapshot();
                    return before == after && Num(data, "count") > 0 ? null : "正文变了或没有命中";
                }
            });
        }

        private static void RunPara(TestRun run, WdSession session)
        {
            Case(run, "fq-para-outline", () =>
            {
                using (FormatProbe doc = Fresh(session, "大纲一段。"))
                {
                    string err = Outline(doc, "大纲一段", 1);
                    if (err != null)
                    {
                        return err;
                    }

                    JObject data = Hit(session, doc, "paragraph", Q("outline_level", "wdOutlineLevel1"), null, out err);
                    if (err != null)
                    {
                        return err;
                    }

                    JObject item = OneItem(data);
                    return item != null && Txt(item, "target_paragraph_codes") == doc.Paragraph("大纲一段")
                        && item["target_codes"] == null
                        ? null
                        : data.ToString();
                }
            });
            Case(run, "fq-para-bar", () =>
            {
                using (FormatProbe doc = Fresh(session, "大纲甲段。\r大纲乙段。"))
                {
                    string err = Outline(doc, "大纲甲段", 1);
                    if (err != null)
                    {
                        return err;
                    }

                    err = Outline(doc, "大纲乙段", 2);
                    if (err != null)
                    {
                        return err;
                    }

                    JObject data = Hit(session, doc, "paragraph", Q("outline_level", "wdOutlineLevel1|wdOutlineLevel2"), null, out err);
                    return err == null && Num(data, "count") == 2 ? null : (err ?? data.ToString());
                }
            });
            Case(run, "fq-para-and", () =>
            {
                using (FormatProbe doc = Fresh(session, "又居中段。\r只大纲段。"))
                {
                    string err = Outline(doc, "又居中段", 1);
                    if (err != null)
                    {
                        return err;
                    }

                    err = Outline(doc, "只大纲段", 1);
                    if (err != null)
                    {
                        return err;
                    }

                    if (!session.SetParagraphAlign(doc.Path, 1, 1))
                    {
                        return "居中失败";
                    }

                    doc.ReadParagraphs();
                    var query = new List<Dictionary<string, object>>
                    {
                        Clause("outline_level", "wdOutlineLevel1"),
                        Clause("alignment", "center")
                    };
                    JObject data = Hit(session, doc, "paragraph", query, "and", out err);
                    if (err != null)
                    {
                        return err;
                    }

                    JObject item = OneItem(data);
                    return item != null && Txt(item, "target_paragraph_codes") == doc.Paragraph("又居中段")
                        ? null
                        : data.ToString();
                }
            });
            Case(run, "fq-para-or", () =>
            {
                using (FormatProbe doc = Fresh(session, "左大纲段。\r居中二级。"))
                {
                    string err = Outline(doc, "左大纲段", 1);
                    if (err != null)
                    {
                        return err;
                    }

                    err = Outline(doc, "居中二级", 2);
                    if (err != null)
                    {
                        return err;
                    }

                    if (!session.SetParagraphAlign(doc.Path, 2, 1))
                    {
                        return "居中失败";
                    }

                    var query = new List<Dictionary<string, object>>
                    {
                        Clause("outline_level", "wdOutlineLevel1"),
                        Clause("alignment", "center")
                    };
                    JObject data = Hit(session, doc, "paragraph", query, "or", out err);
                    return err == null && Num(data, "count") == 2 ? null : (err ?? data.ToString());
                }
            });
            Case(run, "fq-para-cover", () =>
            {
                using (FormatProbe doc = Fresh(session, "样式居中改左。"))
                {
                    if (!doc.Ready)
                    {
                        return doc.Id;
                    }

                    if (!session.AddParagraphStyle(doc.Path, "测居中段", null, null, 1)
                        || !StyleParagraph(doc, "样式居中", "测居中段")
                        || !session.SetParagraphAlign(doc.Path, 1, 0))
                    {
                        return "铺居中覆盖失败";
                    }

                    if (!doc.ReadParagraphs())
                    {
                        return "没有段落码";
                    }

                    JObject center = Hit(session, doc, "paragraph", Q("alignment", "center"), null, out string err);
                    if (center == null)
                    {
                        return err;
                    }

                    JObject left = Hit(session, doc, "paragraph", Q("alignment", "left"), null, out err);
                    if (err != null)
                    {
                        return err;
                    }

                    bool missCenter = Num(center, "count") == 0 || center.ToString().IndexOf(doc.Paragraph("样式居中"), StringComparison.Ordinal) < 0;
                    bool missLeft = Num(left, "count") == 0 || left.ToString().IndexOf(doc.Paragraph("样式居中"), StringComparison.Ordinal) < 0;
                    return missCenter && missLeft ? null : center + " / " + left;
                }
            });
            Case(run, "fq-para-drop", () =>
            {
                using (FormatProbe doc = Fresh(session, "正文甲。"))
                {
                    string err = Outline(doc, "正文甲", 1);
                    if (err != null)
                    {
                        return err;
                    }

                    if (!session.InsertTable(doc.Path, 1, 1, new[] { "格内甲。" })
                        || !session.SetCellLastParagraphOutline(doc.Path, 1))
                    {
                        return "铺格子末段失败";
                    }

                    doc.ReadParagraphs();
                    JObject data = Hit(session, doc, "paragraph", Q("outline_level", "wdOutlineLevel1"), null, out err);
                    if (err != null)
                    {
                        return err;
                    }

                    JObject item = OneItem(data);
                    return item != null && Txt(item, "target_paragraph_codes") == doc.Paragraph("正文甲")
                        ? null
                        : data.ToString();
                }
            });
            Case(run, "fq-para-empty", () =>
            {
                using (FormatProbe doc = Fresh(session, "空段前。\r\r空段后。"))
                {
                    if (!doc.Ready || !session.SetEmptyParagraphOutline(doc.Path, 1) || !doc.ReadParagraphs())
                    {
                        return doc.Ready ? "空段大纲失败" : doc.Id;
                    }

                    string empty = "";
                    foreach (Marker marker in doc.Paragraphs)
                    {
                        if (string.IsNullOrWhiteSpace((marker.Text ?? "").Replace("\r", "").Replace("\a", "")))
                        {
                            empty = marker.Code;
                        }
                    }

                    if (string.IsNullOrEmpty(empty))
                    {
                        return "读文没有空段";
                    }

                    JObject data = Hit(session, doc, "paragraph", Q("outline_level", "wdOutlineLevel1"), null, out string err);
                    if (err != null)
                    {
                        return err;
                    }

                    return data.ToString().IndexOf(empty, StringComparison.Ordinal) >= 0 ? null : data.ToString();
                }
            });
        }

        private static void RunFile(TestRun run, WdSession session)
        {
            Case(run, "fq-file-zero", () =>
            {
                using (FormatProbe doc = Fresh(session, "没有加粗。"))
                {
                    string err = Prepare(doc, true, false);
                    if (err != null)
                    {
                        return err;
                    }

                    JObject data = Hit(session, doc, "character", Q("is_bold", "true"), null, out err);
                    if (err != null && Num(data, "count") != 0)
                    {
                        return err;
                    }

                    JArray file = FileItems("format_search_character.json");
                    return Num(data, "count") == 0 && Arr(data, "items").Count == 0 && file != null && file.Count == 0 && Txt(data, "path").Length > 0
                        ? null
                        : (data == null ? err : data.ToString());
                }
            });
            Case(run, "fq-file-few", () =>
            {
                using (FormatProbe doc = Fresh(session, "少甲。\r少乙。"))
                {
                    string err = Outline(doc, "少甲", 1);
                    if (err != null)
                    {
                        return err;
                    }

                    err = Outline(doc, "少乙", 1);
                    if (err != null)
                    {
                        return err;
                    }

                    JObject data = Hit(session, doc, "paragraph", Q("outline_level", "wdOutlineLevel1"), null, out err);
                    JArray file = FileItems("format_search_paragraph.json");
                    return err == null && Num(data, "count") == 2 && SameItems(Arr(data, "items"), file) ? null : (err ?? data.ToString());
                }
            });
            Case(run, "fq-file-many", () =>
            {
                var lines = new List<string>();
                for (int i = 1; i <= 51; i++)
                {
                    lines.Add("多段" + i.ToString("00") + "。");
                }

                using (FormatProbe doc = Fresh(session, string.Join("\r", lines.ToArray())))
                {
                    if (!doc.Ready || !doc.Touch(OutlineAll))
                    {
                        return doc.Ready ? "铺 51 段大纲失败" : doc.Id;
                    }

                    if (!doc.ReadParagraphs())
                    {
                        return "没有段落码";
                    }

                    JObject data = Hit(session, doc, "paragraph", Q("outline_level", "wdOutlineLevel1"), null, out string err);
                    if (err != null)
                    {
                        return err;
                    }

                    JArray file = FileItems("format_search_paragraph.json");
                    return Num(data, "count") == 51 && data["items"] == null && file != null && file.Count == 51
                        ? null
                        : data.ToString();
                }
            });
            Case(run, "fq-file-overwrite", () =>
            {
                using (FormatProbe doc = Fresh(session, "覆盖加粗再字号。"))
                {
                    string err = Prepare(doc, true, false);
                    if (err != null)
                    {
                        return err;
                    }

                    if (!session.SetTextChar(doc.Path, "覆盖加粗", "", 16, true, false))
                    {
                        return "铺覆盖字失败";
                    }

                    Hit(session, doc, "character", Q("is_bold", "true"), null, out err);
                    JObject data = Hit(session, doc, "character", Q("font_size", "16"), null, out err);
                    JArray file = FileItems("format_search_character.json");
                    string blob = file == null ? "" : file.ToString();
                    return err == null && Txt(data, "path").IndexOf("format_search_character.json", StringComparison.Ordinal) >= 0
                        && blob.IndexOf("16", StringComparison.Ordinal) >= 0
                        && blob.IndexOf("true", StringComparison.Ordinal) < 0
                        ? null
                        : (err ?? blob);
                }
            });
            Case(run, "fq-file-names", () =>
            {
                using (FormatProbe doc = Fresh(session, "两个文件名。"))
                {
                    string err = Prepare(doc, true, true);
                    if (err != null)
                    {
                        return err;
                    }

                    JObject c = Hit(session, doc, "character", Q("font_name", "."), null, out err);
                    if (err != null)
                    {
                        return err;
                    }

                    JObject p = Hit(session, doc, "paragraph", Q("alignment", "."), null, out err);
                    if (err != null)
                    {
                        return err;
                    }

                    bool names = Txt(c, "path").EndsWith("format_search_character.json", StringComparison.Ordinal)
                        && Txt(p, "path").EndsWith("format_search_paragraph.json", StringComparison.Ordinal)
                        && File.Exists(FormatProbe.FileOf(SessionDir, "format_search_character.json"))
                        && File.Exists(FormatProbe.FileOf(SessionDir, "format_search_paragraph.json"));
                    return names ? null : c + " / " + p;
                }
            });
            Case(run, "fq-file-no-layers", () =>
            {
                using (FormatProbe doc = Fresh(session, "文件无层。"))
                {
                    string err = Outline(doc, "文件无层", 1);
                    if (err != null)
                    {
                        return err;
                    }

                    Hit(session, doc, "paragraph", Q("outline_level", "wdOutlineLevel1"), null, out err);
                    JArray file = FileItems("format_search_paragraph.json");
                    if (file == null || file.Count == 0)
                    {
                        return err ?? "文件空";
                    }

                    JObject item = file[0] as JObject;
                    return item != null && item["format_layers"] == null && item["para_format"] != null ? null : file.ToString();
                }
            });
        }

        private static void RunPath(TestRun run, WdSession session)
        {
            Case(run, "fq-path-char", () =>
            {
                using (FormatProbe doc = Fresh(session, "导回加粗。"))
                {
                    string err = Prepare(doc, true, false);
                    if (err != null)
                    {
                        return err;
                    }

                    if (!session.SetTextChar(doc.Path, "导回加粗", "", 0, true, false))
                    {
                        return "铺导回加粗失败";
                    }

                    JObject data = Hit(session, doc, "character", Q("is_bold", "true"), null, out err);
                    if (err != null)
                    {
                        return err;
                    }

                    if (!RewriteBold(FormatProbe.FileOf(SessionDir, "format_search_character.json"), false))
                    {
                        return "改文件失败";
                    }

                    ToolResult applied = session.ApplyDocumentFormat(Map(doc.Id, "path", Txt(data, "path")));
                    bool bold;
                    string font;
                    bool under;
                    session.TryTextChar(doc.Path, "导回加粗", out font, out bold, out under);
                    return applied.Success && !bold ? null : (applied.Error ?? "仍加粗");
                }
            });
            Case(run, "fq-path-para", () =>
            {
                using (FormatProbe doc = Fresh(session, "导回大纲。"))
                {
                    string err = Outline(doc, "导回大纲", 1);
                    if (err != null)
                    {
                        return err;
                    }

                    JObject data = Hit(session, doc, "paragraph", Q("outline_level", "wdOutlineLevel1"), null, out err);
                    if (err != null)
                    {
                        return err;
                    }

                    if (!RewriteOutline(FormatProbe.FileOf(SessionDir, "format_search_paragraph.json"), "wdOutlineLevelBodyText"))
                    {
                        return "改大纲文件失败";
                    }

                    ToolResult applied = session.ApplyParagraphFormat(Map(doc.Id, "path", Txt(data, "path")));
                    JObject again = Hit(session, doc, "paragraph", Q("outline_level", "wdOutlineLevel1"), null, out err);
                    return applied.Success && Num(again, "count") == 0 ? null : (applied.Error ?? again.ToString());
                }
            });
            Case(run, "fq-path-conflict", () =>
            {
                using (FormatProbe doc = Fresh(session, "路径冲突。"))
                {
                    string err = Prepare(doc, true, false);
                    if (err != null)
                    {
                        return err;
                    }

                    string before = doc.Snapshot();
                    var args = Map(doc.Id, "path", "format_search_character.json");
                    args["items"] = new List<Dictionary<string, object>>();
                    ToolResult applied = session.ApplyDocumentFormat(args);
                    return applied != null && !applied.Success && doc.Snapshot() == before ? null : (applied.Error ?? "不该成功");
                }
            });
            Case(run, "fq-path-cross", () =>
            {
                using (FormatProbe doc = Fresh(session, "字符文件交段落。"))
                {
                    string err = Prepare(doc, true, false);
                    if (err != null)
                    {
                        return err;
                    }

                    if (!session.SetTextChar(doc.Path, "字符文件", "", 0, true, false))
                    {
                        return "铺交叉失败";
                    }

                    JObject data = Hit(session, doc, "character", Q("is_bold", "true"), null, out err);
                    if (err != null)
                    {
                        return err;
                    }

                    string before = doc.Snapshot();
                    ToolResult applied = session.ApplyParagraphFormat(Map(doc.Id, "path", Txt(data, "path")));
                    return applied != null && !applied.Success && doc.Snapshot() == before ? null : (applied.Error ?? "不该成功");
                }
            });
            Case(run, "fq-path-para-cross", () =>
            {
                using (FormatProbe doc = Fresh(session, "段落文件交字符。"))
                {
                    string err = Outline(doc, "段落文件", 1);
                    if (err != null)
                    {
                        return err;
                    }

                    JObject data = Hit(session, doc, "paragraph", Q("outline_level", "wdOutlineLevel1"), null, out err);
                    if (err != null)
                    {
                        return err;
                    }

                    string before = doc.Snapshot();
                    ToolResult applied = session.ApplyDocumentFormat(Map(doc.Id, "path", Txt(data, "path")));
                    return applied != null && !applied.Success && doc.Snapshot() == before ? null : (applied.Error ?? "不该成功");
                }
            });
        }

        private static string ArgFail(WdSession session, string kind, List<Dictionary<string, object>> query, string match)
        {
            using (FormatProbe doc = Fresh(session, "入参失败。"))
            {
                if (!doc.Ready)
                {
                    return doc.Id;
                }

                FormatProbe.DeleteSearchFiles(SessionDir);
                ToolResult result = Search(session, doc.Id, kind, query, match);
                return FailedWithoutFile(result, kind == "paragraph" ? "format_search_paragraph.json" : "format_search_character.json")
                    ? null
                    : Dump(result);
            }
        }

        private static string Prepare(FormatProbe doc, bool sentences, bool paragraphs)
        {
            if (!doc.Ready)
            {
                return doc.Id;
            }

            if (sentences && !doc.ReadSentences())
            {
                return "没有句子码";
            }

            if (paragraphs && !doc.ReadParagraphs())
            {
                return "没有段落码";
            }

            return null;
        }

        private static string Outline(FormatProbe doc, string needle, int level)
        {
            if (!doc.Ready)
            {
                return doc.Id;
            }

            if (!doc.Session.SetOutlineLevel(doc.Path, needle, level) || !doc.ReadParagraphs())
            {
                return "大纲失败 " + needle;
            }

            return string.IsNullOrEmpty(doc.Paragraph(needle)) ? "没有这段码 " + needle : null;
        }

        private static JObject Hit(
            WdSession session,
            FormatProbe doc,
            string kind,
            List<Dictionary<string, object>> query,
            string match,
            out string error)
        {
            ToolResult result = Search(session, doc.Id, kind, query, match);
            JObject data = FormatProbe.Data(result);
            if (result == null || !result.Success)
            {
                error = result == null ? "搜索无回包" : (result.Error ?? data.ToString());
                return data;
            }

            error = null;
            return data;
        }

        private static ToolResult Search(
            WdSession session,
            string id,
            string kind,
            List<Dictionary<string, object>> query,
            string match)
        {
            var args = new Dictionary<string, object> { ["channel_id"] = id };
            if (kind != null)
            {
                args["kind"] = kind;
            }

            if (query != null)
            {
                args["query"] = query;
            }

            if (match != null)
            {
                args["match"] = match;
            }

            return session.SearchDocumentFormat(args);
        }

        private static bool FailedWithoutFile(ToolResult result, string name)
        {
            string full = FormatProbe.FileOf(SessionDir, name);
            string error = result == null ? "" : (result.Error ?? "");
            return result != null && !result.Success
                && error.IndexOf("unknown tool", StringComparison.OrdinalIgnoreCase) < 0
                && (string.IsNullOrEmpty(full) || !File.Exists(full));
        }

        private static bool StyleParagraph(FormatProbe doc, string needle, params string[] names)
        {
            return doc.Session.ApplyParagraphStyle(doc.Path, needle, names);
        }

        private static bool BoldHeader(object doc)
        {
            try
            {
                object sections = WdSession.ComGet(doc, "Sections");
                object section = WdSession.ComCall(sections, "Item", 1);
                object headers = WdSession.ComGet(section, "Headers");
                object header = WdSession.ComCall(headers, "Item", 1);
                object range = WdSession.ComGet(header, "Range");
                WdSession.ComSet(range, "Text", "页眉加粗");
                object font = WdSession.ComGet(range, "Font");
                WdSession.ComSet(font, "Bold", 1);
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        private static bool StyleTable(object doc)
        {
            try
            {
                object tables = WdSession.ComGet(doc, "Tables");
                object table = WdSession.ComCall(tables, "Item", 1);
                try
                {
                    WdSession.ComSet(table, "Style", "网格型");
                }
                catch (Exception)
                {
                    WdSession.ComSet(table, "Style", "Table Grid");
                }

                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        private static bool OutlineCellTail(object doc)
        {
            try
            {
                object tables = WdSession.ComGet(doc, "Tables");
                object table = WdSession.ComCall(tables, "Item", 1);
                object cell = WdSession.ComCall(table, "Cell", 1, 1);
                object range = WdSession.ComGet(cell, "Range");
                WdSession.ComSet(range, "Text", "格内甲。\r");
                object paras = WdSession.ComGet(cell, "Paragraphs");
                int count = Convert.ToInt32(WdSession.ComGet(paras, "Count"));
                object last = WdSession.ComCall(paras, "Item", count);
                WdSession.ComSet(last, "OutlineLevel", 1);
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        private static bool OutlineAll(object doc)
        {
            try
            {
                object paras = WdSession.ComGet(doc, "Paragraphs");
                int count = Convert.ToInt32(WdSession.ComGet(paras, "Count"));
                int marked = 0;
                for (int i = 1; i <= count; i++)
                {
                    object para = WdSession.ComCall(paras, "Item", i);
                    string text = Convert.ToString(WdSession.ComGet(WdSession.ComGet(para, "Range"), "Text")) ?? "";
                    if (text.IndexOf("多段", StringComparison.Ordinal) < 0)
                    {
                        continue;
                    }

                    WdSession.ComSet(para, "OutlineLevel", 1);
                    marked++;
                }

                return marked == 51;
            }
            catch (Exception)
            {
                return false;
            }
        }

        private static JArray FileItems(string name)
        {
            string full = FormatProbe.FileOf(SessionDir, name);
            if (string.IsNullOrEmpty(full) || !File.Exists(full))
            {
                return null;
            }

            try
            {
                JObject root = JObject.Parse(File.ReadAllText(full));
                return root["items"] as JArray ?? new JArray();
            }
            catch
            {
                return null;
            }
        }

        private static bool RewriteBold(string full, bool bold)
        {
            try
            {
                JObject root = JObject.Parse(File.ReadAllText(full));
                foreach (JObject item in root["items"])
                {
                    JObject format = item["char_format"] as JObject ?? new JObject();
                    format["is_bold"] = bold;
                    item["char_format"] = format;
                }

                File.WriteAllText(full, root.ToString());
                return true;
            }
            catch
            {
                return false;
            }
        }

        private static bool RewriteOutline(string full, string level)
        {
            try
            {
                JObject root = JObject.Parse(File.ReadAllText(full));
                foreach (JObject item in root["items"])
                {
                    JObject format = item["para_format"] as JObject ?? new JObject();
                    format["outline_level"] = level;
                    item["para_format"] = format;
                }

                File.WriteAllText(full, root.ToString());
                return true;
            }
            catch
            {
                return false;
            }
        }

        private static bool SameItems(JArray left, JArray right)
        {
            if (left == null || right == null || left.Count != right.Count)
            {
                return false;
            }

            for (int i = 0; i < left.Count; i++)
            {
                if (!JToken.DeepEquals(left[i], right[i]))
                {
                    return false;
                }
            }

            return true;
        }

        private static JObject RunOf(JObject data)
        {
            JArray entries = data["entries"] as JArray;
            if (entries == null)
            {
                return null;
            }

            foreach (JObject entry in entries)
            {
                JArray runs = entry["runs"] as JArray;
                if (runs != null && runs.Count > 0)
                {
                    return runs[0] as JObject;
                }
            }

            return null;
        }

        private static JObject OneItem(JObject data)
        {
            JArray items = Arr(data, "items");
            return items.Count == 1 ? items[0] as JObject : null;
        }

        private static Dictionary<string, object> Font(string key, object value)
        {
            return new Dictionary<string, object> { [key] = value };
        }

        private static List<Dictionary<string, object>> Q(string attr, string pattern)
        {
            return new List<Dictionary<string, object>> { Clause(attr, pattern) };
        }

        private static Dictionary<string, object> Clause(string attr, string pattern)
        {
            return new Dictionary<string, object> { ["attr"] = attr, ["pattern"] = pattern };
        }

        private static Dictionary<string, object> Map(string id, string key, object value)
        {
            return new Dictionary<string, object> { ["channel_id"] = id, [key] = value };
        }

        private static JArray Arr(JObject data, string key)
        {
            return data == null ? new JArray() : (data[key] as JArray ?? new JArray());
        }

        private static JObject Obj(JObject data, string key)
        {
            return data == null ? null : data[key] as JObject;
        }

        private static JObject Obj(JToken token, string key)
        {
            JObject data = token as JObject;
            return data == null ? null : data[key] as JObject;
        }

        private static string Txt(JToken token, string key)
        {
            JObject data = token as JObject;
            if (data == null || data[key] == null)
            {
                return "";
            }

            return data[key].ToString();
        }

        private static bool Flag(JObject data, string key)
        {
            if (data == null || data[key] == null)
            {
                return false;
            }

            bool bit;
            return bool.TryParse(data[key].ToString(), out bit) && bit;
        }

        private static int Num(JObject data, string key)
        {
            int n;
            return data != null && data[key] != null && int.TryParse(data[key].ToString(), out n) ? n : -1;
        }

        private static string Dump(ToolResult result)
        {
            if (result == null)
            {
                return "无回包";
            }

            return (result.Error ?? "") + " " + FormatProbe.Data(result);
        }

        private static FormatProbe Fresh(WdSession session, string text)
        {
            return FormatProbe.Open(session, text, SessionDir);
        }

        private static void Case(TestRun run, string id, Func<string> body)
        {
            if (Only != null && !Only.Contains(id))
            {
                return;
            }

            try
            {
                string error = body();
                if (error == null)
                {
                    run.CaseOk(id);
                }
                else
                {
                    run.CaseFail(id, error);
                }
            }
            catch (Exception ex)
            {
                run.CaseFail(id, ex.Message);
            }
        }

        private static void FailWanted(TestRun run, string reason)
        {
            if (Only == null)
            {
                run.CaseFail("fq-", reason);
                return;
            }

            foreach (string id in Only)
            {
                run.CaseFail(id, reason);
            }
        }

        private static HashSet<string> Pick(string filter, string prefix)
        {
            if (string.IsNullOrWhiteSpace(filter)
                || filter.Equals(prefix, StringComparison.OrdinalIgnoreCase)
                || filter.Equals(prefix + "-", StringComparison.OrdinalIgnoreCase))
            {
                return null;
            }

            var picked = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (string part in filter.Split(','))
            {
                string id = part.Trim();
                if (id.Length > 0)
                {
                    picked.Add(id);
                }
            }

            return picked.Count > 0 ? picked : null;
        }
    }
}
