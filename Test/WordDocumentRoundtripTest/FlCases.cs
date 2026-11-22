using System;
using System.Collections.Generic;
using Newtonsoft.Json.Linq;
using Word = Microsoft.Office.Interop.Word;
using WordAddIn1;

namespace WordDocumentRoundtripTest
{
    /// <summary>字符格式来源三层。对照 docs/测试/word测试/字符格式来源三层。前缀 fl-。</summary>
    internal static class FlCases
    {
        public const int Target = 35;

        private static HashSet<string> Only;

        public static void Run(TestRun run, string host, string filter)
        {
            Only = Pick(filter, "fl");
            if (Only != null)
            {
                run.TotalTarget = Only.Count;
            }

            var session = new WdSession(host);
            RunChar(run, session);
            RunPara(run, session);
            RunCodes(run, session);
            RunApply(run, session);
            RunStyle(run, session);
            RunClear(run, session);
            RunSide(run, session);
        }

        private static void RunChar(TestRun run, WdSession session)
        {
            Case(run, "fl-char-run", () =>
            {
                using (FormatProbe doc = Open(session, "直接加粗两字。"))
                {
                    string err = ReadSent(doc);
                    if (err != null)
                    {
                        return err;
                    }

                    if (!session.PaintNeedle(doc.Path, "加粗", 2, Font("Bold", 1), null))
                    {
                        return "铺加粗失败";
                    }

                    JObject run0 = RunContaining(session, doc, doc.Sentence("加粗"), "加粗");
                    JObject layers = Obj(run0, "format_layers");
                    JObject direct = Obj(layers, "direct");
                    JArray style = Arr(layers, "style");
                    return Flag(direct, "is_bold") && style.Count == 0 && Same(Obj(layers, "final"), Obj(run0, "char_format"))
                        ? null
                        : Dump(layers);
                }
            });
            Case(run, "fl-char-para-own", () =>
            {
                using (FormatProbe doc = Open(session, "段样式字号。"))
                {
                    string err = ReadSent(doc);
                    if (err != null)
                    {
                        return err;
                    }

                    if (!session.AddParagraphStyle(doc.Path, "测字号段", 16, null, null) || !SetStyle(doc, "段样式字号", "测字号段"))
                    {
                        return "建字号样式失败";
                    }

                    JObject run0 = FirstRun(session, doc, doc.Sentence("段样式"));
                    JObject layers = Obj(run0, "format_layers");
                    JObject direct = Obj(layers, "direct");
                    JArray style = Arr(layers, "style");
                    bool size = Is16(Txt(direct, "font_size")) && Is16(Txt(Obj(layers, "final"), "font_size"));
                    bool link = style.Count == 0
                        || (Txt(style[0] as JObject, "link") == "paragraph"
                            && Txt(style[0] as JObject, "name").IndexOf("测字号段", StringComparison.Ordinal) >= 0);
                    return size && link ? null : Dump(layers);
                }
            });
            Case(run, "fl-char-link-one", () =>
            {
                using (FormatProbe doc = Open(session, "只有段落链接。"))
                {
                    string err = ReadSent(doc);
                    if (err != null)
                    {
                        return err;
                    }

                    if (!SetStyle(doc, "只有段落链接", "标题 1", "Heading 1", "heading 1"))
                    {
                        return "套标题失败";
                    }

                    JObject run0 = FirstRun(session, doc, doc.Sentence("只有段落"));
                    JArray style = Arr(Obj(run0, "format_layers"), "style");
                    if (style.Count < 1)
                    {
                        return Dump(run0);
                    }

                    JObject first = style[0] as JObject;
                    return Txt(first, "link") == "paragraph" && Txt(first, "name").IndexOf("标题", StringComparison.Ordinal) >= 0
                        ? null
                        : Dump(style);
                }
            });
            Case(run, "fl-char-link-two", () =>
            {
                using (FormatProbe doc = Open(session, "段落链接再超链接。"))
                {
                    string err = ReadSent(doc);
                    if (err != null)
                    {
                        return err;
                    }

                    if (!SetStyle(doc, "段落链接", "标题 1", "Heading 1", "heading 1")
                        || !SetRangeStyle(doc, "超链接", "超链接", "Hyperlink"))
                    {
                        return "套双链接失败";
                    }

                    JObject run0 = RunContaining(session, doc, doc.Sentence("段落链接"), "超链接");
                    JArray style = Arr(Obj(run0, "format_layers"), "style");
                    if (style.Count < 2)
                    {
                        return Dump(style);
                    }

                    return Txt(style[0], "link") == "paragraph" && Txt(style[1], "link") == "character"
                        && Txt(style[1], "name").IndexOf("超链接", StringComparison.Ordinal) >= 0
                        ? null
                        : Dump(style);
                }
            });
            Case(run, "fl-char-cover", () =>
            {
                using (FormatProbe doc = Open(session, "链接加粗再关掉。"))
                {
                    string err = ReadSent(doc);
                    if (err != null)
                    {
                        return err;
                    }

                    if (!session.AddParagraphStyle(doc.Path, "测加粗链接", null, true, null)
                        || !SetStyle(doc, "链接加粗", "测加粗链接")
                        || !session.SetTextChar(doc.Path, "关掉", "", 0, false, false))
                    {
                        return "铺关掉失败";
                    }

                    JObject run0 = RunContaining(session, doc, doc.Sentence("链接加粗"), "关掉");
                    JObject layers = Obj(run0, "format_layers");
                    JObject direct = Obj(layers, "direct");
                    JObject final = Obj(layers, "final");
                    bool off = direct != null && direct["is_bold"] != null && !Flag(direct, "is_bold")
                        && final != null && !Flag(final, "is_bold");
                    return off ? null : Dump(layers);
                }
            });
            Case(run, "fl-char-keep-ten", () =>
            {
                using (FormatProbe doc = Open(session, "默认正文键。"))
                {
                    string err = ReadSent(doc);
                    if (err != null)
                    {
                        return err;
                    }

                    JObject run0 = FirstRun(session, doc, doc.Sentence("默认正文"));
                    JObject layers = Obj(run0, "format_layers");
                    JObject final = Obj(layers, "final");
                    JObject direct = Obj(layers, "direct");
                    bool kept = final != null && final["font_name"] != null && final["font_size"] != null && final["is_bold"] != null;
                    bool bare = direct == null || direct["is_bold"] == null;
                    return kept && bare ? null : Dump(layers);
                }
            });
            Case(run, "fl-char-no-mark", () =>
            {
                using (FormatProbe doc = Open(session, "标记加粗正文不加。"))
                {
                    string err = ReadSent(doc);
                    if (err != null)
                    {
                        return err;
                    }

                    if (!doc.EditParagraph("标记加粗", para =>
                    {
                        object range = WdSession.ComGet(para, "Range");
                        int end = Convert.ToInt32(WdSession.ComGet(range, "End"));
                        int start = Convert.ToInt32(WdSession.ComGet(range, "Start"));
                        if (end <= start)
                        {
                            return;
                        }

                        WdSession.ComSet(range, "Start", end - 1);
                        object font = WdSession.ComGet(range, "Font");
                        WdSession.ComSet(font, "Bold", 1);
                    }))
                    {
                        return "铺段落标记失败";
                    }

                    JObject run0 = FirstRun(session, doc, doc.Sentence("标记加粗"));
                    JObject layers = Obj(run0, "format_layers");
                    bool bold = Flag(Obj(layers, "direct"), "is_bold") || Flag(Obj(layers, "final"), "is_bold");
                    return bold ? Dump(layers) : null;
                }
            });
            Case(run, "fl-char-final-eq", () =>
            {
                using (FormatProbe doc = Open(session, "最终相等。"))
                {
                    string err = ReadSent(doc);
                    if (err != null)
                    {
                        return err;
                    }

                    JObject run0 = FirstRun(session, doc, doc.Sentence("最终相等"));
                    return Same(Obj(Obj(run0, "format_layers"), "final"), Obj(run0, "char_format")) ? null : Dump(run0);
                }
            });
        }

        private static void RunPara(TestRun run, WdSession session)
        {
            Case(run, "fl-para-direct", () =>
            {
                using (FormatProbe doc = Open(session, "直接左对齐。"))
                {
                    string err = ReadPara(doc);
                    if (err != null)
                    {
                        return err;
                    }

                    if (!SetStyle(doc, "直接左对齐", "标题 1", "Heading 1", "heading 1")
                        || !session.SetParagraphAlign(doc.Path, 1, 1))
                    {
                        return "铺居中失败";
                    }

                    JObject layers = ParaLayers(session, doc, doc.Paragraph("直接左对齐"));
                    JObject direct = Obj(layers, "direct");
                    return Txt(direct, "alignment") == "center" && (direct == null || direct["paragraph_style"] == null)
                        ? null
                        : Dump(layers);
                }
            });
            Case(run, "fl-para-style", () =>
            {
                using (FormatProbe doc = Open(session, "只要标题样式。"))
                {
                    string err = ReadPara(doc);
                    if (err != null)
                    {
                        return err;
                    }

                    if (!SetStyle(doc, "只要标题", "标题 1", "Heading 1", "heading 1"))
                    {
                        return "套标题失败";
                    }

                    JObject style = Obj(ParaLayers(session, doc, doc.Paragraph("只要标题")), "style");
                    return Txt(style, "paragraph_style").IndexOf("标题", StringComparison.Ordinal) >= 0 ? null : Dump(style);
                }
            });
            Case(run, "fl-para-body", () =>
            {
                using (FormatProbe doc = Open(session, "没写样式名。"))
                {
                    string err = ReadPara(doc);
                    if (err != null)
                    {
                        return err;
                    }

                    JObject style = Obj(ParaLayers(session, doc, doc.Paragraph("没写样式")), "style");
                    return Txt(style, "paragraph_style") == "正文" ? null : Dump(style);
                }
            });
            Case(run, "fl-para-cover", () =>
            {
                using (FormatProbe doc = Open(session, "样式居中段上改左。"))
                {
                    if (!doc.Ready)
                    {
                        return doc.Id;
                    }

                    if (!session.AddParagraphStyle(doc.Path, "测居中段层", null, null, 1)
                        || !SetStyle(doc, "样式居中", "测居中段层")
                        || !session.SetParagraphAlign(doc.Path, 1, 0)
                        || !doc.ReadParagraphs())
                    {
                        return "铺覆盖失败";
                    }

                    JObject layers = ParaLayers(session, doc, doc.Paragraph("样式居中"));
                    return Txt(Obj(layers, "direct"), "alignment") == "left"
                        && Txt(Obj(layers, "style"), "alignment") == "center"
                        && Txt(Obj(layers, "final"), "alignment") != "center"
                        ? null
                        : Dump(layers);
                }
            });
            Case(run, "fl-para-final-eq", () =>
            {
                using (FormatProbe doc = Open(session, "段落最终相等。"))
                {
                    string err = ReadPara(doc);
                    if (err != null)
                    {
                        return err;
                    }

                    JObject entry = ParaEntry(session, doc, doc.Paragraph("段落最终"));
                    return Same(Obj(Obj(entry, "format_layers"), "final"), Obj(entry, "para_format")) ? null : Dump(entry);
                }
            });
        }

        private static void RunCodes(TestRun run, WdSession session)
        {
            Case(run, "fl-read-p", () =>
            {
                using (FormatProbe doc = Open(session, "按段回段。"))
                {
                    string err = ReadPara(doc);
                    if (err != null)
                    {
                        return err;
                    }

                    string code = doc.Paragraph("按段回段");
                    JObject data = FormatProbe.Data(session.GetFormatContext(Args(doc.Id, "target_paragraph_codes", code)));
                    string blob = data.ToString();
                    return blob.IndexOf(code, StringComparison.Ordinal) >= 0
                        && blob.IndexOf("sentence_codes_in_paragraph", StringComparison.Ordinal) < 0
                        ? null
                        : blob;
                }
            });
            Case(run, "fl-read-s", () =>
            {
                using (FormatProbe doc = Open(session, "按句回句。"))
                {
                    string err = ReadSent(doc);
                    if (err != null)
                    {
                        return err;
                    }

                    string code = doc.Sentence("按句回句");
                    JObject data = FormatProbe.Data(session.GetFormatContext(Args(doc.Id, "target_codes", code)));
                    JObject entry = Entry(data, "code", code);
                    return entry != null && entry["paragraph_code"] == null && entry["para_format"] != null
                        ? null
                        : data.ToString();
                }
            });
        }

        private static void RunApply(TestRun run, WdSession session)
        {
            Case(run, "fl-pf-s-expand", () =>
            {
                using (FormatProbe doc = Open(session, "前半句。后半句。"))
                {
                    string err = ReadSent(doc);
                    if (err != null)
                    {
                        return err;
                    }

                    ToolResult applied = session.ApplyParagraphFormat(ParaArgs(doc.Id, "target_codes", doc.Sentence("前半句"), "center"));
                    int align = AlignOf(doc, "后半句");
                    return applied.Success && align == 1 ? null : (applied.Error ?? ("对齐 " + align));
                }
            });
            Case(run, "fl-pf-s-dedupe", () =>
            {
                using (FormatProbe doc = Open(session, "同段甲句。同段乙句。"))
                {
                    string err = ReadSent(doc);
                    if (err != null)
                    {
                        return err;
                    }

                    string codes = doc.Sentence("同段甲句") + "," + doc.Sentence("同段乙句");
                    ToolResult applied = session.ApplyParagraphFormat(ParaArgs(doc.Id, "target_codes", codes, "center"));
                    string blob = FormatProbe.Data(applied).ToString() + (applied.Error ?? "");
                    return applied.Success && blob.IndexOf("deprecated_target_codes_resolved_to_paragraph", StringComparison.Ordinal) < 0
                        && AlignOf(doc, "同段甲句") == 1
                        ? null
                        : blob;
                }
            });
            Case(run, "fl-pf-s-echo", () =>
            {
                using (FormatProbe doc = Open(session, "回包仍是句。"))
                {
                    string err = ReadSent(doc);
                    if (err != null)
                    {
                        return err;
                    }

                    string code = doc.Sentence("回包仍是句");
                    JObject data = FormatProbe.Data(session.ApplyParagraphFormat(ParaArgs(doc.Id, "target_codes", code, "center")));
                    string blob = data.ToString();
                    return blob.IndexOf(code, StringComparison.Ordinal) >= 0
                        && blob.IndexOf("sentence_codes_in_paragraph", StringComparison.Ordinal) < 0
                        && blob.IndexOf("\"P_", StringComparison.Ordinal) < 0
                        ? null
                        : blob;
                }
            });
            Case(run, "fl-pf-p-echo", () =>
            {
                using (FormatProbe doc = Open(session, "回包仍是段。"))
                {
                    string err = ReadPara(doc);
                    if (err != null)
                    {
                        return err;
                    }

                    string code = doc.Paragraph("回包仍是段");
                    JObject data = FormatProbe.Data(session.ApplyParagraphFormat(ParaArgs(doc.Id, "target_paragraph_codes", code, "center")));
                    string blob = data.ToString();
                    return blob.IndexOf(code, StringComparison.Ordinal) >= 0
                        && blob.IndexOf("sentence_codes_in_paragraph", StringComparison.Ordinal) < 0
                        ? null
                        : blob;
                }
            });
            Case(run, "fl-pf-s-two", () =>
            {
                using (FormatProbe doc = Open(session, "第一段句。\r第二段句。"))
                {
                    string err = ReadSent(doc);
                    if (err != null)
                    {
                        return err;
                    }

                    string codes = doc.Sentence("第一段句") + "," + doc.Sentence("第二段句");
                    ToolResult applied = session.ApplyParagraphFormat(ParaArgs(doc.Id, "target_codes", codes, "center"));
                    return applied.Success && AlignOf(doc, "第一段句") == 1 && AlignOf(doc, "第二段句") == 1
                        ? null
                        : (applied.Error ?? "两段没有都居中");
                }
            });
            Case(run, "fl-pf-s-glue", () =>
            {
                using (FormatProbe doc = Open(session, "黏连前句。"))
                {
                    if (!doc.Ready || !doc.Touch(InsertPageAfter) || !doc.ReadSentences())
                    {
                        return doc.Ready ? "分页失败" : doc.Id;
                    }

                    string code = doc.Sentence("黏连前句");
                    ToolResult applied = session.ApplyParagraphFormat(ParaArgs(doc.Id, "target_codes", code, "center"));
                    int second = AlignAt(doc, 2);
                    string blob = FormatProbe.Data(applied).ToString();
                    return applied.Success && second == 1 && blob.IndexOf(code, StringComparison.Ordinal) >= 0
                        ? null
                        : (applied.Error ?? ("第二段对齐 " + second + " " + blob));
                }
            });
        }

        private static void RunStyle(TestRun run, WdSession session)
        {
            Case(run, "fl-st-both-half", () =>
            {
                using (FormatProbe doc = Open(session, "半截好后面。"))
                {
                    string err = ReadSent(doc);
                    if (err != null)
                    {
                        return err;
                    }

                    JObject data = Act(session, doc, "apply", doc.Sentence("半截好"), "heading_1", "好");
                    string para = ParaStyle(doc, "半截好");
                    string good = CharStyle(doc, "好");
                    string rest = CharStyle(doc, "后面");
                    bool heading = para.IndexOf("标题", StringComparison.Ordinal) >= 0 || para.IndexOf("Heading", StringComparison.OrdinalIgnoreCase) >= 0;
                    bool only = good.IndexOf("标题", StringComparison.Ordinal) >= 0 && rest.IndexOf("标题", StringComparison.Ordinal) < 0;
                    return FlagOk(data) && heading && only ? null : para + " / " + good + " / " + rest + " " + data;
                }
            });
            Case(run, "fl-st-both-whole", () =>
            {
                using (FormatProbe doc = Open(session, "整段标题句。"))
                {
                    string err = ReadSent(doc);
                    if (err != null)
                    {
                        return err;
                    }

                    JObject data = Act(session, doc, "apply", doc.Sentence("整段标题"), "heading_1", null);
                    string para = ParaStyle(doc, "整段标题");
                    string chars = CharStyle(doc, "整段标题");
                    bool heading = para.IndexOf("标题", StringComparison.Ordinal) >= 0 || para.IndexOf("Heading", StringComparison.OrdinalIgnoreCase) >= 0;
                    return FlagOk(data) && heading && chars.Length > 0 ? null : para + " / " + chars + " " + data;
                }
            });
            Case(run, "fl-st-para-half", () =>
            {
                using (FormatProbe doc = Open(session, "列表几字还在。"))
                {
                    string err = ReadSent(doc);
                    if (err != null)
                    {
                        return err;
                    }

                    if (!SetRangeStyle(doc, "还在", "标题 1 字符", "Heading 1 Char", "heading 1 Char"))
                    {
                        return "铺字符样式失败";
                    }

                    JObject data = Act(session, doc, "apply", doc.Sentence("列表几"), "列表段落", "几");
                    string para = ParaStyle(doc, "列表几");
                    string kept = CharStyle(doc, "还在");
                    bool list = para.IndexOf("列表", StringComparison.Ordinal) >= 0 || para.IndexOf("List", StringComparison.OrdinalIgnoreCase) >= 0;
                    return FlagOk(data) && list && kept.IndexOf("标题", StringComparison.Ordinal) >= 0 ? null : para + " / " + kept + " " + data;
                }
            });
            Case(run, "fl-st-char-half", () =>
            {
                using (FormatProbe doc = Open(session, "只套好字。"))
                {
                    string err = ReadSent(doc);
                    if (err != null)
                    {
                        return err;
                    }

                    if (!MaterializeStyle(doc, "超链接", "Hyperlink"))
                    {
                        return "稿里没有超链接样式";
                    }

                    JObject data = Act(session, doc, "apply", doc.Sentence("只套好"), "超链接", "好");
                    string para = ParaStyle(doc, "只套好");
                    string good = CharStyle(doc, "好");
                    bool body = para.IndexOf("正文", StringComparison.Ordinal) >= 0 || para.IndexOf("Normal", StringComparison.OrdinalIgnoreCase) >= 0;
                    return FlagOk(data) && body && good.IndexOf("超链接", StringComparison.Ordinal) >= 0 ? null : para + " / " + good + " " + data;
                }
            });
            Case(run, "fl-st-char-whole", () =>
            {
                using (FormatProbe doc = Open(session, "整段超链接。"))
                {
                    string err = ReadSent(doc);
                    if (err != null)
                    {
                        return err;
                    }

                    if (!MaterializeStyle(doc, "超链接", "Hyperlink"))
                    {
                        return "稿里没有超链接样式";
                    }

                    JObject data = Act(session, doc, "apply", doc.Sentence("整段超链接"), "超链接", null);
                    string para = ParaStyle(doc, "整段超链接");
                    string chars = CharStyle(doc, "整段超链接");
                    bool body = para.IndexOf("正文", StringComparison.Ordinal) >= 0 || para.IndexOf("Normal", StringComparison.OrdinalIgnoreCase) >= 0;
                    return FlagOk(data) && body && chars.IndexOf("超链接", StringComparison.Ordinal) >= 0
                        ? null
                        : para + " / " + chars + " " + data;
                }
            });
            Case(run, "fl-st-multi", () =>
            {
                using (FormatProbe doc = Open(session, "前句不动。尾句要套。\r头句要套。"))
                {
                    string err = ReadSent(doc);
                    if (err != null)
                    {
                        return err;
                    }

                    string codes = doc.Sentence("尾句要套") + "," + doc.Sentence("头句要套");
                    JObject data = Act(session, doc, "apply", codes, "heading_1", null);
                    string first = ParaStyle(doc, "前句不动");
                    string second = ParaStyle(doc, "头句要套");
                    string headChar = CharStyle(doc, "前句不动");
                    string tailChar = CharStyle(doc, "尾句要套");
                    bool both = (first.IndexOf("标题", StringComparison.Ordinal) >= 0 || first.IndexOf("Heading", StringComparison.OrdinalIgnoreCase) >= 0)
                        && (second.IndexOf("标题", StringComparison.Ordinal) >= 0 || second.IndexOf("Heading", StringComparison.OrdinalIgnoreCase) >= 0);
                    bool chars = tailChar.IndexOf("标题", StringComparison.Ordinal) >= 0 && headChar.IndexOf("标题", StringComparison.Ordinal) < 0;
                    return FlagOk(data) && both && chars ? null : first + " / " + second + " / " + headChar + " / " + tailChar + " " + data;
                }
            });
        }

        private static void RunClear(TestRun run, WdSession session)
        {
            Case(run, "fl-cl-codes", () => ClearFails(session, "定位清甲。", "target_codes"));
            Case(run, "fl-cl-items", () =>
            {
                using (FormatProbe doc = Open(session, "条目清甲。"))
                {
                    string err = ReadSent(doc);
                    if (err != null)
                    {
                        return err;
                    }

                    if (!SetStyle(doc, "条目清甲", "标题 1", "Heading 1", "heading 1"))
                    {
                        return "套标题失败";
                    }

                    var item = new Dictionary<string, object> { ["target_codes"] = doc.Sentence("条目清甲") };
                    var args = new Dictionary<string, object>
                    {
                        ["channel_id"] = doc.Id,
                        ["action"] = "clear",
                        ["items"] = new List<Dictionary<string, object>> { item }
                    };
                    ToolResult result = session.DocumentStructure(args);
                    string style = ParaStyle(doc, "条目清甲");
                    return result != null && !result.Success
                        && (style.IndexOf("标题", StringComparison.Ordinal) >= 0 || style.IndexOf("Heading", StringComparison.OrdinalIgnoreCase) >= 0)
                        ? null
                        : (result == null ? "无回包" : result.Error);
                }
            });
            Case(run, "fl-cl-text", () =>
            {
                using (FormatProbe doc = Open(session, "带字清甲。"))
                {
                    string err = ReadSent(doc);
                    if (err != null)
                    {
                        return err;
                    }

                    if (!SetStyle(doc, "带字清甲", "标题 1", "Heading 1", "heading 1"))
                    {
                        return "套标题失败";
                    }

                    var args = Args(doc.Id, "target_codes", doc.Sentence("带字清甲"));
                    args["action"] = "clear";
                    args["text"] = "带字";
                    ToolResult result = session.DocumentStructure(args);
                    string style = ParaStyle(doc, "带字清甲");
                    return result != null && !result.Success
                        && (style.IndexOf("标题", StringComparison.Ordinal) >= 0 || style.IndexOf("Heading", StringComparison.OrdinalIgnoreCase) >= 0)
                        ? null
                        : (result == null ? "无回包" : result.Error);
                }
            });
            Case(run, "fl-cl-all", () =>
            {
                using (FormatProbe doc = Open(session, "全文清甲。"))
                {
                    if (!doc.Ready || !doc.Session.ApplyHeaderStyle(doc.Path, "标题 1", "Heading 1", "heading 1") || !SetStyle(doc, "全文清甲", "标题 1", "Heading 1", "heading 1"))
                    {
                        return doc.Ready ? "铺全文清失败" : doc.Id;
                    }

                    var args = new Dictionary<string, object> { ["channel_id"] = doc.Id, ["action"] = "clear", ["all"] = true };
                    ToolResult result = session.DocumentStructure(args);
                    string body = ParaStyle(doc, "全文清甲");
                    string header = HeaderStyle(doc);
                    bool cleared = body.IndexOf("正文", StringComparison.Ordinal) >= 0 || body.IndexOf("Normal", StringComparison.OrdinalIgnoreCase) >= 0;
                    bool headerKept = header.IndexOf("标题", StringComparison.Ordinal) >= 0 || header.IndexOf("Heading", StringComparison.OrdinalIgnoreCase) >= 0;
                    return result != null && result.Success && cleared && headerKept ? null : body + " / " + header + " " + (result == null ? "" : result.Error);
                }
            });
            Case(run, "fl-cl-all-locate", () =>
            {
                using (FormatProbe doc = Open(session, "全文带定位。"))
                {
                    string err = ReadSent(doc);
                    if (err != null)
                    {
                        return err;
                    }

                    string before = doc.Snapshot();
                    var args = Args(doc.Id, "target_codes", doc.Sentence("全文带定位"));
                    args["action"] = "clear";
                    args["all"] = true;
                    ToolResult result = session.DocumentStructure(args);
                    return result != null && !result.Success && doc.Snapshot() == before ? null : (result == null ? "无回包" : result.Error);
                }
            });
        }

        private static void RunSide(TestRun run, WdSession session)
        {
            Case(run, "fl-cf-direct", () =>
            {
                using (FormatProbe doc = Open(session, "只写直接加粗。"))
                {
                    string err = ReadSent(doc);
                    if (err != null)
                    {
                        return err;
                    }

                    string before = ParaStyle(doc, "只写直接");
                    var format = new Dictionary<string, object> { ["is_bold"] = true };
                    var args = Args(doc.Id, "target_codes", doc.Sentence("只写直接"));
                    args["text"] = "加粗";
                    args["char_format"] = format;
                    ToolResult applied = session.ApplyDocumentFormat(args);
                    bool bold;
                    string font;
                    bool under;
                    session.TryTextChar(doc.Path, "加粗", out font, out bold, out under);
                    return applied.Success && bold && ParaStyle(doc, "只写直接") == before ? null : (applied.Error ?? "没加粗或段样式变了");
                }
            });
            Case(run, "fl-kb-echo", () =>
            {
                using (FormatProbe doc = Open(session, "知识库回包段。"))
                {
                    string err = ReadPara(doc);
                    if (err != null)
                    {
                        return err;
                    }

                    var format = new Dictionary<string, object> { ["alignment"] = "center" };
                    var args = Args(doc.Id, "target_paragraph_codes", doc.Paragraph("知识库回包"));
                    args["para_format"] = format;
                    ToolResult applied = session.ApplySourceParagraphFormat(args);
                    string blob = FormatProbe.Data(applied).ToString();
                    return applied.Success && blob.IndexOf("affected_paragraph_codes", StringComparison.Ordinal) >= 0
                        && blob.IndexOf("sentence_codes_in_paragraph", StringComparison.Ordinal) < 0
                        ? null
                        : (applied.Error ?? blob);
                }
            });
            Case(run, "fl-display-same", () =>
            {
                using (FormatProbe doc = Open(session, "聚类不带层。"))
                {
                    if (!doc.Ready || !doc.ReadSentences())
                    {
                        return doc.Ready ? "没有句子码" : doc.Id;
                    }

                    ToolResult shown = session.GetFormatDisplay(new Dictionary<string, object> { ["source_channel_id"] = doc.Id });
                    string blob = FormatProbe.Data(shown).ToString();
                    return shown.Success && blob.IndexOf("format_layers", StringComparison.Ordinal) < 0
                        && (blob.IndexOf("entries", StringComparison.Ordinal) >= 0 || blob.IndexOf("para_format_clusters", StringComparison.Ordinal) >= 0)
                        ? null
                        : (shown.Error ?? blob);
                }
            });
        }

        private static string ClearFails(WdSession session, string text, string field)
        {
            using (FormatProbe doc = Open(session, text))
            {
                string err = ReadSent(doc);
                if (err != null)
                {
                    return err;
                }

                if (!SetStyle(doc, text.Replace("。", ""), "标题 1", "Heading 1", "heading 1"))
                {
                    return "套标题失败";
                }

                var args = Args(doc.Id, field, doc.Sentence(text.Replace("。", "")));
                args["action"] = "clear";
                ToolResult result = session.DocumentStructure(args);
                string style = ParaStyle(doc, text.Replace("。", ""));
                bool kept = style.IndexOf("标题", StringComparison.Ordinal) >= 0 || style.IndexOf("Heading", StringComparison.OrdinalIgnoreCase) >= 0;
                return result != null && !result.Success && kept ? null : (result == null ? "无回包" : result.Error + " " + style);
            }
        }

        private static bool MaterializeStyle(FormatProbe doc, params string[] names)
        {
            return doc.Session.MaterializeCharacterStyle(doc.Path, names);
        }

        private static bool InsertPageAfter(object doc)
        {
            try
            {
                object content = WdSession.ComGet(doc, "Content");
                int end = Convert.ToInt32(WdSession.ComGet(content, "End"));
                object range = WdSession.ComCall(doc, "Range", end - 1, end - 1);
                WdSession.ComCall(range, "InsertBreak", 7);
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        private static bool SetStyle(FormatProbe doc, string needle, params string[] names)
        {
            return doc.Session.ApplyParagraphStyle(doc.Path, needle, names);
        }

        private static bool SetRangeStyle(FormatProbe doc, string needle, params string[] names)
        {
            return doc.Session.ApplyRangeStyle(doc.Path, needle, names);
        }

        private static string ParaStyle(FormatProbe doc, string needle)
        {
            string name = "";
            doc.EditParagraph(needle, para =>
            {
                try
                {
                    object style = WdSession.ComGet(para, "Style");
                    name = Convert.ToString(WdSession.ComGet(style, "NameLocal")) ?? "";
                }
                catch (Exception)
                {
                    name = "";
                }
            });
            return name;
        }

        private static string HeaderStyle(FormatProbe doc)
        {
            if (doc.Session.IsWps)
            {
                return doc.Session.PrimaryHeaderStyle(doc.Path);
            }

            string name = "";
            doc.Session.OnWordDoc(doc.Path, word =>
            {
                try
                {
                    Word.HeaderFooter header = word.Sections[1].Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary];
                    Word.Style style = header.Range.Paragraphs[1].get_Style() as Word.Style;
                    name = style == null ? "" : (style.NameLocal ?? "");
                }
                catch (Exception)
                {
                    name = "";
                }

                return true;
            });
            return name;
        }

        private static string CharStyle(FormatProbe doc, string needle)
        {
            string name = "";
            doc.Touch(d =>
            {
                object range = Find(d, needle);
                if (range == null)
                {
                    return false;
                }

                try
                {
                    object style = WdSession.ComGet(range, "CharacterStyle");
                    if (style != null)
                    {
                        name = Convert.ToString(WdSession.ComGet(style, "NameLocal")) ?? "";
                    }
                }
                catch (Exception)
                {
                    name = "";
                }

                return true;
            });
            return name ?? "";
        }

        private static int AlignOf(FormatProbe doc, string needle)
        {
            int align = -1;
            doc.EditParagraph(needle, para =>
            {
                try
                {
                    object format = WdSession.ComGet(para, "Format");
                    if (format == null)
                    {
                        format = WdSession.ComGet(para, "ParagraphFormat");
                    }

                    align = Convert.ToInt32(WdSession.ComGet(format, "Alignment"));
                }
                catch (Exception)
                {
                    align = -1;
                }
            });
            return align;
        }

        private static int AlignAt(FormatProbe doc, int index)
        {
            int align = -1;
            doc.Touch(d =>
            {
                try
                {
                    object paras = WdSession.ComGet(d, "Paragraphs");
                    object para = WdSession.ComCall(paras, "Item", index);
                    object format = WdSession.ComGet(para, "Format");
                    align = Convert.ToInt32(WdSession.ComGet(format, "Alignment"));
                    return true;
                }
                catch (Exception)
                {
                    return false;
                }
            });
            return align;
        }

        private static object Find(object doc, string needle)
        {
            try
            {
                object content = WdSession.ComGet(doc, "Content");
                object find = WdSession.ComGet(content, "Find");
                WdSession.ComCall(find, "ClearFormatting");
                object hit = WdSession.ComCall(find, "Execute", needle);
                if (hit is bool && (bool)hit)
                {
                    return WdSession.ComGet(find, "Parent");
                }
            }
            catch (Exception)
            {
            }

            return null;
        }

        private static JObject Act(WdSession session, FormatProbe doc, string action, string codes, string structure, string text)
        {
            var args = Args(doc.Id, "target_codes", codes);
            args["action"] = action;
            args["structure"] = structure;
            if (!string.IsNullOrEmpty(text))
            {
                args["text"] = text;
            }

            ToolResult result = session.DocumentStructure(args);
            JObject data = FormatProbe.Data(result);
            data["_success"] = result != null && result.Success;
            data["_error"] = result == null ? "" : (result.Error ?? "");
            return data;
        }

        private static JObject FirstRun(WdSession session, FormatProbe doc, string code)
        {
            return RunContaining(session, doc, code, null);
        }

        private static JObject RunContaining(WdSession session, FormatProbe doc, string code, string text)
        {
            JObject data = FormatProbe.Data(session.GetFormatContext(Args(doc.Id, "target_codes", code)));
            JArray entries = data["entries"] as JArray;
            if (entries == null)
            {
                return null;
            }

            foreach (JObject entry in entries)
            {
                JArray runs = entry["runs"] as JArray;
                if (runs == null)
                {
                    continue;
                }

                foreach (JObject run in runs)
                {
                    if (text == null || Txt(run, "text").IndexOf(text, StringComparison.Ordinal) >= 0)
                    {
                        return run;
                    }
                }
            }

            return null;
        }

        private static JObject ParaLayers(WdSession session, FormatProbe doc, string code)
        {
            return Obj(ParaEntry(session, doc, code), "format_layers");
        }

        private static JObject ParaEntry(WdSession session, FormatProbe doc, string code)
        {
            JObject data = FormatProbe.Data(session.GetFormatContext(Args(doc.Id, "target_paragraph_codes", code)));
            return Entry(data, "paragraph_code", code);
        }

        private static JObject Entry(JObject data, string key, string code)
        {
            JArray entries = data["entries"] as JArray;
            if (entries == null)
            {
                return null;
            }

            foreach (JObject entry in entries)
            {
                if (Txt(entry, key) == code)
                {
                    return entry;
                }
            }

            return entries.Count > 0 ? entries[0] as JObject : null;
        }

        private static string ReadSent(FormatProbe doc)
        {
            if (!doc.Ready)
            {
                return doc.Id;
            }

            return doc.ReadSentences() ? null : "没有句子码";
        }

        private static string ReadPara(FormatProbe doc)
        {
            if (!doc.Ready)
            {
                return doc.Id;
            }

            return doc.ReadParagraphs() ? null : "没有段落码";
        }

        private static Dictionary<string, object> Args(string id, string key, string value)
        {
            return new Dictionary<string, object> { ["channel_id"] = id, [key] = value };
        }

        private static Dictionary<string, object> ParaArgs(string id, string key, string value, string alignment)
        {
            Dictionary<string, object> args = Args(id, key, value);
            args["para_format"] = new Dictionary<string, object> { ["alignment"] = alignment };
            return args;
        }

        private static Dictionary<string, object> Font(string key, object value)
        {
            return new Dictionary<string, object> { [key] = value };
        }

        private static bool FlagOk(JObject data)
        {
            return data != null && data["_success"] != null && data["_success"].Value<bool>();
        }

        private static bool Same(JObject left, JObject right)
        {
            if (left == null || right == null)
            {
                return false;
            }

            return JToken.DeepEquals(left, right);
        }

        private static JObject Obj(JObject data, string key)
        {
            return data == null ? null : data[key] as JObject;
        }

        private static JArray Arr(JObject data, string key)
        {
            return data == null ? new JArray() : (data[key] as JArray ?? new JArray());
        }

        private static string Txt(JToken token, string key)
        {
            return Txt(token as JObject, key);
        }

        private static string Txt(JObject data, string key)
        {
            return data == null || data[key] == null ? "" : data[key].ToString();
        }

        private static bool Is16(string size)
        {
            return size == "16" || size == "16.0";
        }

        private static bool Flag(JObject data, string key)
        {
            bool bit;
            return data != null && data[key] != null && bool.TryParse(data[key].ToString(), out bit) && bit;
        }

        private static string Dump(JToken token)
        {
            return token == null ? "空" : token.ToString();
        }

        private static FormatProbe Open(WdSession session, string text)
        {
            return FormatProbe.Open(session, text, "");
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
