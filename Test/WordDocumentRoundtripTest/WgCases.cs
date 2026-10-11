using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using WordAddIn1;

namespace WordDocumentRoundtripTest
{
    /// <summary>
    /// 第 5 波：读格式上下文、套字符 / 段版式、页边距。铺稿用 COM，抄当次 S_/P_/T_。
    /// </summary>
    internal static class WgCases
    {
        public const int Target = 44;

        private const string SentA = "易写套格式甲句。";
        private const string SentB = "易写套格式乙句。";
        private const string Dup = "重复套格式夹具。";
        private const int AlignCenter = 1;
        private const double PageLeft = 120;

        public static void Run(
            TestRun run,
            string host,
            int wordBefore,
            int wpsBefore,
            bool wordWinBefore,
            bool wpsWinBefore)
        {
            var s = new WdSession(host);
            RunNegEarly(run, s);
            string warmPath = WarmWpsChannel(s);
            try
            {
                RunContextAndChar(run, s);
                RunCopyAndPara(run, s);
                RunTable(run, s);
                RunAmbiguity(run, s);
                RunPage(run, s);
                RunNegLate(run, s);
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
            s.TypePlain(path, SentA);
            s.GetContent(id);
            string code = FirstCode(ReadBody(s, id), "S");
            if (!string.IsNullOrEmpty(code))
            {
                s.GetFormatContext(CtxSent(id, code));
            }

            return path;
        }

        private static void BurnFirstWpsOpen(WdSession s)
        {
            if (!s.IsWps)
            {
                return;
            }

            string path = s.NewDocxPath();
            ToolResult opened = s.Open(path, true, s.App);
            if (!opened.Success)
            {
                return;
            }

            string id = WdSession.Str(opened, "channel_id");
            s.ActivateDoc(path);
            s.TypePlain(path, SentA);
            s.GetContent(id);
            string code = FirstCode(ReadBody(s, id), "S");
            if (!string.IsNullOrEmpty(code))
            {
                s.ApplyDocumentFormat(CharArgs(id, code, BoldUnderline()));
            }

            s.CloseDocByPath(path);
        }

        private static void RunNegEarly(TestRun run, WdSession s)
        {
            var dummyChar = CharArgs(null, "S_00001", BoldUnderline());
            ToolResult noCh = s.ApplyDocumentFormat(dummyChar);
            if (!noCh.Success
                && (Contains(noCh.Error, "打开")
                    || Contains(noCh.Error, "unsupported")
                    || Contains(noCh.Error, "browser")
                    || Contains(noCh.Error, "渠道")))
            {
                run.CaseOk("wg-neg-no-channel");
            }
            else if (noCh.Success)
            {
                run.CaseSkip("wg-neg-no-channel", "现场已有默认文字渠，不关用户稿");
            }
            else
            {
                run.CaseFail("wg-neg-no-channel", noCh.Error ?? "应失败或跳过");
            }

            ExpectFail(run, "wg-neg-unknown-id", s.ApplyDocumentFormat(CharArgs("w999", "S_00001", BoldUnderline())), "未知");

            ExpectFail(
                run,
                "wg-neg-char-no-target",
                s.ApplyDocumentFormat(new Dictionary<string, object> { ["char_format"] = BoldUnderline() }),
                "target");
            ExpectFail(
                run,
                "wg-neg-char-mutex",
                s.ApplyDocumentFormat(new Dictionary<string, object>
                {
                    ["target_codes"] = "S_00001",
                    ["target_paragraph_codes"] = "P_00001",
                    ["char_format"] = BoldUnderline()
                }),
                "互斥");
            ExpectFail(
                run,
                "wg-neg-char-no-fmt",
                s.ApplyDocumentFormat(new Dictionary<string, object> { ["target_codes"] = "S_00001" }),
                "char_format");
            ExpectFail(
                run,
                "wg-neg-char-empty-fmt",
                s.ApplyDocumentFormat(new Dictionary<string, object>
                {
                    ["target_codes"] = "S_00001",
                    ["char_format"] = new Dictionary<string, object>()
                }),
                "char_format");
            ExpectFail(
                run,
                "wg-neg-char-inherit",
                s.ApplyDocumentFormat(new Dictionary<string, object>
                {
                    ["target_codes"] = "S_00001",
                    ["char_format"] = BoldUnderline(),
                    ["inherit_from"] = "S_00001"
                }),
                "inherit");
            ExpectFail(
                run,
                "wg-neg-char-match-role",
                s.ApplyDocumentFormat(new Dictionary<string, object>
                {
                    ["target_codes"] = "S_00001",
                    ["char_format"] = BoldUnderline(),
                    ["match_role"] = "body"
                }),
                "match_role");
            ExpectFail(
                run,
                "wg-neg-char-t-code",
                s.ApplyDocumentFormat(CharArgs(null, "T_00001", BoldUnderline())),
                "T_");
            ExpectFail(
                run,
                "wg-neg-page-both-off",
                s.ApplyPageSetup(new Dictionary<string, object>
                {
                    ["apply_page_setup"] = false,
                    ["apply_headers_footers"] = false
                }),
                "false");
        }

        private static void RunContextAndChar(TestRun run, WdSession s)
        {
            BurnFirstWpsOpen(s);

            string path, id, text;
            if (!TryPlant(run, s, SentA + "\r" + SentB, out path, out id, out text,
                "wg-neg-ctx-no-target", "wg-neg-ctx-mutex",
                "wg-neg-para-no-fmt", "wg-neg-para-inherit",
                "wg-ctx-sent", "wg-ctx-sent-keys", "wg-ctx-sent-echo",
                "wg-ctx-para", "wg-ctx-para-align", "wg-ctx-para-echo",
                "wg-char-ok", "wg-char-read", "wg-char-text", "wg-char-fields"))
            {
                return;
            }

            if (!s.SetParagraphChar(path, 1, "黑体", 16, true)
                || !s.SetParagraphAlign(path, 1, AlignCenter))
            {
                FailMany(run, "COM 铺黑体/居中失败",
                    "wg-neg-ctx-no-target", "wg-neg-ctx-mutex",
                    "wg-neg-para-no-fmt", "wg-neg-para-inherit",
                    "wg-ctx-sent", "wg-ctx-sent-keys", "wg-ctx-sent-echo",
                    "wg-ctx-para", "wg-ctx-para-align", "wg-ctx-para-echo",
                    "wg-char-ok", "wg-char-read", "wg-char-text", "wg-char-fields");
                s.CloseDocByPath(path);
                return;
            }

            s.GetContent(id);
            text = ReadBody(s, id);
            string sent = FirstCode(text, "S");
            string para = FirstParagraphCode(s, id);
            if (string.IsNullOrEmpty(sent) || string.IsNullOrEmpty(para))
            {
                FailMany(run, "无可抄 S_ / P_",
                    "wg-neg-ctx-no-target", "wg-neg-ctx-mutex",
                    "wg-neg-para-no-fmt", "wg-neg-para-inherit",
                    "wg-ctx-sent", "wg-ctx-sent-keys", "wg-ctx-sent-echo",
                    "wg-ctx-para", "wg-ctx-para-align", "wg-ctx-para-echo",
                    "wg-char-ok", "wg-char-read", "wg-char-text", "wg-char-fields");
                s.CloseDocByPath(path);
                return;
            }

            ExpectFail(
                run,
                "wg-neg-ctx-no-target",
                s.GetFormatContext(new Dictionary<string, object> { ["channel_id"] = id }),
                "target");
            ExpectFail(
                run,
                "wg-neg-ctx-mutex",
                s.GetFormatContext(new Dictionary<string, object>
                {
                    ["channel_id"] = id,
                    ["target_codes"] = sent,
                    ["target_paragraph_codes"] = para
                }),
                "互斥");
            ExpectFail(
                run,
                "wg-neg-para-no-fmt",
                s.ApplyParagraphFormat(new Dictionary<string, object>
                {
                    ["channel_id"] = id,
                    ["target_paragraph_codes"] = para
                }),
                "para_format");
            ExpectFail(
                run,
                "wg-neg-para-inherit",
                s.ApplyParagraphFormat(new Dictionary<string, object>
                {
                    ["channel_id"] = id,
                    ["target_paragraph_codes"] = para,
                    ["para_format"] = CenterPara(),
                    ["inherit_from_paragraph"] = para
                }),
                "inherit");

            ToolResult ctx = s.GetFormatContext(CtxSent(id, sent));
            JObject entry = FirstEntry(ctx);
            JObject cf = Obj(entry, "char_format");
            Ok(run, "wg-ctx-sent", ctx.Success && entry != null, ctx.Success ? "无 entries" : (ctx.Error ?? ""));
            Ok(run, "wg-ctx-sent-keys",
                ctx.Success
                    && FontLooksLike(Txt(cf, "font_name"), "黑体")
                    && Flag(cf, "is_bold"),
                ctx.Success
                    ? "font_name=" + Txt(cf, "font_name") + " bold=" + Flag(cf, "is_bold")
                    : (ctx.Error ?? ""));
            Ok(run, "wg-ctx-sent-echo",
                ctx.Success && Contains(WdSession.Str(ctx, "target_codes"), sent),
                ctx.Success ? "回包无 target_codes" : (ctx.Error ?? ""));

            ToolResult pctx = s.GetFormatContext(CtxPara(id, para));
            JObject pentry = FirstEntry(pctx);
            JObject pf = Obj(pentry, "para_format");
            Ok(run, "wg-ctx-para", pctx.Success && pentry != null, pctx.Success ? "无段落 entries" : (pctx.Error ?? ""));
            Ok(run, "wg-ctx-para-align",
                pctx.Success && string.Equals(Txt(pf, "alignment"), "center", StringComparison.OrdinalIgnoreCase),
                pctx.Success ? "alignment=" + Txt(pf, "alignment") : (pctx.Error ?? ""));
            Ok(run, "wg-ctx-para-echo",
                pctx.Success && Contains(WdSession.Str(pctx, "target_paragraph_codes"), para),
                pctx.Success ? "回包无 target_paragraph_codes" : (pctx.Error ?? ""));

            ToolResult apply = s.ApplyDocumentFormat(CharArgs(id, sent, new Dictionary<string, object>
            {
                ["font_name"] = "楷体",
                ["font_size"] = 14,
                ["is_bold"] = false,
                ["has_underline"] = true
            }));
            Ok(run, "wg-char-ok", apply.Success, apply.Error ?? "套字符失败");
            JObject after = FirstEntry(s.GetFormatContext(CtxSent(id, sent)));
            JObject afterCf = Obj(after, "char_format");
            Ok(run, "wg-char-read",
                apply.Success
                    && FontLooksLike(Txt(afterCf, "font_name"), "楷体")
                    && !Flag(afterCf, "is_bold")
                    && Flag(afterCf, "has_underline"),
                apply.Success
                    ? "回读 font=" + Txt(afterCf, "font_name")
                        + " bold=" + Flag(afterCf, "is_bold")
                        + " ul=" + Flag(afterCf, "has_underline")
                    : (apply.Error ?? ""));
            Ok(run, "wg-char-text",
                apply.Success && StillHas(s, id, SentA) && StillHas(s, id, SentB),
                apply.Success ? "套字符改了正文" : (apply.Error ?? ""));
            Ok(run, "wg-char-fields",
                apply.Success
                    && BlobHas(apply, sent)
                    && BlobHas(apply, "explicit")
                    && !WdSession.HasKey(apply, "snapshot")
                    && !WdSession.HasKey(apply, "display"),
                apply.Success ? "回包缺 applied / explicit，或带了 snapshot" : (apply.Error ?? ""));
            s.CloseDocByPath(path);
        }

        private static void RunCopyAndPara(TestRun run, WdSession s)
        {
            string path, id, text;
            if (!TryPlant(run, s, SentA + "\r" + SentB, out path, out id, out text,
                "wg-char-copy", "wg-char-p", "wg-para-ok", "wg-para-read", "wg-para-text", "wg-para-indent"))
            {
                return;
            }

            if (!s.SetParagraphChar(path, 1, "黑体", 16, true)
                || !s.SetParagraphChar(path, 2, "宋体", 12, false))
            {
                FailMany(run, "COM 铺两段字体失败",
                    "wg-char-copy", "wg-char-p", "wg-para-ok", "wg-para-read", "wg-para-text", "wg-para-indent");
                s.CloseDocByPath(path);
                return;
            }

            s.GetContent(id);
            List<string> sCodes = ListCodes(ReadBody(s, id), "S");
            List<string> pCodes = ListParagraphCodes(s, id);
            if (sCodes.Count < 2 || pCodes.Count < 2)
            {
                FailMany(run, "两句/两段码不够",
                    "wg-char-copy", "wg-char-p", "wg-para-ok", "wg-para-read", "wg-para-text", "wg-para-indent");
                s.CloseDocByPath(path);
                return;
            }

            JObject srcCf = Obj(FirstEntry(s.GetFormatContext(CtxSent(id, sCodes[0]))), "char_format");
            Dictionary<string, object> copied = ToDict(srcCf);
            if (copied.Count == 0)
            {
                FailMany(run, "源句 char_format 空",
                    "wg-char-copy", "wg-char-p", "wg-para-ok", "wg-para-read", "wg-para-text", "wg-para-indent");
                s.CloseDocByPath(path);
                return;
            }

            ToolResult copy = s.ApplyDocumentFormat(CharArgs(id, sCodes[1], copied));
            JObject dst = Obj(FirstEntry(s.GetFormatContext(CtxSent(id, sCodes[1]))), "char_format");
            Ok(run, "wg-char-copy",
                copy.Success
                    && FontLooksLike(Txt(dst, "font_name"), Txt(srcCf, "font_name"))
                    && Flag(dst, "is_bold")
                    && StillHas(s, id, SentA)
                    && StillHas(s, id, SentB),
                copy.Success
                    ? "抄到乙句 font=" + Txt(dst, "font_name") + " bold=" + Flag(dst, "is_bold")
                    : (copy.Error ?? ""));

            ToolResult onP = s.ApplyDocumentFormat(new Dictionary<string, object>
            {
                ["channel_id"] = id,
                ["target_paragraph_codes"] = pCodes[0],
                ["char_format"] = new Dictionary<string, object> { ["has_strikethrough"] = true }
            });
            JObject pChar = Obj(FirstEntry(s.GetFormatContext(CtxSent(id, sCodes[0]))), "char_format");
            Ok(run, "wg-char-p",
                onP.Success && Flag(pChar, "has_strikethrough"),
                onP.Success ? "P_ 套删除线未回读到" : (onP.Error ?? ""));

            ToolResult para = s.ApplyParagraphFormat(new Dictionary<string, object>
            {
                ["channel_id"] = id,
                ["target_paragraph_codes"] = pCodes[1],
                ["para_format"] = CenterPara()
            });
            JObject paraFmt = Obj(FirstEntry(s.GetFormatContext(CtxPara(id, pCodes[1]))), "para_format");
            Ok(run, "wg-para-ok", para.Success, para.Error ?? "套段失败");
            Ok(run, "wg-para-read",
                para.Success && string.Equals(Txt(paraFmt, "alignment"), "center", StringComparison.OrdinalIgnoreCase),
                para.Success ? "alignment=" + Txt(paraFmt, "alignment") : (para.Error ?? ""));
            Ok(run, "wg-para-text",
                para.Success && StillHas(s, id, SentA) && StillHas(s, id, SentB),
                para.Success ? "套段改了正文" : (para.Error ?? ""));

            ToolResult indent = s.ApplyParagraphFormat(new Dictionary<string, object>
            {
                ["channel_id"] = id,
                ["target_paragraph_codes"] = pCodes[0],
                ["para_format"] = new Dictionary<string, object>
                {
                    ["first_line_indent"] = 24,
                    ["space_before"] = 12
                }
            });
            JObject indFmt = Obj(FirstEntry(s.GetFormatContext(CtxPara(id, pCodes[0]))), "para_format");
            Ok(run, "wg-para-indent",
                indent.Success
                    && Near(Num(indFmt, "first_line_indent"), 24, 1)
                    && Near(Num(indFmt, "space_before"), 12, 1)
                    && StillHas(s, id, SentA),
                indent.Success
                    ? "indent=" + Num(indFmt, "first_line_indent") + " before=" + Num(indFmt, "space_before")
                    : (indent.Error ?? ""));
            s.CloseDocByPath(path);
        }

        private static void RunTable(TestRun run, WdSession s)
        {
            string path = s.NewDocxPath();
            ToolResult opened = s.Open(path, true, s.App);
            if (!opened.Success)
            {
                FailMany(run, opened.Error ?? "新建失败",
                    "wg-tbl-ctx", "wg-tbl-char", "wg-tbl-cell", "wg-tbl-wrong-t");
                return;
            }

            string id = WdSession.Str(opened, "channel_id");
            if (!s.InsertTable(path, 2, 2, new[] { "甲", "乙", "丙", "丁" }))
            {
                FailMany(run, "COM 建表失败",
                    "wg-tbl-ctx", "wg-tbl-char", "wg-tbl-cell", "wg-tbl-wrong-t");
                s.CloseDocByPath(path);
                return;
            }

            ToolResult got = s.GetContent(id);
            string body = got.Success ? WdSession.Str(got, "document_content") : "";
            Match marker = Regex.Match(body, @"\[S_([0-9a-zA-Z]{5})-start\]\[in_table=(T_[^\]]+)\]");
            if (!got.Success || !marker.Success)
            {
                FailMany(run, got.Success ? "表内无 [in_table=T_]" : (got.Error ?? "读表失败"),
                    "wg-tbl-ctx", "wg-tbl-char", "wg-tbl-cell", "wg-tbl-wrong-t");
                s.CloseDocByPath(path);
                return;
            }

            string code = "S_" + marker.Groups[1].Value;
            string tableId = marker.Groups[2].Value;
            var ctxArgs = CtxSent(id, code);
            ctxArgs["in_table"] = tableId;
            ToolResult ctx = s.GetFormatContext(ctxArgs);
            Ok(run, "wg-tbl-ctx", ctx.Success && FirstEntry(ctx) != null, ctx.Success ? "表内 context 无 entry" : (ctx.Error ?? ""));

            var applyArgs = CharArgs(id, code, new Dictionary<string, object> { ["is_bold"] = true });
            applyArgs["in_table"] = tableId;
            ToolResult apply = s.ApplyDocumentFormat(applyArgs);
            JObject cellCf = Obj(FirstEntry(s.GetFormatContext(ctxArgs)), "char_format");
            Ok(run, "wg-tbl-char", apply.Success && Flag(cellCf, "is_bold"), apply.Success ? "表内加粗未回读到" : (apply.Error ?? ""));
            Ok(run, "wg-tbl-cell",
                apply.Success && StillHas(s, id, "甲") && StillHas(s, id, "丁"),
                apply.Success ? "套表内格式改了格字" : (apply.Error ?? ""));

            var wrong = CharArgs(id, code, new Dictionary<string, object> { ["has_underline"] = true });
            wrong["in_table"] = "T_zzzzz";
            ToolResult bad = s.ApplyDocumentFormat(wrong);
            JObject afterWrong = Obj(FirstEntry(s.GetFormatContext(ctxArgs)), "char_format");
            Ok(run, "wg-tbl-wrong-t",
                !bad.Success && !Flag(afterWrong, "has_underline") && StillHas(s, id, "甲"),
                bad.Success ? "错 T_ 仍成功" : (bad.Error ?? "错 T_ 却写下划线"));
            s.CloseDocByPath(path);
        }

        private static void RunAmbiguity(TestRun run, WdSession s)
        {
            string path, id, text;
            if (!TryPlant(run, s, "独特套格式甲句。\r" + Dup + "\r" + Dup, out path, out id, out text,
                "wg-amb-fail", "wg-amb-p"))
            {
                return;
            }

            string dupCode = FindDupCode(text);
            if (string.IsNullOrEmpty(dupCode))
            {
                FailMany(run, "未找到重复 S_", "wg-amb-fail", "wg-amb-p");
                s.CloseDocByPath(path);
                return;
            }

            ToolResult amb = s.ApplyDocumentFormat(CharArgs(id, dupCode, new Dictionary<string, object>
            {
                ["has_underline"] = true
            }));
            Ok(run, "wg-amb-fail",
                !amb.Success && StillHas(s, id, Dup),
                amb.Success ? "重复码仍套上" : (amb.Error ?? "夹具被改"));

            List<string> pCodes = ListParagraphCodes(s, id);
            string stamped = "";
            foreach (string code in pCodes)
            {
                if (code.IndexOf('_') != code.LastIndexOf('_'))
                {
                    stamped = code;
                    break;
                }
            }

            if (string.IsNullOrEmpty(stamped))
            {
                FailMany(run, "paragraph display 无 P_xxxxx_N", "wg-amb-p");
                s.CloseDocByPath(path);
                return;
            }

            ToolResult one = s.ApplyDocumentFormat(new Dictionary<string, object>
            {
                ["channel_id"] = id,
                ["target_codes"] = stamped,
                ["char_format"] = new Dictionary<string, object> { ["has_underline"] = true }
            });
            Ok(run, "wg-amb-p",
                one.Success && StillHas(s, id, Dup) && StillHas(s, id, "独特套格式甲句。"),
                one.Success ? "P_xxxxx_N 套完丢了字" : (one.Error ?? ""));
            s.CloseDocByPath(path);
        }

        private static void RunPage(TestRun run, WdSession s)
        {
            string path, id, text;
            if (!TryPlant(run, s, SentA, out path, out id, out text,
                "wg-page-ok", "wg-page-text", "wg-page-margin", "wg-page-fields"))
            {
                return;
            }

            ToolResult page = s.ApplyPageSetup(new Dictionary<string, object>
            {
                ["channel_id"] = id,
                ["apply_page_setup"] = true,
                ["apply_headers_footers"] = false,
                ["left_margin"] = PageLeft
            });
            Ok(run, "wg-page-ok", page.Success, page.Error ?? "页设置失败");
            Ok(run, "wg-page-text",
                page.Success && StillHas(s, id, SentA),
                page.Success ? "改边距动了正文" : (page.Error ?? ""));

            double left;
            bool read = s.TryReadLeftMargin(path, out left);
            Ok(run, "wg-page-margin",
                page.Success && read && Near(left, PageLeft, 3),
                page.Success
                    ? (read ? "LeftMargin=" + left.ToString("F1") : "COM 读边距失败")
                    : (page.Error ?? ""));
            Ok(run, "wg-page-fields",
                page.Success
                    && WdSession.Bool(page, "page_setup_applied")
                    && !WdSession.Bool(page, "apply_headers_footers"),
                page.Success ? "回包缺 page_setup_applied 或仍开了页眉页脚" : (page.Error ?? ""));
            s.CloseDocByPath(path);
        }

        private static void RunNegLate(TestRun run, WdSession s)
        {
            var dummy = CharArgs(null, "S_00001", BoldUnderline());
            string xlsx = Path.Combine(Path.GetTempPath(), "ew-wg-" + Guid.NewGuid().ToString("N") + ".xlsx");
            ToolResult xl = s.Open(xlsx, true, "excel");
            if (!xl.Success)
            {
                run.CaseFail("wg-neg-wrong-kind-excel", xl.Error ?? "建 xlsx 失败");
            }
            else
            {
                dummy["channel_id"] = WdSession.Str(xl, "channel_id");
                ExpectFail(run, "wg-neg-wrong-kind-excel", s.ApplyDocumentFormat(dummy), "excel");
                WdSession.CloseExcelByPath(WdSession.Str(xl, "path"));
            }

            dummy = CharArgs(null, "S_00001", BoldUnderline());
            string pptx = Path.Combine(Path.GetTempPath(), "ew-wg-" + Guid.NewGuid().ToString("N") + ".pptx");
            ToolResult ppt = s.Open(pptx, true, "powerpoint");
            if (!ppt.Success)
            {
                run.CaseFail("wg-neg-wrong-kind-ppt", ppt.Error ?? "建 pptx 失败");
            }
            else
            {
                dummy["channel_id"] = WdSession.Str(ppt, "channel_id");
                ExpectFail(run, "wg-neg-wrong-kind-ppt", s.ApplyDocumentFormat(dummy), "ppt");
                WdSession.ClosePptByPath(WdSession.Str(ppt, "path"));
            }

            string closePath = s.NewDocxPath();
            ToolResult opened = s.Open(closePath, true, s.App);
            if (!opened.Success)
            {
                run.CaseFail("wg-neg-closed", opened.Error ?? "再建稿失败");
                return;
            }

            string cid = WdSession.Str(opened, "channel_id");
            if (!s.CloseDocByPath(closePath))
            {
                run.CaseFail("wg-neg-closed", "COM 关稿失败");
                return;
            }

            ToolResult closed = s.ApplyDocumentFormat(CharArgs(cid, "S_00001", BoldUnderline()));
            if (closed != null && !closed.Success
                && (Contains(closed.Error, "关闭")
                    || Contains(closed.Error, "找不到")
                    || Contains(closed.Error, "未知")
                    || Contains(closed.Error, "unsupported")))
            {
                run.CaseOk("wg-neg-closed");
            }
            else if (closed != null && !closed.Success)
            {
                run.CaseFail("wg-neg-closed", closed.Error ?? "");
            }
            else
            {
                run.CaseFail("wg-neg-closed", "关稿后仍成功");
            }
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
            Ok(run, "wg-no-quit", wordOk && wpsOk, wordDetail + " " + wpsDetail);
        }

        private static bool TryPlant(
            TestRun run,
            WdSession s,
            string plain,
            out string path,
            out string id,
            out string text,
            params string[] failIds)
        {
            path = s.NewDocxPath();
            id = null;
            text = "";
            ToolResult opened = s.Open(path, true, s.App);
            if (!opened.Success)
            {
                FailMany(run, opened.Error ?? "新建失败", failIds);
                return false;
            }

            id = WdSession.Str(opened, "channel_id");
            if (!s.TypePlain(path, plain))
            {
                FailMany(run, "COM 写句失败", failIds);
                s.CloseDocByPath(path);
                return false;
            }

            ToolResult got = s.GetContent(id);
            if (!got.Success)
            {
                FailMany(run, got.Error ?? "读失败", failIds);
                s.CloseDocByPath(path);
                return false;
            }

            text = WdSession.Str(got, "document_content");
            if (string.IsNullOrEmpty(FirstCode(text, "S")))
            {
                FailMany(run, "无可抄 S_", failIds);
                s.CloseDocByPath(path);
                return false;
            }

            return true;
        }

        private static Dictionary<string, object> CtxSent(string channelId, string codes)
        {
            var args = new Dictionary<string, object> { ["target_codes"] = codes ?? "" };
            if (channelId != null)
            {
                args["channel_id"] = channelId;
            }

            return args;
        }

        private static Dictionary<string, object> CtxPara(string channelId, string codes)
        {
            var args = new Dictionary<string, object> { ["target_paragraph_codes"] = codes ?? "" };
            if (channelId != null)
            {
                args["channel_id"] = channelId;
            }

            return args;
        }

        private static Dictionary<string, object> CharArgs(
            string channelId,
            string codes,
            Dictionary<string, object> charFormat)
        {
            var args = new Dictionary<string, object>
            {
                ["target_codes"] = codes ?? "",
                ["char_format"] = charFormat ?? new Dictionary<string, object>()
            };
            if (channelId != null)
            {
                args["channel_id"] = channelId;
            }

            return args;
        }

        private static Dictionary<string, object> BoldUnderline()
        {
            return new Dictionary<string, object>
            {
                ["is_bold"] = true,
                ["has_underline"] = true
            };
        }

        private static Dictionary<string, object> CenterPara()
        {
            return new Dictionary<string, object> { ["alignment"] = "center" };
        }

        private static List<string> ListCodes(string text, string prefix)
        {
            var list = new List<string>();
            if (string.IsNullOrEmpty(text))
            {
                return list;
            }

            foreach (Match m in Regex.Matches(text, @"\[" + prefix + @"_([0-9a-zA-Z]{5})-start\]"))
            {
                list.Add(prefix + "_" + m.Groups[1].Value);
            }

            return list;
        }

        private static string FirstCode(string text, string prefix)
        {
            List<string> list = ListCodes(text, prefix);
            return list.Count == 0 ? "" : list[0];
        }

        private static string FirstParagraphCode(WdSession s, string id)
        {
            List<string> list = ListParagraphCodes(s, id);
            return list.Count == 0 ? "" : list[0];
        }

        private static List<string> ListParagraphCodes(WdSession s, string id)
        {
            ToolResult got = s.GetContent(id, "paragraph");
            string text = got.Success ? WdSession.Str(got, "document_content") : "";
            var list = new List<string>();
            if (string.IsNullOrEmpty(text))
            {
                return list;
            }

            foreach (Match m in Regex.Matches(text, @"\[P_([0-9a-zA-Z]{5}(?:_[1-9][0-9]*)?)(?:-start| start)\]"))
            {
                list.Add("P_" + m.Groups[1].Value);
            }

            return list;
        }

        private static string FindDupCode(string text)
        {
            var seen = new HashSet<string>(StringComparer.Ordinal);
            foreach (string code in ListCodes(text, "S"))
            {
                if (!seen.Add(code))
                {
                    return code;
                }
            }

            return "";
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

        private static JObject DataObj(ToolResult r)
        {
            if (r == null || r.Data == null)
            {
                return null;
            }

            try
            {
                return JObject.Parse(JsonConvert.SerializeObject(r.Data));
            }
            catch
            {
                return null;
            }
        }

        private static JObject FirstEntry(ToolResult r)
        {
            JObject data = DataObj(r);
            JArray entries = data == null ? null : data["entries"] as JArray;
            if (entries == null || entries.Count == 0)
            {
                return null;
            }

            return entries[0] as JObject;
        }

        private static JObject Obj(JObject parent, string key)
        {
            if (parent == null || string.IsNullOrEmpty(key))
            {
                return null;
            }

            return parent[key] as JObject;
        }

        private static string Txt(JObject o, string key)
        {
            if (o == null || o[key] == null || o[key].Type == JTokenType.Null)
            {
                return "";
            }

            return o[key].ToString();
        }

        private static bool Flag(JObject o, string key)
        {
            if (o == null || o[key] == null || o[key].Type == JTokenType.Null)
            {
                return false;
            }

            if (o[key].Type == JTokenType.Boolean)
            {
                return o[key].Value<bool>();
            }

            string v = o[key].ToString();
            return string.Equals(v, "true", StringComparison.OrdinalIgnoreCase) || v == "1";
        }

        private static double Num(JObject o, string key)
        {
            if (o == null || o[key] == null || o[key].Type == JTokenType.Null)
            {
                return 0;
            }

            try
            {
                return o[key].Value<double>();
            }
            catch
            {
                double d;
                return double.TryParse(o[key].ToString(), out d) ? d : 0;
            }
        }

        private static Dictionary<string, object> ToDict(JObject o)
        {
            var dict = new Dictionary<string, object>();
            if (o == null)
            {
                return dict;
            }

            foreach (JProperty p in o.Properties())
            {
                if (p.Value == null || p.Value.Type == JTokenType.Null)
                {
                    continue;
                }

                if (p.Value.Type == JTokenType.Boolean)
                {
                    dict[p.Name] = p.Value.Value<bool>();
                }
                else if (p.Value.Type == JTokenType.Integer || p.Value.Type == JTokenType.Float)
                {
                    dict[p.Name] = p.Value.Value<double>();
                }
                else
                {
                    dict[p.Name] = p.Value.ToString();
                }
            }

            return dict;
        }

        private static bool FontLooksLike(string actual, string expect)
        {
            if (string.IsNullOrEmpty(actual) || string.IsNullOrEmpty(expect))
            {
                return false;
            }

            if (actual.IndexOf(expect, StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return true;
            }

            if (expect.IndexOf("黑", StringComparison.Ordinal) >= 0)
            {
                return actual.IndexOf("黑", StringComparison.Ordinal) >= 0
                    || actual.IndexOf("SimHei", StringComparison.OrdinalIgnoreCase) >= 0;
            }

            if (expect.IndexOf("楷", StringComparison.Ordinal) >= 0)
            {
                return actual.IndexOf("楷", StringComparison.Ordinal) >= 0
                    || actual.IndexOf("KaiTi", StringComparison.OrdinalIgnoreCase) >= 0;
            }

            if (expect.IndexOf("宋", StringComparison.Ordinal) >= 0)
            {
                return actual.IndexOf("宋", StringComparison.Ordinal) >= 0
                    || actual.IndexOf("SimSun", StringComparison.OrdinalIgnoreCase) >= 0;
            }

            return false;
        }

        private static bool Near(double actual, double expect, double tol)
        {
            return Math.Abs(actual - expect) <= tol;
        }

        private static bool BlobHas(ToolResult r, string needle)
        {
            if (Contains(r == null ? null : r.Error, needle))
            {
                return true;
            }

            if (r == null || r.Data == null)
            {
                return false;
            }

            try
            {
                return Contains(JsonConvert.SerializeObject(r.Data), needle);
            }
            catch
            {
                return Contains(Convert.ToString(r.Data), needle);
            }
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
    }
}
