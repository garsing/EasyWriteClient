using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using Newtonsoft.Json;
using WordAddIn1;

namespace WordDocumentRoundtripTest
{
    /// <summary>
    /// 第 4 波：只测 from_file 插 / 换。铺稿用 COM，抄当次 S_/T_。
    /// </summary>
    internal static class WfCases
    {
        public const int Target = 44;

        private const string SentA = "易写从文件甲句。";
        private const string SentB = "易写从文件乙句。";
        private const string SentC = "易写从文件丙句。";
        private const string SentNew = "易写从文件新句。";
        private const string SentIns = "易写从文件插入句。";
        private const string SentFront = "易写从文件文首句。";
        private const string EmptyIns = "易写从文件空稿句。";
        private const string MultiA = "易写从文件多句甲。";
        private const string MultiB = "易写从文件多句乙。";
        private const string ContA = "易写从文件新甲。";
        private const string ContB = "易写从文件新乙。";
        private const string AfterTable = "易写从文件表后句。";
        private const string CellIns = "易写从文件戊。";
        private const string Dup = "重复从文件夹具。";
        private const string Unique = "独特从文件甲句。";

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
                RunEmptyAndContract(run, s);
                RunInsert(run, s);
                RunReplace(run, s);
                RunTable(run, s);
                RunAmbiguity(run, s);
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
            s.GetContent(id);
            s.InsertFromFile(InsertArgs(id, WriteTxt(EmptyIns)));
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
            if (opened.Success)
            {
                string id = WdSession.Str(opened, "channel_id");
                s.ActivateDoc(path);
                s.InsertFromFile(InsertArgs(id, WriteTxt(EmptyIns)));
                s.CloseDocByPath(path);
            }
        }

        private static void RunNegEarly(TestRun run, WdSession s)
        {
            string txt = WriteTxt(EmptyIns);
            ToolResult noCh = s.InsertFromFile(InsertArgs(null, txt));
            if (!noCh.Success
                && (Contains(noCh.Error, "打开")
                    || Contains(noCh.Error, "unsupported")
                    || Contains(noCh.Error, "browser")
                    || Contains(noCh.Error, "渠道")))
            {
                run.CaseOk("wf-neg-no-channel");
            }
            else if (noCh.Success)
            {
                run.CaseSkip("wf-neg-no-channel", "现场已有默认文字渠，不关用户稿");
            }
            else
            {
                run.CaseFail("wf-neg-no-channel", noCh.Error ?? "应失败或跳过");
            }

            ExpectFail(run, "wf-neg-unknown-id", s.InsertFromFile(InsertArgs("w999", txt)), "未知");
        }

        private static void RunEmptyAndContract(TestRun run, WdSession s)
        {
            BurnFirstWpsOpen(s);

            string emptyPath = s.NewDocxPath();
            ToolResult emptyOpen = s.Open(emptyPath, true, s.App);
            if (!emptyOpen.Success)
            {
                FailMany(run, emptyOpen.Error ?? "再建空稿失败",
                    "wf-empty-omit", "wf-empty-omit-text",
                    "wf-neg-ins-no-path", "wf-neg-rep-no-path",
                    "wf-neg-ins-no-fmt", "wf-neg-rep-no-fmt", "wf-neg-empty-fmt", "wf-neg-inherit",
                    "wf-neg-rep-no-codes", "wf-neg-file-missing", "wf-neg-file-empty",
                    "wf-neg-pos-end", "wf-neg-locator-mutex", "wf-neg-intable-alone",
                    "wf-empty-start", "wf-empty-start-text");
                return;
            }

            string emptyId = WdSession.Str(emptyOpen, "channel_id");
            s.ActivateDoc(emptyPath);
            string emptyTxt = WriteTxt(EmptyIns);
            ToolResult omit = s.InsertFromFile(InsertArgs(emptyId, emptyTxt));
            Ok(run, "wf-empty-omit", omit.Success, omit.Error ?? "空稿省略定位失败");
            Ok(run, "wf-empty-omit-text",
                omit.Success && StillHas(s, emptyId, EmptyIns),
                omit.Success ? "空稿省略定位未写出" : (omit.Error ?? ""));

            ExpectFail(
                run,
                "wf-neg-ins-no-path",
                s.InsertFromFile(new Dictionary<string, object>
                {
                    ["channel_id"] = emptyId,
                    ["format"] = DefaultFormat()
                }),
                "path");

            ExpectFail(
                run,
                "wf-neg-rep-no-path",
                s.ReplaceFromFile(new Dictionary<string, object>
                {
                    ["channel_id"] = emptyId,
                    ["original_codes"] = "S_aaaaa",
                    ["format"] = DefaultFormat()
                }),
                "path");

            if (!s.TypePlain(emptyPath, SentA + SentB))
            {
                FailMany(run, "COM 写句失败",
                    "wf-neg-ins-no-fmt", "wf-neg-rep-no-fmt", "wf-neg-empty-fmt", "wf-neg-inherit",
                    "wf-neg-rep-no-codes", "wf-neg-file-missing", "wf-neg-file-empty",
                    "wf-neg-pos-end", "wf-neg-locator-mutex", "wf-neg-intable-alone",
                    "wf-empty-start", "wf-empty-start-text");
                s.CloseDocByPath(emptyPath);
                return;
            }

            ToolResult got = s.GetContent(emptyId);
            string code = FirstCode(WdSession.Str(got, "document_content"), "S");
            if (!got.Success || string.IsNullOrEmpty(code))
            {
                FailMany(run, got.Success ? "无可抄 S_" : (got.Error ?? "读失败"),
                    "wf-neg-ins-no-fmt", "wf-neg-rep-no-fmt", "wf-neg-empty-fmt", "wf-neg-inherit",
                    "wf-neg-rep-no-codes", "wf-neg-file-missing", "wf-neg-file-empty",
                    "wf-neg-pos-end", "wf-neg-locator-mutex", "wf-neg-intable-alone",
                    "wf-empty-start", "wf-empty-start-text");
                s.CloseDocByPath(emptyPath);
                return;
            }

            string insTxt = WriteTxt(SentIns);
            string newTxt = WriteTxt(SentNew);

            var insNoFmt = InsertArgs(emptyId, insTxt, code);
            insNoFmt.Remove("format");
            ToolResult noFmtIns = s.InsertFromFile(insNoFmt);
            Ok(run, "wf-neg-ins-no-fmt",
                !noFmtIns.Success && BlobHas(noFmtIns, "char_format") && !StillHas(s, emptyId, SentIns),
                noFmtIns.Success ? "缺 format 仍插入" : (noFmtIns.Error ?? "文案无 char_format"));

            var repNoFmt = ReplaceArgs(emptyId, newTxt, code);
            repNoFmt.Remove("format");
            ToolResult noFmtRep = s.ReplaceFromFile(repNoFmt);
            Ok(run, "wf-neg-rep-no-fmt",
                !noFmtRep.Success && BlobHas(noFmtRep, "char_format") && StillHas(s, emptyId, SentA),
                noFmtRep.Success ? "缺 format 仍成功" : (noFmtRep.Error ?? "文案无 char_format 或已改字"));

            var emptyFmt = ReplaceArgs(emptyId, newTxt, code);
            emptyFmt["format"] = EmptyCharFormat();
            ToolResult emptyFmtR = s.ReplaceFromFile(emptyFmt);
            Ok(run, "wf-neg-empty-fmt",
                !emptyFmtR.Success && (BlobHas(emptyFmtR, "空对象") || BlobHas(emptyFmtR, "char_format")),
                emptyFmtR.Success ? "空 char_format 仍成功" : (emptyFmtR.Error ?? ""));

            var inherit = InsertArgs(emptyId, insTxt, code);
            inherit["inherit_from"] = code;
            ToolResult inheritR = s.InsertFromFile(inherit);
            Ok(run, "wf-neg-inherit",
                !inheritR.Success && BlobHas(inheritR, "inherit") && StillHas(s, emptyId, SentA),
                inheritR.Success ? "inherit_from 仍成功" : (inheritR.Error ?? "文案无 inherit"));

            ExpectFail(
                run,
                "wf-neg-rep-no-codes",
                s.ReplaceFromFile(new Dictionary<string, object>
                {
                    ["channel_id"] = emptyId,
                    ["path"] = newTxt,
                    ["format"] = DefaultFormat()
                }),
                "original_codes");

            ToolResult missing = s.InsertFromFile(InsertArgs(
                emptyId,
                Path.Combine(Path.GetTempPath(), "ew-wf-missing-" + Guid.NewGuid().ToString("N") + ".txt"),
                code));
            Ok(run, "wf-neg-file-missing",
                !missing.Success && BlobHas(missing, "不存在") && StillHas(s, emptyId, SentA),
                missing.Success ? "缺文件仍成功" : (missing.Error ?? "文案无不存在"));

            string emptyFile = Path.Combine(Path.GetTempPath(), "ew-wf-empty-" + Guid.NewGuid().ToString("N") + ".txt");
            File.WriteAllText(emptyFile, "");
            ToolResult emptyFileR = s.InsertFromFile(InsertArgs(emptyId, emptyFile, code));
            Ok(run, "wf-neg-file-empty",
                !emptyFileR.Success && BlobHas(emptyFileR, "空") && StillHas(s, emptyId, SentA),
                emptyFileR.Success ? "空文件仍成功" : (emptyFileR.Error ?? "文案无空"));

            var posEnd = InsertArgs(emptyId, insTxt);
            posEnd["target_position"] = "end";
            ToolResult posEndR = s.InsertFromFile(posEnd);
            Ok(run, "wf-neg-pos-end",
                !posEndR.Success && BlobHas(posEndR, "start"),
                posEndR.Success ? "position=end 仍成功" : (posEndR.Error ?? ""));

            var mutex = InsertArgs(emptyId, insTxt, code);
            mutex["target_position"] = "start";
            ToolResult mutexR = s.InsertFromFile(mutex);
            Ok(run, "wf-neg-locator-mutex",
                !mutexR.Success && BlobHas(mutexR, "五选一"),
                mutexR.Success ? "双定位仍成功" : (mutexR.Error ?? "文案无五选一"));

            var inTableAlone = InsertArgs(emptyId, insTxt);
            inTableAlone["in_table"] = "T_zzzzz";
            ToolResult inTableR = s.InsertFromFile(inTableAlone);
            Ok(run, "wf-neg-intable-alone",
                !inTableR.Success && BlobHas(inTableR, "in_table"),
                inTableR.Success ? "单独 in_table 仍成功" : (inTableR.Error ?? "文案无 in_table"));

            s.CloseDocByPath(emptyPath);

            string startPath = s.NewDocxPath();
            ToolResult startOpen = s.Open(startPath, true, s.App);
            if (!startOpen.Success)
            {
                FailMany(run, startOpen.Error ?? "再建空稿失败", "wf-empty-start", "wf-empty-start-text");
                return;
            }

            string startId = WdSession.Str(startOpen, "channel_id");
            var startArgs = InsertArgs(startId, WriteTxt(EmptyIns));
            startArgs["target_position"] = "start";
            ToolResult start = s.InsertFromFile(startArgs);
            Ok(run, "wf-empty-start", start.Success, start.Error ?? "空稿 position=start 失败");
            Ok(run, "wf-empty-start-text",
                start.Success && StillHas(s, startId, EmptyIns),
                start.Success ? "空稿 start 未写出" : (start.Error ?? ""));
            s.CloseDocByPath(startPath);
        }

        private static void RunInsert(TestRun run, WdSession s)
        {
            string path, id, text;
            if (!TryPlant(run, s, SentA + SentB, out path, out id, out text,
                "wf-ins-after", "wf-ins-keep", "wf-ins-new", "wf-ins-start-front",
                "wf-ins-fields", "wf-no-write-key"))
            {
                return;
            }

            string code = FirstCode(text, "S");
            ToolResult wr = s.InsertFromFile(InsertArgs(id, WriteTxt(SentIns), code));
            Ok(run, "wf-ins-after", wr.Success, wr.Error ?? "insert 失败");
            if (!wr.Success)
            {
                FailMany(run, wr.Error ?? "insert 失败",
                    "wf-ins-keep", "wf-ins-new", "wf-ins-start-front",
                    "wf-ins-fields", "wf-no-write-key");
                s.CloseDocByPath(path);
                return;
            }

            Ok(run, "wf-ins-keep", StillHas(s, id, SentA), "原句丢了");
            Ok(run, "wf-ins-new", StillHas(s, id, SentIns), "插入句未出现");
            Ok(run, "wf-ins-fields",
                (WdSession.HasKey(wr, "path") || WdSession.HasKey(wr, "insert_file"))
                && WdSession.HasKey(wr, "characters_inserted")
                && WdSession.HasKey(wr, "process_result")
                && WdSession.HasKey(wr, "message"),
                "缺 insert 回包字段");
            Ok(run, "wf-no-write-key",
                !WdSession.HasKey(wr, "snapshot")
                && !WdSession.HasKey(wr, "display")
                && !WdSession.HasKey(wr, "display_contents"),
                "Data 带了 snapshot 旁路");

            var frontArgs = InsertArgs(id, WriteTxt(SentFront));
            frontArgs["target_position"] = "start";
            ToolResult front = s.InsertFromFile(frontArgs);
            string after = front.Success ? ReadBody(s, id) : "";
            int frontAt = after.IndexOf(SentFront, StringComparison.Ordinal);
            int oldAt = after.IndexOf(SentA, StringComparison.Ordinal);
            Ok(run, "wf-ins-start-front",
                front.Success && frontAt >= 0 && (oldAt < 0 || frontAt < oldAt),
                front.Success ? "文首句未落到最前" : (front.Error ?? "start 失败"));
            s.CloseDocByPath(path);

            string aliasPath, aliasId, aliasText;
            if (!TryPlant(run, s, SentA + SentB, out aliasPath, out aliasId, out aliasText, "wf-ins-alias"))
            {
                return;
            }

            string aliasCode = FirstCode(aliasText, "S");
            var aliasArgs = InsertArgs(aliasId, WriteTxt(SentIns), aliasCode);
            aliasArgs.Remove("path");
            aliasArgs["insert_file"] = WriteTxt(SentIns);
            ToolResult alias = s.InsertFromFile(aliasArgs);
            Ok(run, "wf-ins-alias",
                alias.Success && StillHas(s, aliasId, SentIns),
                alias.Success ? "insert_file 别名未写出" : (alias.Error ?? "insert_file 失败"));
            s.CloseDocByPath(aliasPath);

            string multiPath, multiId, multiText;
            if (!TryPlant(run, s, SentA + SentB, out multiPath, out multiId, out multiText, "wf-ins-multi"))
            {
                return;
            }

            string multiCode = FirstCode(multiText, "S");
            ToolResult multi = s.InsertFromFile(InsertArgs(multiId, WriteTxt(MultiA + MultiB), multiCode));
            Ok(run, "wf-ins-multi",
                multi.Success
                && StillHas(s, multiId, MultiA)
                && StillHas(s, multiId, MultiB)
                && StillHas(s, multiId, SentA),
                multi.Success ? "多句 txt 未完整写入" : (multi.Error ?? "多句 insert 失败"));
            s.CloseDocByPath(multiPath);
        }

        private static void RunReplace(TestRun run, WdSession s)
        {
            string path, id, text;
            if (!TryPlant(run, s, SentA + SentB, out path, out id, out text,
                "wf-rep-ok", "wf-rep-new", "wf-rep-gone", "wf-rep-fields"))
            {
                return;
            }

            string code = FirstCode(text, "S");
            ToolResult wr = s.ReplaceFromFile(ReplaceArgs(id, WriteTxt(SentNew), code));
            if (!wr.Success)
            {
                FailMany(run, wr.Error ?? "replace 失败",
                    "wf-rep-ok", "wf-rep-new", "wf-rep-gone", "wf-rep-fields");
                s.CloseDocByPath(path);
                return;
            }

            Ok(run, "wf-rep-ok", true, "");
            Ok(run, "wf-rep-new", StillHas(s, id, SentNew), "新句未出现");
            Ok(run, "wf-rep-gone", !StillHas(s, id, SentA), "旧句还在");
            Ok(run, "wf-rep-fields",
                (WdSession.HasKey(wr, "path") || WdSession.HasKey(wr, "replace_file"))
                && WdSession.HasKey(wr, "original_codes")
                && WdSession.HasKey(wr, "characters_replaced")
                && WdSession.HasKey(wr, "process_result"),
                "缺 replace 回包字段");
            s.CloseDocByPath(path);

            string aliasPath, aliasId, aliasText;
            if (!TryPlant(run, s, SentA + SentB, out aliasPath, out aliasId, out aliasText, "wf-rep-alias"))
            {
                return;
            }

            string aliasCode = FirstCode(aliasText, "S");
            var aliasArgs = ReplaceArgs(aliasId, WriteTxt(SentNew), aliasCode);
            aliasArgs.Remove("path");
            aliasArgs["replace_file"] = WriteTxt(SentNew);
            ToolResult alias = s.ReplaceFromFile(aliasArgs);
            Ok(run, "wf-rep-alias",
                alias.Success && StillHas(s, aliasId, SentNew) && !StillHas(s, aliasId, SentA),
                alias.Success ? "replace_file 别名未换掉" : (alias.Error ?? "replace_file 失败"));
            s.CloseDocByPath(aliasPath);

            string contPath, contId, contText;
            if (!TryPlant(run, s, SentA + SentB + SentC, out contPath, out contId, out contText,
                "wf-rep-cont-ok", "wf-rep-cont-text", "wf-rep-cont-gap"))
            {
                return;
            }

            List<string> codes = ListCodes(contText, "S");
            if (codes.Count < 3)
            {
                FailMany(run, "S_ 不足 3 个", "wf-rep-cont-ok", "wf-rep-cont-text", "wf-rep-cont-gap");
                s.CloseDocByPath(contPath);
                return;
            }

            ToolResult ok = s.ReplaceFromFile(ReplaceArgs(contId, WriteTxt(ContA + ContB), codes[0] + "," + codes[1]));
            Ok(run, "wf-rep-cont-ok", ok.Success, ok.Error ?? "连续码 replace 失败");
            Ok(run, "wf-rep-cont-text",
                ok.Success && StillHas(s, contId, ContA) && StillHas(s, contId, ContB) && StillHas(s, contId, SentC),
                ok.Success ? "连续替换正文不对" : (ok.Error ?? ""));

            string gapCodes = ok.Success
                ? (FirstCode(ReadBody(s, contId), "S") + "," + LastCode(ReadBody(s, contId), "S"))
                : (codes[0] + "," + codes[2]);
            ToolResult gap = s.ReplaceFromFile(ReplaceArgs(contId, WriteTxt("易写从文件空隙句。"), gapCodes));
            Ok(run, "wf-rep-cont-gap",
                !gap.Success
                && (BlobHas(gap, "不连续") || BlobHas(gap, "continuity"))
                && StillHas(s, contId, SentC),
                gap.Success ? "不连续码仍成功" : (gap.Error ?? "文案无「不连续」"));
            s.CloseDocByPath(contPath);
        }

        private static void RunTable(TestRun run, WdSession s)
        {
            string path = s.NewDocxPath();
            ToolResult opened = s.Open(path, true, s.App);
            if (!opened.Success)
            {
                FailMany(run, opened.Error ?? "新建失败",
                    "wf-tbl-ins", "wf-tbl-cell", "wf-tbl-wrong-t", "wf-tbl-after");
                return;
            }

            string id = WdSession.Str(opened, "channel_id");
            if (!s.InsertTable(path, 2, 2, new[] { "甲", "乙", "丙", "丁" }))
            {
                FailMany(run, "COM 建表失败",
                    "wf-tbl-ins", "wf-tbl-cell", "wf-tbl-wrong-t", "wf-tbl-after");
                s.CloseDocByPath(path);
                return;
            }

            ToolResult got = s.GetContent(id);
            string body = got.Success ? WdSession.Str(got, "document_content") : "";
            Match marker = Regex.Match(body, @"\[S_([0-9a-zA-Z]{5})-start\]\[in_table=(T_[^\]]+)\]");
            if (!got.Success || !marker.Success)
            {
                FailMany(run, got.Success ? "表内末句无 [in_table=T_]" : (got.Error ?? "读表失败"),
                    "wf-tbl-ins", "wf-tbl-cell", "wf-tbl-wrong-t", "wf-tbl-after");
                s.CloseDocByPath(path);
                return;
            }

            string code = "S_" + marker.Groups[1].Value;
            string tableId = marker.Groups[2].Value;
            var ins = InsertArgs(id, WriteTxt(CellIns), code);
            ins["in_table"] = tableId;
            ToolResult wr = s.InsertFromFile(ins);
            Ok(run, "wf-tbl-ins", wr.Success, wr.Error ?? "表内 insert 失败");
            Ok(run, "wf-tbl-cell",
                wr.Success && StillHas(s, id, "戊") && StillHas(s, id, "丁"),
                wr.Success ? "表内插入后格字不对" : (wr.Error ?? ""));

            var wrong = InsertArgs(id, WriteTxt("易写从文件己。"), code);
            wrong["in_table"] = "T_zzzzz";
            ToolResult wrongR = s.InsertFromFile(wrong);
            Ok(run, "wf-tbl-wrong-t",
                !wrongR.Success && !StillHas(s, id, "己"),
                wrongR.Success ? "错 T_ 仍成功" : (wrongR.Error ?? "错 T_ 却改了字"));

            ToolResult gotAfter = s.GetContent(id);
            string bodyAfter = gotAfter.Success ? WdSession.Str(gotAfter, "document_content") : "";
            Match tableMark = Regex.Match(bodyAfter, @"id=""(T_[^""]+)""");
            if (!tableMark.Success)
            {
                tableMark = Regex.Match(bodyAfter, @"\[in_table=(T_[^\]]+)\]");
            }

            string liveTableId = tableMark.Success ? tableMark.Groups[1].Value : tableId;
            var after = InsertArgs(id, WriteTxt(AfterTable));
            after["target_table"] = liveTableId;
            ToolResult afterR = s.InsertFromFile(after);
            Ok(run, "wf-tbl-after",
                afterR.Success && StillHas(s, id, AfterTable) && StillHas(s, id, "丁"),
                afterR.Success
                    ? "表后句未出现或表内丁丢了"
                    : ("target_table 失败: " + (afterR.Error ?? "(null)")));
            s.CloseDocByPath(path);
        }

        private static void RunAmbiguity(TestRun run, WdSession s)
        {
            string path, id, text;
            if (!TryPlant(run, s, Unique + "\r" + Dup + "\r" + Dup, out path, out id, out text,
                "wf-amb-fail", "wf-amb-no-write"))
            {
                return;
            }

            string uniqueCode = FirstCode(text, "S");
            string dupCode = FindDupCode(text);
            if (string.IsNullOrEmpty(dupCode) || string.Equals(uniqueCode, dupCode, StringComparison.Ordinal))
            {
                FailMany(run, "未形成重复 S_", "wf-amb-fail", "wf-amb-no-write");
                s.CloseDocByPath(path);
                return;
            }

            ToolResult amb = s.ReplaceFromFile(ReplaceArgs(id, WriteTxt(SentNew), dupCode));
            Ok(run, "wf-amb-fail", !amb.Success, amb.Success ? "重复码仍成功" : "");
            Ok(run, "wf-amb-no-write",
                !amb.Success && StillHas(s, id, Dup) && !StillHas(s, id, SentNew),
                amb.Success ? "歧义仍写了字" : "夹具已被改");
            s.CloseDocByPath(path);
        }

        private static void RunNegLate(TestRun run, WdSession s)
        {
            string txt = WriteTxt(EmptyIns);
            string xlsx = Path.Combine(Path.GetTempPath(), "ew-wf-" + Guid.NewGuid().ToString("N") + ".xlsx");
            ToolResult xl = s.Open(xlsx, true, "excel");
            if (!xl.Success)
            {
                run.CaseFail("wf-neg-wrong-kind-excel", xl.Error ?? "建 xlsx 失败");
            }
            else
            {
                ExpectFail(run, "wf-neg-wrong-kind-excel",
                    s.InsertFromFile(InsertArgs(WdSession.Str(xl, "channel_id"), txt)), "excel");
                WdSession.CloseExcelByPath(WdSession.Str(xl, "path"));
            }

            string pptx = Path.Combine(Path.GetTempPath(), "ew-wf-" + Guid.NewGuid().ToString("N") + ".pptx");
            ToolResult ppt = s.Open(pptx, true, "powerpoint");
            if (!ppt.Success)
            {
                run.CaseFail("wf-neg-wrong-kind-ppt", ppt.Error ?? "建 pptx 失败");
            }
            else
            {
                ExpectFail(run, "wf-neg-wrong-kind-ppt",
                    s.InsertFromFile(InsertArgs(WdSession.Str(ppt, "channel_id"), txt)), "ppt");
                WdSession.ClosePptByPath(WdSession.Str(ppt, "path"));
            }

            string closePath = s.NewDocxPath();
            ToolResult opened = s.Open(closePath, true, s.App);
            if (!opened.Success)
            {
                run.CaseFail("wf-neg-closed", opened.Error ?? "再建稿失败");
                return;
            }

            string cid = WdSession.Str(opened, "channel_id");
            if (!s.CloseDocByPath(closePath))
            {
                run.CaseFail("wf-neg-closed", "COM 关稿失败");
                return;
            }

            ToolResult closed = s.InsertFromFile(InsertArgs(cid, txt));
            if (closed != null && !closed.Success
                && (Contains(closed.Error, "关闭")
                    || Contains(closed.Error, "找不到")
                    || Contains(closed.Error, "未知")
                    || Contains(closed.Error, "unsupported")))
            {
                run.CaseOk("wf-neg-closed");
            }
            else if (closed != null && !closed.Success)
            {
                run.CaseFail("wf-neg-closed", closed.Error ?? "");
            }
            else
            {
                run.CaseFail("wf-neg-closed", "关稿后仍成功");
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
            Ok(run, "wf-no-quit", wordOk && wpsOk, wordDetail + " " + wpsDetail);
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

        private static Dictionary<string, object> InsertArgs(string channelId, string path, string targetCodes = null)
        {
            var args = new Dictionary<string, object>
            {
                ["path"] = path,
                ["format"] = DefaultFormat()
            };
            if (channelId != null)
            {
                args["channel_id"] = channelId;
            }

            if (!string.IsNullOrEmpty(targetCodes))
            {
                args["target_codes"] = targetCodes;
            }

            return args;
        }

        private static Dictionary<string, object> ReplaceArgs(string channelId, string path, string originalCodes)
        {
            var args = new Dictionary<string, object>
            {
                ["path"] = path,
                ["original_codes"] = originalCodes ?? "",
                ["format"] = DefaultFormat()
            };
            if (channelId != null)
            {
                args["channel_id"] = channelId;
            }

            return args;
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

        private static Dictionary<string, object> EmptyCharFormat()
        {
            return new Dictionary<string, object>
            {
                ["char_format"] = new Dictionary<string, object>()
            };
        }

        private static string WriteTxt(string text)
        {
            string path = Path.Combine(Path.GetTempPath(), "ew-wf-" + Guid.NewGuid().ToString("N") + ".txt");
            File.WriteAllText(path, text ?? "");
            return path;
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

        private static string LastCode(string text, string prefix)
        {
            List<string> list = ListCodes(text, prefix);
            return list.Count == 0 ? "" : list[list.Count - 1];
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
