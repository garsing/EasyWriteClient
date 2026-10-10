using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using Newtonsoft.Json;
using WordAddIn1;

namespace WordDocumentRoundtripTest
{
    /// <summary>
    /// 第 3 波：只测 F_process_document_actions。铺稿用 COM，抄当次 S_/T_。
    /// </summary>
    internal static class WaCases
    {
        public const int Target = 42;

        private const string SentA = "易写改字甲句。";
        private const string SentB = "易写改字乙句。";
        private const string SentC = "易写改字丙句。";
        private const string SentNew = "易写改字新句。";
        private const string SentIns = "易写改字插入句。";
        private const string SentFront = "易写改字文首句。";
        private const string EmptyIns = "易写空稿写入句。";
        private const string Dup = "重复改字夹具。";
        private const string Unique = "独特改字甲句。";

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
            RunEmpty(run, s);
            RunReplace(run, s);
            RunDelete(run, s);
            RunInsert(run, s);
            RunContinuity(run, s);
            RunAmbiguity(run, s);
            RunTable(run, s);
            RunNegLate(run, s);
            RunNoQuit(run, wordBefore, wpsBefore, wordWinBefore, wpsWinBefore);
        }

        private static void RunNegEarly(TestRun run, WdSession s)
        {
            ToolResult noCh = s.ProcessActions(null, InsertOmit(EmptyIns));
            if (!noCh.Success && Contains(noCh.Error, "打开"))
            {
                run.CaseOk("wa-neg-no-channel");
            }
            else if (noCh.Success)
            {
                run.CaseSkip("wa-neg-no-channel", "现场已有默认文字渠，不关用户稿");
            }
            else
            {
                run.CaseFail("wa-neg-no-channel", noCh.Error ?? "应失败或跳过");
            }

            ExpectFail(run, "wa-neg-unknown-id", s.ProcessActions("w999", InsertOmit(EmptyIns)), "未知");
        }

        private static void RunEmpty(TestRun run, WdSession s)
        {
            string path = s.NewDocxPath();
            ToolResult opened = s.Open(path, true, s.App);
            if (!opened.Success)
            {
                FailMany(run, opened.Error ?? "新建失败",
                    "wa-neg-no-action", "wa-neg-replace-no-orig", "wa-neg-replace-no-fmt",
                    "wa-neg-insert-no-fmt", "wa-neg-empty-fmt", "wa-neg-inherit", "wa-neg-pos-end",
                    "wa-empty-omit", "wa-empty-omit-text", "wa-empty-start", "wa-empty-start-text");
                return;
            }

            string id = WdSession.Str(opened, "channel_id");
            var onlyId = new Dictionary<string, object> { ["channel_id"] = id };
            ExpectFail(run, "wa-neg-no-action", s.ProcessActionsRaw(onlyId), "action");

            ExpectFail(
                run,
                "wa-neg-replace-no-orig",
                s.ProcessActions(id, One("replace", new Dictionary<string, object>
                {
                    ["new"] = SentNew,
                    ["format"] = DefaultFormat()
                })),
                "original");

            if (!s.TypePlain(path, SentA + SentB))
            {
                FailMany(run, "COM 写句失败",
                    "wa-neg-replace-no-fmt", "wa-neg-insert-no-fmt", "wa-neg-empty-fmt",
                    "wa-neg-inherit", "wa-neg-pos-end",
                    "wa-empty-omit", "wa-empty-omit-text", "wa-empty-start", "wa-empty-start-text");
                s.CloseDocByPath(path);
                return;
            }

            ToolResult got = s.GetContent(id);
            string code = FirstCode(WdSession.Str(got, "document_content"), "S");
            if (!got.Success || string.IsNullOrEmpty(code))
            {
                FailMany(run, got.Success ? "无可抄 S_" : (got.Error ?? "读失败"),
                    "wa-neg-replace-no-fmt", "wa-neg-insert-no-fmt", "wa-neg-empty-fmt",
                    "wa-neg-inherit", "wa-neg-pos-end",
                    "wa-empty-omit", "wa-empty-omit-text", "wa-empty-start", "wa-empty-start-text");
                s.CloseDocByPath(path);
                return;
            }

            ToolResult noFmt = s.ProcessActions(id, Replace(code, SentNew, format: null));
            Ok(run, "wa-neg-replace-no-fmt",
                !noFmt.Success && BlobHas(noFmt, "char_format") && StillHas(s, id, SentA),
                noFmt.Success ? "缺 format 仍成功" : (noFmt.Error ?? "文案无 char_format 或已改字"));

            ToolResult insNoFmt = s.ProcessActions(id, InsertAfter(code, SentIns, format: null));
            Ok(run, "wa-neg-insert-no-fmt",
                !insNoFmt.Success && BlobHas(insNoFmt, "char_format") && !StillHas(s, id, SentIns),
                insNoFmt.Success ? "缺 format 仍插入" : (insNoFmt.Error ?? "文案无 char_format"));

            ToolResult emptyFmt = s.ProcessActions(id, Replace(code, SentNew, EmptyCharFormat()));
            Ok(run, "wa-neg-empty-fmt",
                !emptyFmt.Success && (BlobHas(emptyFmt, "空对象") || BlobHas(emptyFmt, "char_format")),
                emptyFmt.Success ? "空 char_format 仍成功" : (emptyFmt.Error ?? ""));

            var inheritDetail = new Dictionary<string, object>
            {
                ["original"] = Locator(code),
                ["new"] = SentNew,
                ["format"] = DefaultFormat(),
                ["inherit_from"] = code
            };
            ToolResult inherit = s.ProcessActions(id, One("replace", inheritDetail));
            Ok(run, "wa-neg-inherit",
                !inherit.Success && BlobHas(inherit, "inherit"),
                inherit.Success ? "inherit_from 仍成功" : (inherit.Error ?? "文案无 inherit"));

            ToolResult posEnd = s.ProcessActions(id, InsertAt(new Dictionary<string, object> { ["position"] = "end" }, SentIns));
            Ok(run, "wa-neg-pos-end",
                !posEnd.Success && BlobHas(posEnd, "start"),
                posEnd.Success ? "position=end 仍成功" : (posEnd.Error ?? ""));

            s.CloseDocByPath(path);

            string emptyPath = s.NewDocxPath();
            ToolResult emptyOpen = s.Open(emptyPath, true, s.App);
            if (!emptyOpen.Success)
            {
                FailMany(run, emptyOpen.Error ?? "再建空稿失败",
                    "wa-empty-omit", "wa-empty-omit-text", "wa-empty-start", "wa-empty-start-text");
                return;
            }

            string emptyId = WdSession.Str(emptyOpen, "channel_id");
            s.ActivateDoc(emptyPath);
            ToolResult omit = s.ProcessActions(emptyId, InsertOmit(EmptyIns));
            Ok(run, "wa-empty-omit", omit.Success, omit.Error ?? "空稿省略 target 失败");
            Ok(run, "wa-empty-omit-text",
                omit.Success && StillHas(s, emptyId, EmptyIns),
                omit.Success ? "空稿省略 target 未写出" : (omit.Error ?? ""));
            s.CloseDocByPath(emptyPath);

            string startPath = s.NewDocxPath();
            ToolResult startOpen = s.Open(startPath, true, s.App);
            if (!startOpen.Success)
            {
                FailMany(run, startOpen.Error ?? "再建空稿失败", "wa-empty-start", "wa-empty-start-text");
                return;
            }

            string startId = WdSession.Str(startOpen, "channel_id");
            ToolResult start = s.ProcessActions(startId, InsertStart(EmptyIns));
            Ok(run, "wa-empty-start", start.Success, start.Error ?? "空稿 position=start 失败");
            Ok(run, "wa-empty-start-text",
                start.Success && StillHas(s, startId, EmptyIns),
                start.Success ? "空稿 start 未写出" : (start.Error ?? ""));
            s.CloseDocByPath(startPath);
        }

        private static void RunReplace(TestRun run, WdSession s)
        {
            string path, id, text;
            if (!TryPlant(run, s, SentA + SentB, out path, out id, out text,
                "wa-rep-ok", "wa-rep-new", "wa-rep-gone", "wa-rep-executed", "wa-rep-results",
                "wa-fields", "wa-no-write-key"))
            {
                return;
            }

            string code = FirstCode(text, "S");
            ToolResult wr = s.ProcessActions(id, Replace(code, SentNew));
            if (!wr.Success)
            {
                FailMany(run, wr.Error ?? "replace 失败",
                    "wa-rep-ok", "wa-rep-new", "wa-rep-gone", "wa-rep-executed", "wa-rep-results",
                    "wa-fields", "wa-no-write-key");
                s.CloseDocByPath(path);
                return;
            }

            Ok(run, "wa-rep-ok", true, "");
            bool hasNew = StillHas(s, id, SentNew);
            bool gone = !StillHas(s, id, SentA);
            Ok(run, "wa-rep-new", hasNew, "新句未出现");
            Ok(run, "wa-rep-gone", gone, "旧句还在");
            Ok(run, "wa-rep-executed", WdSession.IntOf(wr, "executed_actions") >= 1, "executed_actions=0");
            Ok(run, "wa-rep-results", WdSession.HasKey(wr, "results"), "无 results");
            Ok(run, "wa-fields",
                WdSession.HasKey(wr, "total_actions")
                && WdSession.HasKey(wr, "executed_actions")
                && WdSession.HasKey(wr, "results"),
                "缺回包字段");
            Ok(run, "wa-no-write-key",
                !WdSession.HasKey(wr, "snapshot")
                && !WdSession.HasKey(wr, "display")
                && !WdSession.HasKey(wr, "display_contents"),
                "Data 带了 snapshot 旁路");
            s.CloseDocByPath(path);
        }

        private static void RunDelete(TestRun run, WdSession s)
        {
            string path, id, text;
            if (!TryPlant(run, s, SentA + SentB, out path, out id, out text,
                "wa-del-ok", "wa-del-gone", "wa-del-keep-other"))
            {
                return;
            }

            string code = FirstCode(text, "S");
            ToolResult wr = s.ProcessActions(id, Delete(code));
            Ok(run, "wa-del-ok", wr.Success, wr.Error ?? "delete 失败");
            if (!wr.Success)
            {
                FailMany(run, wr.Error ?? "delete 失败", "wa-del-gone", "wa-del-keep-other");
                s.CloseDocByPath(path);
                return;
            }

            Ok(run, "wa-del-gone", !StillHas(s, id, SentA), "甲句未删掉");
            Ok(run, "wa-del-keep-other", StillHas(s, id, SentB), "乙句被误删");
            s.CloseDocByPath(path);
        }

        private static void RunInsert(TestRun run, WdSession s)
        {
            string path, id, text;
            if (!TryPlant(run, s, SentA + SentB, out path, out id, out text,
                "wa-ins-after", "wa-ins-keep", "wa-ins-new", "wa-ins-start-front"))
            {
                return;
            }

            string code = FirstCode(text, "S");
            ToolResult wr = s.ProcessActions(id, InsertAfter(code, SentIns));
            Ok(run, "wa-ins-after", wr.Success, wr.Error ?? "insert 失败");
            if (!wr.Success)
            {
                FailMany(run, wr.Error ?? "insert 失败", "wa-ins-keep", "wa-ins-new", "wa-ins-start-front");
                s.CloseDocByPath(path);
                return;
            }

            Ok(run, "wa-ins-keep", StillHas(s, id, SentA), "原句丢了");
            Ok(run, "wa-ins-new", StillHas(s, id, SentIns), "插入句未出现");

            ToolResult front = s.ProcessActions(id, InsertStart(SentFront));
            string after = front.Success ? ReadBody(s, id) : "";
            int frontAt = after.IndexOf(SentFront, StringComparison.Ordinal);
            int oldAt = after.IndexOf(SentA, StringComparison.Ordinal);
            Ok(run, "wa-ins-start-front",
                front.Success && frontAt >= 0 && (oldAt < 0 || frontAt < oldAt),
                front.Success ? "文首句未落到最前" : (front.Error ?? "start 失败"));
            s.CloseDocByPath(path);
        }

        private static void RunContinuity(TestRun run, WdSession s)
        {
            string path, id, text;
            if (!TryPlant(run, s, SentA + SentB + SentC, out path, out id, out text,
                "wa-cont-ok", "wa-cont-text", "wa-cont-gap"))
            {
                return;
            }

            List<string> codes = ListCodes(text, "S");
            if (codes.Count < 3)
            {
                FailMany(run, "S_ 不足 3 个", "wa-cont-ok", "wa-cont-text", "wa-cont-gap");
                s.CloseDocByPath(path);
                return;
            }

            ToolResult ok = s.ProcessActions(id, Replace(codes[0] + "," + codes[1], "新甲句。新乙句。"));
            Ok(run, "wa-cont-ok", ok.Success, ok.Error ?? "连续码 replace 失败");
            Ok(run, "wa-cont-text",
                ok.Success && StillHas(s, id, "新甲句") && StillHas(s, id, "新乙句") && StillHas(s, id, SentC),
                ok.Success ? "连续替换正文不对" : (ok.Error ?? ""));

            string gapCodes = ok.Success
                ? (FirstCode(ReadBody(s, id), "S") + "," + LastCode(ReadBody(s, id), "S"))
                : (codes[0] + "," + codes[2]);
            ToolResult gap = s.ProcessActions(id, Replace(gapCodes, "空隙句。"));

            Ok(run, "wa-cont-gap",
                !gap.Success && (BlobHas(gap, "不连续") || BlobHas(gap, "continuity")),
                gap.Success ? "不连续码仍成功" : (gap.Error ?? "文案无「不连续」"));
            s.CloseDocByPath(path);
        }

        private static void RunAmbiguity(TestRun run, WdSession s)
        {
            string path, id, text;
            if (!TryPlant(run, s, Unique + "\r" + Dup + "\r" + Dup, out path, out id, out text,
                "wa-amb-fail", "wa-amb-no-write", "wa-amb-suggested", "wa-amb-batch"))
            {
                return;
            }

            string uniqueCode = FirstCode(text, "S");
            string dupCode = FindDupCode(text);
            if (string.IsNullOrEmpty(dupCode) || string.Equals(uniqueCode, dupCode, StringComparison.Ordinal))
            {
                FailMany(run, "未形成重复 S_（note/树应对同一码出现两次）",
                    "wa-amb-fail", "wa-amb-no-write", "wa-amb-suggested", "wa-amb-batch");
                s.CloseDocByPath(path);
                return;
            }

            ToolResult amb = s.ProcessActions(id, Replace(dupCode, SentNew));
            Ok(run, "wa-amb-fail", !amb.Success, amb.Success ? "重复码仍成功" : "");
            Ok(run, "wa-amb-no-write",
                !amb.Success && StillHas(s, id, Dup) && !StillHas(s, id, SentNew),
                amb.Success ? "歧义仍写了字" : "夹具已被改");
            Ok(run, "wa-amb-suggested",
                !amb.Success && BlobHas(amb, "suggested_locator_codes"),
                amb.Success ? "未失败" : "回包无 suggested_locator_codes");

            var batch = new List<Dictionary<string, object>>();
            batch.AddRange(Replace(uniqueCode, SentNew));
            batch.AddRange(Replace(dupCode, "另一处新句。"));
            ToolResult both = s.ProcessActions(id, batch);
            Ok(run, "wa-amb-batch",
                !both.Success
                && StillHas(s, id, Unique)
                && WdSession.IntOf(both, "executed_actions") == 0,
                both.Success ? "歧义整批仍执行" : "独特句被改或 executed≠0");
            s.CloseDocByPath(path);
        }

        private static void RunTable(TestRun run, WdSession s)
        {
            string path = s.NewDocxPath();
            ToolResult opened = s.Open(path, true, s.App);
            if (!opened.Success)
            {
                FailMany(run, opened.Error ?? "新建失败",
                    "wa-tbl-marker", "wa-tbl-rep", "wa-tbl-cell", "wa-tbl-wrong-t");
                return;
            }

            string id = WdSession.Str(opened, "channel_id");
            if (!s.InsertTable(path, 2, 2, new[] { "甲", "乙", "丙", "丁" }))
            {
                FailMany(run, "COM 建表失败",
                    "wa-tbl-marker", "wa-tbl-rep", "wa-tbl-cell", "wa-tbl-wrong-t");
                s.CloseDocByPath(path);
                return;
            }

            ToolResult got = s.GetContent(id);
            string body = got.Success ? WdSession.Str(got, "document_content") : "";
            Match marker = Regex.Match(body, @"\[S_([0-9a-zA-Z]{5})-start\]\[in_table=(T_[^\]]+)\]");
            Ok(run, "wa-tbl-marker",
                got.Success && marker.Success,
                got.Success ? "表内末句无 [in_table=T_]" : (got.Error ?? "读表失败"));
            if (!got.Success || !marker.Success)
            {
                FailMany(run, got.Success ? "无 in_table 标记" : (got.Error ?? "读表失败"),
                    "wa-tbl-rep", "wa-tbl-cell", "wa-tbl-wrong-t");
                s.CloseDocByPath(path);
                return;
            }

            string code = "S_" + marker.Groups[1].Value;
            string tableId = marker.Groups[2].Value;
            ToolResult wr = s.ProcessActions(id, ReplaceInTable(code, tableId, "戊。"));
            Ok(run, "wa-tbl-rep", wr.Success, wr.Error ?? "表内 replace 失败");
            Ok(run, "wa-tbl-cell",
                wr.Success && StillHas(s, id, "戊") && !StillHas(s, id, "丁"),
                wr.Success ? "表内格字未换成戊" : (wr.Error ?? ""));

            ToolResult wrong = s.ProcessActions(id, ReplaceInTable(code, "T_zzzzz", "己。"));
            Ok(run, "wa-tbl-wrong-t",
                !wrong.Success && !StillHas(s, id, "己"),
                wrong.Success ? "错 T_ 仍成功" : (wrong.Error ?? "错 T_ 却改了字"));
            s.CloseDocByPath(path);
        }

        private static void RunNegLate(TestRun run, WdSession s)
        {
            string xlsx = Path.Combine(Path.GetTempPath(), "ew-wa-" + Guid.NewGuid().ToString("N") + ".xlsx");
            ToolResult xl = s.Open(xlsx, true, "excel");
            if (!xl.Success)
            {
                run.CaseFail("wa-neg-wrong-kind-excel", xl.Error ?? "建 xlsx 失败");
            }
            else
            {
                ExpectFail(run, "wa-neg-wrong-kind-excel",
                    s.ProcessActions(WdSession.Str(xl, "channel_id"), InsertOmit(EmptyIns)), "excel");
                WdSession.CloseExcelByPath(WdSession.Str(xl, "path"));
            }

            string pptx = Path.Combine(Path.GetTempPath(), "ew-wa-" + Guid.NewGuid().ToString("N") + ".pptx");
            ToolResult ppt = s.Open(pptx, true, "powerpoint");
            if (!ppt.Success)
            {
                run.CaseFail("wa-neg-wrong-kind-ppt", ppt.Error ?? "建 pptx 失败");
            }
            else
            {
                ExpectFail(run, "wa-neg-wrong-kind-ppt",
                    s.ProcessActions(WdSession.Str(ppt, "channel_id"), InsertOmit(EmptyIns)), "ppt");
                WdSession.ClosePptByPath(WdSession.Str(ppt, "path"));
            }

            string closePath = s.NewDocxPath();
            ToolResult opened = s.Open(closePath, true, s.App);
            if (!opened.Success)
            {
                run.CaseFail("wa-neg-closed", opened.Error ?? "再建稿失败");
                return;
            }

            string cid = WdSession.Str(opened, "channel_id");
            if (!s.CloseDocByPath(closePath))
            {
                run.CaseFail("wa-neg-closed", "COM 关稿失败");
                return;
            }

            ToolResult closed = s.ProcessActions(cid, InsertOmit(EmptyIns));
            if (closed != null && !closed.Success
                && (Contains(closed.Error, "关闭")
                    || Contains(closed.Error, "找不到")
                    || Contains(closed.Error, "未知")
                    || Contains(closed.Error, "unsupported")))
            {
                run.CaseOk("wa-neg-closed");
            }
            else if (closed != null && !closed.Success)
            {
                run.CaseFail("wa-neg-closed", closed.Error ?? "");
            }
            else
            {
                run.CaseFail("wa-neg-closed", "关稿后仍成功");
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
            Ok(run, "wa-no-quit", wordOk && wpsOk, wordDetail + " " + wpsDetail);
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

        private static List<Dictionary<string, object>> One(string type, Dictionary<string, object> detail)
        {
            return new List<Dictionary<string, object>>
            {
                new Dictionary<string, object>
                {
                    ["type"] = type,
                    ["detail"] = detail
                }
            };
        }

        private static List<Dictionary<string, object>> Replace(string codes, string neu, Dictionary<string, object> format)
        {
            var detail = new Dictionary<string, object>
            {
                ["original"] = Locator(codes),
                ["new"] = neu
            };
            if (format != null)
            {
                detail["format"] = format;
            }

            return One("replace", detail);
        }

        private static List<Dictionary<string, object>> Replace(string codes, string neu)
        {
            return Replace(codes, neu, DefaultFormat());
        }

        private static List<Dictionary<string, object>> ReplaceInTable(string codes, string tableId, string neu)
        {
            return One("replace", new Dictionary<string, object>
            {
                ["original"] = new Dictionary<string, object>
                {
                    ["codes"] = codes,
                    ["in_table"] = tableId
                },
                ["new"] = neu,
                ["format"] = DefaultFormat()
            });
        }

        private static List<Dictionary<string, object>> Delete(string codes)
        {
            return One("delete", new Dictionary<string, object>
            {
                ["deletion"] = Locator(codes)
            });
        }

        private static List<Dictionary<string, object>> InsertAfter(string codes, string insert, Dictionary<string, object> format)
        {
            var detail = new Dictionary<string, object>
            {
                ["target"] = Locator(codes),
                ["insert"] = insert
            };
            if (format != null)
            {
                detail["format"] = format;
            }

            return One("insert", detail);
        }

        private static List<Dictionary<string, object>> InsertAfter(string codes, string insert)
        {
            return InsertAfter(codes, insert, DefaultFormat());
        }

        private static List<Dictionary<string, object>> InsertStart(string insert)
        {
            return InsertAt(new Dictionary<string, object> { ["position"] = "start" }, insert);
        }

        private static List<Dictionary<string, object>> InsertOmit(string insert)
        {
            return One("insert", new Dictionary<string, object>
            {
                ["insert"] = insert,
                ["format"] = DefaultFormat()
            });
        }

        private static List<Dictionary<string, object>> InsertAt(Dictionary<string, object> target, string insert)
        {
            return One("insert", new Dictionary<string, object>
            {
                ["target"] = target,
                ["insert"] = insert,
                ["format"] = DefaultFormat()
            });
        }

        private static Dictionary<string, object> Locator(string codes)
        {
            return new Dictionary<string, object> { ["codes"] = codes };
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
