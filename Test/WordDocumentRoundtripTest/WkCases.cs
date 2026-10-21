using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using WordAddIn1;

namespace WordDocumentRoundtripTest
{
    /// <summary>
    /// 第 9 波：检查点（list / restore）+ 数据溯源（apply / get）。
    /// 用例清单见 docs/测试/word测试/检查点与溯源/用例清单.md。
    /// </summary>
    internal static class WkCases
    {
        public const int Target = 46;
        public const int CaretTarget = 2;

        private const string SentA = "易写检查点甲句。";
        private const string SentB = "易写检查点乙句。";
        private const string SentC = "易写检查点丙句。";
        private const string ProvSent = "易写溯源原句。";
        private const string ProvSent2 = "易写溯源乙句。";
        private const string SrcA = "易写溯源说明甲";
        private const string SrcB = "易写溯源说明乙";
        private const string TblCell = "易写巩固表甲";
        private const string TblSrc = "易写表来源";
        private const string ImgSrc = "易写图来源";
        private const string Page2Anchor = "易写检查点第二页锚。";

        public static void Run(
            TestRun run,
            string host,
            int wordBefore,
            int wpsBefore,
            bool wordWinBefore,
            bool wpsWinBefore,
            string caseFilter = "")
        {
            var s = new WdSession(host);
            BurnFirstWpsOpen(s);
            if (IsCaretOnly(caseFilter))
            {
                RunCkptCaret(run, s);
                return;
            }

            RunCkptNeg(run, s);
            RunCkptEmpty(run, s);
            RunCkptPos(run, s);
            RunCkptTwo(run, s);
            RunCkptCaret(run, s);
            RunProvNeg(run, s);
            RunProvEmpty(run, s);
            RunProvSentence(run, s);
            RunProvTwo(run, s);
            RunProvTableImage(run, s);
            RunProvMix(run, s);
            RunProvCkpt(run, s);
            RunNoQuit(run, wordBefore, wpsBefore, wordWinBefore, wpsWinBefore);
        }

        public static bool IsCaretOnly(string caseFilter)
        {
            return !string.IsNullOrEmpty(caseFilter)
                && caseFilter.IndexOf("caret", StringComparison.OrdinalIgnoreCase) >= 0;
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
                s.GetContent(id);
                s.ProcessActions(id, ReplaceOrInsertWarm());
                s.CloseDocByPath(path);
            }
        }

        // ---------- A. 检查点负例 ----------

        private static void RunCkptNeg(TestRun run, WdSession s)
        {
            ExpectNoChannel(run, "wk-neg-no-channel", s.ListOperations(null));
            ExpectFail(run, "wk-neg-unknown-id", s.ListOperations("w999"), "未知");
            ExpectFail(run, "wk-rst-unknown-id", s.RestoreCheckpoint("w999", "1"), "未知");

            string[] ids = { "wk-rst-no-ts", "wk-rst-unk-ts" };
            string path, id;
            if (!TryOpen(run, s, out path, out id, ids))
            {
                return;
            }

            ExpectFail(run, "wk-rst-no-ts", s.RestoreCheckpoint(id, null), "timestamp");
            ExpectFail(run, "wk-rst-unk-ts", s.RestoreCheckpoint(id, "not-a-copy"), "不存在");
            s.CloseDocByPath(path);
        }

        // ---------- B. 空稿 list ----------

        private static void RunCkptEmpty(TestRun run, WdSession s)
        {
            string[] ids = { "wk-list-empty", "wk-list-com-only" };
            string path, id;
            if (!TryOpen(run, s, out path, out id, ids))
            {
                return;
            }

            ToolResult empty = s.ListOperations(id);
            Ok(run, "wk-list-empty",
                empty.Success && Contains(DataText(empty), "尚无操作记录"),
                empty.Success ? "list=" + Snip(DataText(empty)) : (empty.Error ?? ""));

            if (!s.TypePlain(path, SentA))
            {
                FailMany(run, "COM 写甲失败", "wk-list-com-only");
                s.CloseDocByPath(path);
                return;
            }

            s.ActivateDoc(path);
            ToolResult comOnly = s.ListOperations(id);
            Ok(run, "wk-list-com-only",
                comOnly.Success && Contains(DataText(comOnly), "尚无操作记录"),
                comOnly.Success ? "COM 写后不应有副本: " + Snip(DataText(comOnly)) : (comOnly.Error ?? ""));
            s.CloseDocByPath(path);
        }

        // ---------- C. 检查点正例 ----------

        private static void RunCkptPos(TestRun run, WdSession s)
        {
            string[] ids =
            {
                "wk-write-ok", "wk-list-copy", "wk-list-ts", "wk-after-new",
                "wk-rst-ok", "wk-rst-old-back", "wk-rst-new-gone"
            };
            string path, id;
            if (!TryOpen(run, s, out path, out id, ids))
            {
                return;
            }

            if (!s.TypePlain(path, SentA))
            {
                FailMany(run, "COM 写甲失败", ids);
                s.CloseDocByPath(path);
                return;
            }

            s.ActivateDoc(path);
            string body = ReadBody(s, id);
            string code = FirstCode(body, "S");
            if (string.IsNullOrEmpty(code))
            {
                FailMany(run, "回读无甲句编码（body=" + Snip(body) + "）", ids);
                s.CloseDocByPath(path);
                return;
            }

            ToolResult wrote = s.ProcessActions(id, Replace(code, SentB));
            Ok(run, "wk-write-ok", wrote.Success, wrote.Error ?? "改字失败");

            ToolResult listed = s.ListOperations(id);
            string listText = DataText(listed);
            string ts = FirstCopyTs(listText);
            Ok(run, "wk-list-copy",
                listed.Success && Contains(listText, "副本"),
                listed.Success ? "list=" + Snip(listText) : (listed.Error ?? ""));
            Ok(run, "wk-list-ts",
                listed.Success && !string.IsNullOrEmpty(ts),
                listed.Success ? "抄不出 timestamp: " + Snip(listText) : (listed.Error ?? ""));

            string after = ReadBody(s, id);
            Ok(run, "wk-after-new",
                wrote.Success && Contains(after, SentB) && !Contains(after, SentA),
                wrote.Success ? "改后正文不对: " + Snip(after) : (wrote.Error ?? ""));

            if (string.IsNullOrEmpty(ts))
            {
                FailMany(run, "无 timestamp，不能 restore", "wk-rst-ok", "wk-rst-old-back", "wk-rst-new-gone");
                s.CloseDocByPath(path);
                return;
            }

            ToolResult rst = s.RestoreCheckpoint(id, ts);
            Ok(run, "wk-rst-ok",
                rst.Success && Contains(DataText(rst), "已恢复"),
                rst.Success ? "回包=" + Snip(DataText(rst)) : (rst.Error ?? ""));
            string back = ReadBody(s, id);
            Ok(run, "wk-rst-old-back",
                rst.Success && Contains(back, SentA),
                rst.Success ? "甲没回来: " + Snip(back) : (rst.Error ?? ""));
            Ok(run, "wk-rst-new-gone",
                rst.Success && !Contains(back, SentB),
                rst.Success ? "乙还在: " + Snip(back) : (rst.Error ?? ""));
            s.CloseDocByPath(path);
        }

        private static void RunCkptTwo(TestRun run, WdSession s)
        {
            string[] ids = { "wk-two-rst-first" };
            string path, id;
            if (!TryOpen(run, s, out path, out id, ids))
            {
                return;
            }

            if (!s.TypePlain(path, SentA))
            {
                FailMany(run, "COM 写甲失败", ids);
                s.CloseDocByPath(path);
                return;
            }

            s.ActivateDoc(path);
            string codeA = FirstCode(ReadBody(s, id), "S");
            ToolResult first = s.ProcessActions(id, Replace(codeA, SentB));
            string ts1 = FirstCopyTs(DataText(s.ListOperations(id)));
            string codeB = FirstCode(ReadBody(s, id), "S");
            s.ProcessActions(id, Replace(codeB, SentC));
            ToolResult rst = string.IsNullOrEmpty(ts1)
                ? new ToolResult { Success = false, Error = "第一次 list 无副本" }
                : s.RestoreCheckpoint(id, ts1);
            string back = ReadBody(s, id);
            Ok(run, "wk-two-rst-first",
                first.Success && rst.Success
                    && Contains(back, SentA)
                    && !Contains(back, SentB)
                    && !Contains(back, SentC),
                rst.Success ? "两次改后 restore 第一次不对: " + Snip(back) : (rst.Error ?? first.Error ?? ""));
            s.CloseDocByPath(path);
        }

        private static void RunCkptCaret(TestRun run, WdSession s)
        {
            string[] ids = { "wk-rst-caret-page2", "wk-rst-caret-scroll" };
            string path, id;
            if (!TryOpen(run, s, out path, out id, ids))
            {
                return;
            }

            if (!s.TypeLongCheckpointBody(path, SentA, Page2Anchor))
            {
                FailMany(run, "COM 写长文失败", ids);
                s.CloseDocByPath(path);
                return;
            }

            s.ActivateDoc(path);
            if (!s.TryGoToPage(path, 2))
            {
                FailMany(run, "GoTo 第2页失败", ids);
                s.CloseDocByPath(path);
                return;
            }

            int pageBefore;
            int startBefore;
            int scrollBefore;
            if (!s.TryReadCaret(path, out pageBefore, out startBefore, out scrollBefore) || pageBefore < 2)
            {
                FailMany(run, "建副本前光标不在第2页（page=" + pageBefore + "）", ids);
                s.CloseDocByPath(path);
                return;
            }

            string body = ReadBody(s, id);
            string code = FirstCode(body, "S");
            if (string.IsNullOrEmpty(code))
            {
                FailMany(run, "长文回读无甲句编码（body=" + Snip(body) + "）", ids);
                s.CloseDocByPath(path);
                return;
            }

            if (!s.TryGoToPage(path, 2))
            {
                FailMany(run, "改字前再次落到第2页失败", ids);
                s.CloseDocByPath(path);
                return;
            }

            ToolResult wrote = s.ProcessActions(id, Replace(code, SentB));
            if (!wrote.Success)
            {
                FailMany(run, "改字失败: " + (wrote.Error ?? ""), ids);
                s.CloseDocByPath(path);
                return;
            }

            string ts = FirstCopyTs(DataText(s.ListOperations(id)));
            if (string.IsNullOrEmpty(ts))
            {
                FailMany(run, "无 timestamp，不能 restore", ids);
                s.CloseDocByPath(path);
                return;
            }

            s.TryGoToPage(path, 1);
            ToolResult rst = s.RestoreCheckpoint(id, ts);
            int pageAfter;
            int startAfter;
            int scrollAfter;
            bool read = s.TryReadCaret(path, out pageAfter, out startAfter, out scrollAfter);
            string back = ReadBody(s, id);
            Ok(run, "wk-rst-caret-page2",
                rst.Success && read && pageAfter >= 2 && Contains(back, SentA) && !Contains(back, SentB),
                rst.Success
                    ? "回滚后页=" + pageAfter + " start=" + startAfter + "（建副本时页=" + pageBefore
                      + " start=" + startBefore + "）正文=" + Snip(back)
                    : (rst.Error ?? ""));
            Ok(run, "wk-rst-caret-scroll",
                rst.Success && read && pageAfter >= 2 && scrollAfter > 0,
                rst.Success
                    ? "回滚后滚动=" + scrollAfter + "%（建副本时 " + scrollBefore
                      + "%）页=" + pageAfter
                    : (rst.Error ?? ""));
            s.CloseDocByPath(path);
        }

        // ---------- E. 溯源负例 ----------

        private static void RunProvNeg(TestRun run, WdSession s)
        {
            ExpectNoChannel(run, "wv-neg-no-channel",
                s.ApplyDataProvenance(new Dictionary<string, object>
                {
                    ["items"] = new List<object> { SentItem("S_00001", SrcA) }
                }));
            ExpectFail(run, "wv-neg-unknown-id",
                s.ApplyDataProvenance(new Dictionary<string, object>
                {
                    ["channel_id"] = "w999",
                    ["items"] = new List<object> { SentItem("S_00001", SrcA) }
                }),
                "未知");

            string[] ids =
            {
                "wv-neg-no-items", "wv-neg-items-empty", "wv-neg-no-code",
                "wv-neg-s-sources", "wv-neg-t-source",
                "wv-neg-unk-s", "wv-neg-unk-t",
                "wv-neg-bad-type", "wv-neg-web-nourl", "wv-neg-dup"
            };
            string path, id;
            if (!TryOpen(run, s, out path, out id, ids))
            {
                return;
            }

            ExpectFail(run, "wv-neg-no-items",
                s.ApplyDataProvenance(new Dictionary<string, object> { ["channel_id"] = id }),
                "items");
            ExpectFail(run, "wv-neg-items-empty",
                s.ApplyDataProvenance(new Dictionary<string, object>
                {
                    ["channel_id"] = id,
                    ["items"] = new List<object>()
                }),
                "items");
            ExpectFail(run, "wv-neg-no-code",
                s.ApplyDataProvenance(ApplyArgs(id, new List<object>
                {
                    new Dictionary<string, object>()
                })),
                "code");
            ExpectFail(run, "wv-neg-s-sources",
                s.ApplyDataProvenance(ApplyArgs(id, new List<object>
                {
                    new Dictionary<string, object>
                    {
                        ["anchor"] = new Dictionary<string, object> { ["code"] = "S_00001" },
                        ["sources"] = new List<object> { TextSource(SrcA) }
                    }
                })),
                "source");
            ExpectFail(run, "wv-neg-t-source",
                s.ApplyDataProvenance(ApplyArgs(id, new List<object>
                {
                    new Dictionary<string, object>
                    {
                        ["anchor"] = new Dictionary<string, object> { ["code"] = "T_00001" },
                        ["source"] = TextSource(SrcA)
                    }
                })),
                "sources");
            ExpectFail(run, "wv-neg-unk-s",
                s.ApplyDataProvenance(ApplyArgs(id, new List<object> { SentItem("S_99999", SrcA) })),
                "定位");
            ExpectFail(run, "wv-neg-unk-t",
                s.ApplyDataProvenance(ApplyArgs(id, new List<object> { ObjItem("T_99999", TblSrc) })),
                "表格");

            if (!s.TypePlain(path, ProvSent))
            {
                FailMany(run, "COM 写溯源句失败", "wv-neg-bad-type", "wv-neg-web-nourl", "wv-neg-dup");
                s.CloseDocByPath(path);
                return;
            }

            s.ActivateDoc(path);
            string code = FirstCode(ReadBody(s, id), "S");
            if (string.IsNullOrEmpty(code))
            {
                FailMany(run, "回读无溯源句编码", "wv-neg-bad-type", "wv-neg-web-nourl", "wv-neg-dup");
                s.CloseDocByPath(path);
                return;
            }

            ExpectFail(run, "wv-neg-web-nourl",
                s.ApplyDataProvenance(ApplyArgs(id, new List<object>
                {
                    new Dictionary<string, object>
                    {
                        ["anchor"] = new Dictionary<string, object> { ["code"] = FirstCode(ReadBody(s, id), "S") },
                        ["source"] = new Dictionary<string, object> { ["type"] = "web", ["title"] = "t" }
                    }
                })),
                "url");
            ExpectFail(run, "wv-neg-bad-type",
                s.ApplyDataProvenance(ApplyArgs(id, new List<object>
                {
                    new Dictionary<string, object>
                    {
                        ["anchor"] = new Dictionary<string, object> { ["code"] = FirstCode(ReadBody(s, id), "S") },
                        ["source"] = new Dictionary<string, object> { ["type"] = "foo", ["text"] = SrcA }
                    }
                })),
                "type");
            string dupCode = FirstCode(ReadBody(s, id), "S");
            ExpectFail(run, "wv-neg-dup",
                s.ApplyDataProvenance(ApplyArgs(id, new List<object>
                {
                    SentItem(dupCode, SrcA),
                    SentItem(dupCode, SrcB)
                })),
                "重复");
            s.CloseDocByPath(path);
        }

        // ---------- F. 溯源空读 ----------

        private static void RunProvEmpty(TestRun run, WdSession s)
        {
            string[] ids = { "wv-get-empty", "wv-get-no-write" };
            string path, id;
            if (!TryOpen(run, s, out path, out id, ids))
            {
                return;
            }

            ToolResult empty = s.GetDataProvenance(id);
            Ok(run, "wv-get-empty",
                empty.Success && !WdSession.Bool(empty, "has_provenance"),
                empty.Success ? "空稿不应有溯源" : (empty.Error ?? ""));

            if (!s.TypePlain(path, ProvSent))
            {
                FailMany(run, "COM 写溯源句失败", "wv-get-no-write");
                s.CloseDocByPath(path);
                return;
            }

            s.ActivateDoc(path);
            ToolResult got = s.GetDataProvenance(id);
            string body = ReadBody(s, id);
            Ok(run, "wv-get-no-write",
                got.Success && !WdSession.Bool(got, "has_provenance") && Contains(body, ProvSent),
                got.Success ? "未 apply 却有溯源或原句丢了" : (got.Error ?? ""));
            s.CloseDocByPath(path);
        }

        // ---------- G. 溯源句子 ----------

        private static void RunProvSentence(TestRun run, WdSession s)
        {
            string[] ids =
            {
                "wv-s-ok", "wv-s-fields", "wv-s-mark", "wv-s-ref",
                "wv-s-get", "wv-s-cover"
            };
            string path, id;
            if (!TryOpen(run, s, out path, out id, ids))
            {
                return;
            }

            if (!s.TypePlain(path, ProvSent))
            {
                FailMany(run, "COM 写溯源句失败", ids);
                s.CloseDocByPath(path);
                return;
            }

            s.ActivateDoc(path);
            string code = FirstCode(ReadBody(s, id), "S");
            if (string.IsNullOrEmpty(code))
            {
                FailMany(run, "回读无溯源句编码", ids);
                s.CloseDocByPath(path);
                return;
            }

            ToolResult applied = s.ApplyDataProvenance(ApplyArgs(id, new List<object> { SentItem(code, SrcA) }));
            Ok(run, "wv-s-ok", applied.Success, applied.Error ?? "apply 失败");
            Ok(run, "wv-s-fields",
                applied.Success
                    && WdSession.IntOf(applied, "sentence_count") == 1
                    && WdSession.IntOf(applied, "item_count") == 1,
                applied.Success
                    ? "sentence_count=" + WdSession.IntOf(applied, "sentence_count")
                        + " item_count=" + WdSession.IntOf(applied, "item_count")
                    : (applied.Error ?? ""));

            string body = ReadBody(s, id);
            Ok(run, "wv-s-mark",
                applied.Success && Contains(body, ProvSent) && Contains(body, "[1]"),
                applied.Success ? "缺原句或 [1]: " + Snip(body) : (applied.Error ?? ""));
            Ok(run, "wv-s-ref",
                applied.Success && Contains(body, "参考数据来源"),
                applied.Success ? "无参考区: " + Snip(body) : (applied.Error ?? ""));

            ToolResult got = s.GetDataProvenance(id);
            Ok(run, "wv-s-get",
                got.Success && WdSession.Bool(got, "has_provenance") && HasItemN(got, 1),
                got.Success ? "get 无 n=1" : (got.Error ?? ""));

            string code2 = FirstCode(ReadBody(s, id), "S");
            if (string.IsNullOrEmpty(code2))
            {
                code2 = code;
            }

            ToolResult refused = s.ApplyDataProvenance(ApplyArgs(id, new List<object> { SentItem(code2, SrcB) }));
            string still = ReadBody(s, id);
            ToolResult cleared = s.ClearDataProvenance(id);
            string restored = ReadBody(s, id);
            string codeBack = FirstCode(restored, "S");
            ToolResult again = s.ApplyDataProvenance(ApplyArgs(id, new List<object> { SentItem(codeBack, SrcB) }));
            string after = ReadBody(s, id);
            Ok(run, "wv-s-cover",
                !refused.Success
                    && Contains(refused.Error ?? "", "已有引用")
                    && Contains(still, SrcA)
                    && !Contains(still, SrcB)
                    && cleared.Success
                    && codeBack == code
                    && again.Success
                    && Contains(after, SrcB)
                    && !Contains(after, SrcA),
                !refused.Success
                    ? "清后码=" + codeBack + " 原码=" + code + " again=" + (again.Error ?? Snip(after))
                    : ("再 apply 仍成功: " + Snip(still)));
            s.CloseDocByPath(path);
        }

        private static void RunProvTwo(TestRun run, WdSession s)
        {
            string[] ids = { "wv-s-two" };
            string path, id;
            if (!TryOpen(run, s, out path, out id, ids))
            {
                return;
            }

            if (!s.TypePlain(path, ProvSent + ProvSent2))
            {
                FailMany(run, "COM 写两句失败", ids);
                s.CloseDocByPath(path);
                return;
            }

            s.ActivateDoc(path);
            List<string> codes = ListCodes(ReadBody(s, id), "S");
            if (codes.Count < 2)
            {
                FailMany(run, "回读不足两个 S_（n=" + codes.Count + "）", ids);
                s.CloseDocByPath(path);
                return;
            }

            ToolResult applied = s.ApplyDataProvenance(ApplyArgs(id, new List<object>
            {
                SentItem(codes[0], SrcA),
                SentItem(codes[1], SrcB)
            }));
            string body = ReadBody(s, id);
            Ok(run, "wv-s-two",
                applied.Success
                    && WdSession.IntOf(applied, "sentence_count") == 2
                    && Contains(body, "[1]")
                    && Contains(body, "[2]"),
                applied.Success
                    ? "两句溯源不对 count=" + WdSession.IntOf(applied, "sentence_count") + " " + Snip(body)
                    : (applied.Error ?? ""));
            s.CloseDocByPath(path);
        }

        // ---------- H. 表 / 图 ----------

        private static void RunProvTableImage(TestRun run, WdSession s)
        {
            string[] tblIds = { "wv-t-ok", "wv-t-star" };
            string path, id;
            if (TryOpen(run, s, out path, out id, tblIds))
            {
                if (!s.InsertTable(path, 1, 1, new[] { TblCell }))
                {
                    FailMany(run, "COM 建表失败", tblIds);
                    s.CloseDocByPath(path);
                }
                else
                {
                    s.ActivateDoc(path);
                    string tid = FirstTableId(ReadBody(s, id));
                    if (string.IsNullOrEmpty(tid))
                    {
                        FailMany(run, "回读无 T_", tblIds);
                    }
                    else
                    {
                        ToolResult applied = s.ApplyDataProvenance(ApplyArgs(id, new List<object> { ObjItem(tid, TblSrc) }));
                        string body = ReadBody(s, id);
                        Ok(run, "wv-t-ok",
                            applied.Success && WdSession.IntOf(applied, "table_count") == 1,
                            applied.Success ? "table_count=" + WdSession.IntOf(applied, "table_count") : (applied.Error ?? ""));
                        ToolResult againTbl = s.ApplyDataProvenance(ApplyArgs(id, new List<object> { ObjItem(tid, TblSrc) }));
                        string afterTbl = ReadBody(s, id);
                        Ok(run, "wv-t-star",
                            applied.Success
                                && Contains(body, TblCell)
                                && Contains(body, "*")
                                && Contains(body, TblSrc)
                                && !Contains(body, "[1]")
                                && !againTbl.Success
                                && Contains(againTbl.Error ?? "", "已有引用")
                                && Contains(afterTbl, TblSrc),
                            applied.Success ? "表下说明或再写入: " + (againTbl.Error ?? Snip(afterTbl)) : (applied.Error ?? ""));
                    }

                    s.CloseDocByPath(path);
                }
            }

            string[] imgIds = { "wv-i-ok" };
            if (!TryOpen(run, s, out path, out id, imgIds))
            {
                return;
            }

            ToolResult ins = s.InsertImage(new Dictionary<string, object>
            {
                ["channel_id"] = id,
                ["path"] = FixturePng.Ensure()
            });
            string imageId = WdSession.Str(ins, "image_id");
            if (!ins.Success || !imageId.StartsWith("I_", StringComparison.Ordinal))
            {
                FailMany(run, ins.Error ?? "插图无 I_", imgIds);
                s.CloseDocByPath(path);
                return;
            }

            ToolResult img = s.ApplyDataProvenance(ApplyArgs(id, new List<object> { ObjItem(imageId, ImgSrc) }));
            ToolResult againImg = s.ApplyDataProvenance(ApplyArgs(id, new List<object> { ObjItem(imageId, ImgSrc) }));
            Ok(run, "wv-i-ok",
                img.Success
                    && WdSession.IntOf(img, "image_count") == 1
                    && !againImg.Success
                    && Contains(againImg.Error ?? "", "已有引用"),
                img.Success ? (againImg.Error ?? "image_count=" + WdSession.IntOf(img, "image_count")) : (img.Error ?? ""));
            s.CloseDocByPath(path);
        }

        private static void RunProvMix(TestRun run, WdSession s)
        {
            string[] ids = { "wv-mix" };
            string path, id;
            if (!TryOpen(run, s, out path, out id, ids))
            {
                return;
            }

            if (!s.TypePlain(path, ProvSent) || !s.InsertTable(path, 1, 1, new[] { TblCell }))
            {
                FailMany(run, "COM 铺一句+表失败", ids);
                s.CloseDocByPath(path);
                return;
            }

            s.ActivateDoc(path);
            string body = ReadBody(s, id);
            string code = FirstCode(body, "S");
            string tid = FirstTableId(body);
            if (string.IsNullOrEmpty(code) || string.IsNullOrEmpty(tid))
            {
                FailMany(run, "回读缺 S_ 或 T_（body=" + Snip(body) + "）", ids);
                s.CloseDocByPath(path);
                return;
            }

            ToolResult applied = s.ApplyDataProvenance(ApplyArgs(id, new List<object>
            {
                SentItem(code, SrcA),
                ObjItem(tid, TblSrc)
            }));
            string after = ReadBody(s, id);
            Ok(run, "wv-mix",
                applied.Success
                    && WdSession.IntOf(applied, "sentence_count") == 1
                    && WdSession.IntOf(applied, "table_count") == 1
                    && Contains(after, "[1]")
                    && Contains(after, "*"),
                applied.Success
                    ? "混合回包/正文不对 s=" + WdSession.IntOf(applied, "sentence_count")
                        + " t=" + WdSession.IntOf(applied, "table_count") + " " + Snip(after)
                    : (applied.Error ?? ""));
            s.CloseDocByPath(path);
        }

        // ---------- I. 交叉 ----------

        private static void RunProvCkpt(TestRun run, WdSession s)
        {
            string[] ids = { "wv-ckpt-list", "wv-ckpt-rst" };
            string path, id;
            if (!TryOpen(run, s, out path, out id, ids))
            {
                return;
            }

            if (!s.TypePlain(path, ProvSent))
            {
                FailMany(run, "COM 写溯源句失败", ids);
                s.CloseDocByPath(path);
                return;
            }

            s.ActivateDoc(path);
            string code = FirstCode(ReadBody(s, id), "S");
            if (string.IsNullOrEmpty(code))
            {
                FailMany(run, "回读无溯源句编码", ids);
                s.CloseDocByPath(path);
                return;
            }

            ToolResult applied = s.ApplyDataProvenance(ApplyArgs(id, new List<object> { SentItem(code, SrcA) }));
            ToolResult listed = s.ListOperations(id);
            string listText = DataText(listed);
            string ts = LastCopyTs(listText);
            Ok(run, "wv-ckpt-list",
                applied.Success && listed.Success
                    && Contains(listText, "副本")
                    && (Contains(listText, "apply_data_provenance") || Contains(listText, "F_apply_data_provenance")),
                listed.Success ? "list=" + Snip(listText) : (listed.Error ?? applied.Error ?? ""));

            if (!applied.Success || string.IsNullOrEmpty(ts))
            {
                FailMany(run, "apply 失败或无副本，不能 restore: " + (applied.Error ?? Snip(listText)), "wv-ckpt-rst");
                s.CloseDocByPath(path);
                return;
            }

            ToolResult rst = s.RestoreCheckpoint(id, ts);
            string back = ReadBody(s, id);
            Ok(run, "wv-ckpt-rst",
                rst.Success
                    && Contains(back, ProvSent)
                    && !Contains(back, "[1]")
                    && !Contains(back, "参考数据来源"),
                rst.Success ? "restore 后溯源还在: " + Snip(back) : (rst.Error ?? ""));
            s.CloseDocByPath(path);
        }

        // ---------- 哨兵 ----------

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
                run.CaseOk("wk-no-quit");
                run.CaseOk("wv-no-quit");
            }
            else
            {
                run.CaseFail("wk-no-quit", why);
                run.CaseFail("wv-no-quit", why);
            }
        }

        // ---------- 帮手 ----------

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

        private static List<Dictionary<string, object>> Replace(string codes, string neu)
        {
            return new List<Dictionary<string, object>>
            {
                new Dictionary<string, object>
                {
                    ["type"] = "replace",
                    ["detail"] = new Dictionary<string, object>
                    {
                        ["original"] = new Dictionary<string, object> { ["codes"] = codes },
                        ["new"] = neu,
                        ["format"] = DefaultFormat()
                    }
                }
            };
        }

        private static List<Dictionary<string, object>> ReplaceOrInsertWarm()
        {
            return new List<Dictionary<string, object>>
            {
                new Dictionary<string, object>
                {
                    ["type"] = "insert",
                    ["detail"] = new Dictionary<string, object>
                    {
                        ["insert"] = "易写预热句。",
                        ["format"] = DefaultFormat()
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

        private static Dictionary<string, object> ApplyArgs(string channelId, List<object> items)
        {
            return new Dictionary<string, object>
            {
                ["channel_id"] = channelId,
                ["items"] = items
            };
        }

        private static Dictionary<string, object> SentItem(string code, string text)
        {
            return new Dictionary<string, object>
            {
                ["anchor"] = new Dictionary<string, object> { ["code"] = code },
                ["source"] = TextSource(text)
            };
        }

        private static Dictionary<string, object> ObjItem(string code, string text)
        {
            return new Dictionary<string, object>
            {
                ["anchor"] = new Dictionary<string, object> { ["code"] = code },
                ["sources"] = new List<object> { TextSource(text) }
            };
        }

        private static Dictionary<string, object> TextSource(string text)
        {
            return new Dictionary<string, object>
            {
                ["type"] = "text",
                ["text"] = text
            };
        }

        private static string DataText(ToolResult r)
        {
            if (r == null || r.Data == null)
            {
                return "";
            }

            if (r.Data is string s)
            {
                return s;
            }

            try
            {
                return JsonConvert.SerializeObject(r.Data);
            }
            catch
            {
                return r.Data.ToString() ?? "";
            }
        }

        private static string FirstCopyTs(string listText)
        {
            if (string.IsNullOrEmpty(listText))
            {
                return "";
            }

            Match m = Regex.Match(listText, @"副本(\S+)");
            return m.Success ? m.Groups[1].Value : "";
        }

        private static string LastCopyTs(string listText)
        {
            if (string.IsNullOrEmpty(listText))
            {
                return "";
            }

            MatchCollection ms = Regex.Matches(listText, @"副本(\S+)");
            return ms.Count == 0 ? "" : ms[ms.Count - 1].Groups[1].Value;
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

        private static string FirstTableId(string body)
        {
            if (string.IsNullOrEmpty(body))
            {
                return "";
            }

            Match m = Regex.Match(body, @"T_[A-Za-z0-9]+");
            return m.Success ? m.Value : "";
        }

        private static bool HasItemN(ToolResult r, int n)
        {
            Dictionary<string, object> d = r == null ? null : r.Data as Dictionary<string, object>;
            if (d == null || !d.ContainsKey("items") || d["items"] == null)
            {
                return false;
            }

            IEnumerable items = d["items"] as IEnumerable;
            if (items == null)
            {
                return false;
            }

            foreach (object item in items)
            {
                JObject jo = item as JObject ?? JObject.FromObject(item);
                if (jo["n"] != null && jo.Value<int>("n") == n)
                {
                    return true;
                }
            }

            return false;
        }

        private static void ExpectNoChannel(TestRun run, string id, ToolResult r)
        {
            if (!r.Success
                && (Contains(r.Error, "打开")
                    || Contains(r.Error, "unsupported")
                    || Contains(r.Error, "browser")
                    || Contains(r.Error, "渠道")))
            {
                run.CaseOk(id);
            }
            else if (r.Success)
            {
                run.CaseSkip(id, "现场已有默认文字渠，不关用户稿");
            }
            else
            {
                run.CaseFail(id, r.Error ?? "应失败或跳过");
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

        private static string Snip(string body)
        {
            if (string.IsNullOrEmpty(body))
            {
                return "";
            }

            return body.Length <= 80 ? body : body.Substring(0, 80);
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

        private static class FixturePng
        {
            private static string _path;

            public static string Ensure()
            {
                if (_path != null && File.Exists(_path))
                {
                    return _path;
                }

                string path = Path.Combine(
                    Path.GetTempPath(),
                    "ew-wk-" + Guid.NewGuid().ToString("N") + ".png");
                using (var bmp = new System.Drawing.Bitmap(32, 16))
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
