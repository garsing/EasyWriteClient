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
    /// 批注：列出 / 增加 / 删除。用例清单见 docs/测试/word测试/批注/用例清单.md。
    /// </summary>
    internal static class WnCases
    {
        public const int Target = 34;

        private const string SentA = "易写批注甲句。";
        private const string SentB = "易写批注乙句。";
        private const string SentC = "易写批注丙句。";
        private const string SentD = "易写批注丁句。";
        private const string NoteA = "易写批注意见甲";
        private const string NoteB = "易写批注意见乙";
        private const string NoteC = "易写批注意见丙";
        private const string NoteD = "易写批注意见丁";
        private const string SelWord = "批注甲";

        public static void Run(
            TestRun run,
            string host,
            int wordBefore,
            int wpsBefore,
            bool wordWinBefore,
            bool wpsWinBefore)
        {
            var s = new WdSession(host);
            RunContract(run, s);
            RunAdds(run, s);
            RunPartial(run, s);
            RunRestore(run, s);
            RunDeletes(run, s);
            RunList(run, s);
            RunNoQuit(run, wordBefore, wpsBefore, wordWinBefore, wpsWinBefore);
        }

        private static void RunContract(TestRun run, WdSession s)
        {
            ExpectNoChannel(run, "wn-neg-no-channel", s.ProcessComments(null, Arr(AddItem(NoteA, null))));

            string[] ids =
            {
                "wn-neg-no-comments", "wn-neg-empty", "wn-neg-blank", "wn-neg-mixed",
                "wn-neg-author", "wn-neg-del-extra", "wn-neg-del-index"
            };
            string path, id;
            if (!TryOpen(run, s, out path, out id, ids))
            {
                return;
            }

            string before = Ops(s, id);
            ExpectFail(run, "wn-neg-no-comments", s.ProcessComments(id, null), "comments");
            ExpectFail(run, "wn-neg-empty", s.ProcessComments(id, Arr()), "不能为空");
            ExpectFail(run, "wn-neg-blank", s.ProcessComments(id, Arr(AddItem("  ", null))), "text");
            ExpectFail(
                run,
                "wn-neg-mixed",
                s.ProcessComments(id, Arr(AddItem(NoteA, new Dictionary<string, object>
                {
                    ["codes"] = "S_00001",
                    ["position"] = "start"
                }))),
                "locate_failed");
            var withAuthor = AddItem(NoteA, null);
            withAuthor["author"] = "不是当前用户";
            ExpectFail(run, "wn-neg-author", s.ProcessComments(id, Arr(withAuthor)), "author");
            ExpectFail(
                run,
                "wn-neg-del-extra",
                s.ProcessComments(id, Arr(new Dictionary<string, object>
                {
                    ["type"] = "delete",
                    ["index"] = 1,
                    ["target"] = new Dictionary<string, object> { ["position"] = "start" }
                })),
                "index");
            ToolResult badIndex = s.ProcessComments(id, Arr(DelItem(0)));
            string after = Ops(s, id);
            ToolResult listed = s.ListComments(id, null);
            Ok(run, "wn-neg-del-index",
                badIndex != null
                    && !badIndex.Success
                    && Contains(badIndex.Error, "1")
                    && string.Equals(before, after, StringComparison.Ordinal)
                    && listed.Success
                    && Comments(listed).Count == 0,
                badIndex != null && !badIndex.Success
                    ? "契约后稿或记录变了 list=" + Snip(after)
                    : (badIndex == null ? "应失败" : (badIndex.Error ?? "")));

            s.CloseDocByPath(path);
        }

        private static void RunAdds(TestRun run, WdSession s)
        {
            RunAddSentence(run, s);
            RunAddParagraph(run, s);
            RunAddSelection(run, s);
            RunAddCursor(run, s);
            RunAddPin(run, s);
            RunAddTable(run, s);
            RunAddStart(run, s);
            RunAddCell(run, s);
            RunAddImage(run, s);
        }

        private static void RunAddSentence(TestRun run, WdSession s)
        {
            string path, id;
            if (!TryOpen(run, s, out path, out id, "wn-add-sent"))
            {
                return;
            }

            if (!s.TypePlain(path, SentA))
            {
                FailMany(run, "COM 写甲失败", "wn-add-sent");
                s.CloseDocByPath(path);
                return;
            }

            s.ActivateDoc(path);
            string code = FirstCode(ReadBody(s, id), "S");
            ToolResult added = s.ProcessComments(id, Arr(AddItem(NoteA, TargetCodes(code))));
            JObject data = Obj(added);
            ToolResult listed = s.ListComments(id, null);
            JArray rows = Comments(listed);
            string body = ReadBody(s, id);
            string author = rows.Count == 0 ? "" : (rows[0].Value<string>("author") ?? "").Trim();
            string scope = rows.Count == 0 ? "" : (rows[0].Value<string>("scope_text") ?? "");
            bool point = rows.Count > 0 && rows[0].Value<bool?>("point") == true;
            Ok(run, "wn-add-sent",
                added.Success
                    && Msg(data) == "已处理 1 条批注。"
                    && Anchor(data, 0) == "sentence"
                    && listed.Success
                    && rows.Count == 1
                    && author.Length > 0
                    && scope.IndexOf(SentA.TrimEnd('。'), StringComparison.Ordinal) >= 0
                    && !point
                    && body.IndexOf(SentA, StringComparison.Ordinal) >= 0
                    && body.IndexOf(NoteA, StringComparison.Ordinal) < 0,
                added.Success
                    ? "anchor=" + Anchor(data, 0) + " author=" + author + " scope=" + Snip(scope) + " body=" + Snip(body)
                    : (added.Error ?? ""));
            s.CloseDocByPath(path);
        }

        private static void RunAddParagraph(TestRun run, WdSession s)
        {
            string path, id;
            if (!TryOpen(run, s, out path, out id, "wn-add-para"))
            {
                return;
            }

            if (!s.TypePlain(path, SentA + SentB))
            {
                FailMany(run, "COM 写段失败", "wn-add-para");
                s.CloseDocByPath(path);
                return;
            }

            s.ActivateDoc(path);
            string para = FirstParagraph(ReadBody(s, id, "paragraph"));
            ToolResult added = s.ProcessComments(id, Arr(AddItem(NoteA, new Dictionary<string, object>
            {
                ["paragraph_codes"] = para
            })));
            JArray rows = Comments(s.ListComments(id, null));
            string scope = rows.Count == 0 ? "" : (rows[0].Value<string>("scope_text") ?? "");
            Ok(run, "wn-add-para",
                added.Success
                    && Anchor(Obj(added), 0) == "paragraph"
                    && scope.IndexOf("甲句", StringComparison.Ordinal) >= 0
                    && scope.IndexOf("乙句", StringComparison.Ordinal) >= 0,
                added.Success ? "scope=" + Snip(scope) + " para=" + para : (added.Error ?? ""));
            s.CloseDocByPath(path);
        }

        private static void RunAddSelection(TestRun run, WdSession s)
        {
            string path, id;
            if (!TryOpen(run, s, out path, out id, "wn-add-sel"))
            {
                return;
            }

            if (!s.TypePlain(path, SentA) || !s.ActivateDoc(path) || !s.SelectPlain(path, SelWord, false))
            {
                FailMany(run, "铺选区失败", "wn-add-sel");
                s.CloseDocByPath(path);
                return;
            }

            ToolResult added = s.ProcessComments(id, Arr(AddItem(NoteA, null)));
            JArray rows = Comments(s.ListComments(id, null));
            string scope = rows.Count == 0 ? "" : (rows[0].Value<string>("scope_text") ?? "");
            string body = ReadBody(s, id);
            Ok(run, "wn-add-sel",
                added.Success
                    && Anchor(Obj(added), 0) == "selection"
                    && rows.Count == 1
                    && scope.IndexOf(SelWord, StringComparison.Ordinal) >= 0
                    && scope.IndexOf("易写", StringComparison.Ordinal) < 0
                    && rows[0].Value<bool?>("point") != true
                    && body.IndexOf(SentA, StringComparison.Ordinal) >= 0,
                added.Success ? "scope=" + Snip(scope) : (added.Error ?? ""));
            s.CloseDocByPath(path);
        }

        private static void RunAddCursor(TestRun run, WdSession s)
        {
            string path, id;
            if (!TryOpen(run, s, out path, out id, "wn-add-cursor"))
            {
                return;
            }

            if (!s.TypePlain(path, SentA) || !s.ActivateDoc(path) || !s.SelectPlain(path, SentA, true))
            {
                FailMany(run, "铺光标失败", "wn-add-cursor");
                s.CloseDocByPath(path);
                return;
            }

            ToolResult added = s.ProcessComments(id, Arr(AddItem(NoteA, null)));
            JArray rows = Comments(s.ListComments(id, null));
            string body = ReadBody(s, id);
            Ok(run, "wn-add-cursor",
                added.Success
                    && Anchor(Obj(added), 0) == "cursor"
                    && rows.Count == 1
                    && body.IndexOf(SentA, StringComparison.Ordinal) >= 0
                    && body.IndexOf(NoteA, StringComparison.Ordinal) < 0,
                added.Success
                    ? "anchor=" + Anchor(Obj(added), 0) + " point=" + (rows.Count == 0 ? "?" : rows[0]["point"] + "")
                    : (added.Error ?? ""));
            s.CloseDocByPath(path);
        }

        private static void RunAddPin(TestRun run, WdSession s)
        {
            string path, id;
            if (!TryOpen(run, s, out path, out id, "wn-add-pin"))
            {
                return;
            }

            if (!s.TypePlain(path, SentA + "\r" + SentB) || !s.ActivateDoc(path) || !s.SelectPlain(path, SentA, false))
            {
                FailMany(run, "铺选区失败", "wn-add-pin");
                s.CloseDocByPath(path);
                return;
            }

            ToolResult added = s.ProcessComments(id, Arr(AddItem(NoteA, null), AddItem(NoteB, null)));
            JArray rows = Comments(s.ListComments(id, null));
            bool both = rows.Count == 2
                && (rows[0].Value<string>("scope_text") ?? "").IndexOf("甲句", StringComparison.Ordinal) >= 0
                && (rows[1].Value<string>("scope_text") ?? "").IndexOf("甲句", StringComparison.Ordinal) >= 0;
            Ok(run, "wn-add-pin",
                added.Success && Num(Obj(added), "add_count") == 2 && both,
                added.Success ? "rows=" + rows.Count : (added.Error ?? ""));
            s.CloseDocByPath(path);
        }

        private static void RunAddTable(TestRun run, WdSession s)
        {
            string path, id;
            if (!TryOpen(run, s, out path, out id, "wn-add-table"))
            {
                return;
            }

            if (!s.InsertTable(path, 2, 2, new[] { "易写批注表甲", "", "", "" }))
            {
                FailMany(run, "插表失败", "wn-add-table");
                s.CloseDocByPath(path);
                return;
            }

            s.ActivateDoc(path);
            string body = ReadBody(s, id);
            string table = FirstTableId(body);
            ToolResult added = s.ProcessComments(id, Arr(AddItem(NoteA, new Dictionary<string, object>
            {
                ["table"] = table
            })));
            JArray rows = Comments(s.ListComments(id, null));
            string after = ReadBody(s, id);
            Ok(run, "wn-add-table",
                added.Success
                    && Anchor(Obj(added), 0) == "table"
                    && rows.Count >= 1
                    && rows[rows.Count - 1].Value<bool?>("point") == true
                    && after.IndexOf(table, StringComparison.Ordinal) >= 0,
                added.Success ? "table=" + table + " point=" + (rows.Count == 0 ? "?" : rows[rows.Count - 1]["point"] + "") : (added.Error ?? "无表号 " + Snip(body)));
            s.CloseDocByPath(path);
        }

        private static void RunAddStart(TestRun run, WdSession s)
        {
            string path, id;
            if (!TryOpen(run, s, out path, out id, "wn-add-start"))
            {
                return;
            }

            if (!s.TypePlain(path, SentA))
            {
                FailMany(run, "COM 写甲失败", "wn-add-start");
                s.CloseDocByPath(path);
                return;
            }

            s.ActivateDoc(path);
            ToolResult added = s.ProcessComments(id, Arr(AddItem(NoteA, new Dictionary<string, object>
            {
                ["position"] = "start"
            })));
            JArray rows = Comments(s.ListComments(id, null));
            string body = ReadBody(s, id);
            Ok(run, "wn-add-start",
                added.Success
                    && Anchor(Obj(added), 0) == "document_start"
                    && rows.Count == 1
                    && body.IndexOf(SentA, StringComparison.Ordinal) >= 0
                    && body.IndexOf(NoteA, StringComparison.Ordinal) < 0,
                added.Success
                    ? "anchor=" + Anchor(Obj(added), 0) + " point=" + (rows.Count == 0 ? "?" : rows[0]["point"] + "")
                    : (added.Error ?? ""));
            s.CloseDocByPath(path);
        }

        private static void RunAddCell(TestRun run, WdSession s)
        {
            string path, id;
            if (!TryOpen(run, s, out path, out id, "wn-add-cell"))
            {
                return;
            }

            if (!s.InsertTable(path, 2, 2, new[] { "易写批注表甲", "", "", "" }))
            {
                FailMany(run, "插表失败", "wn-add-cell");
                s.CloseDocByPath(path);
                return;
            }

            s.ActivateDoc(path);
            string cell = FirstEmptyCell(ReadBody(s, id));
            ToolResult added = s.ProcessComments(id, Arr(AddItem(NoteA, TargetCodes(cell))));
            string after = ReadBody(s, id);
            Ok(run, "wn-add-cell",
                added.Success
                    && Anchor(Obj(added), 0) == "empty_cell"
                    && !string.IsNullOrEmpty(cell)
                    && after.IndexOf(NoteA, StringComparison.Ordinal) < 0
                    && after.IndexOf(cell, StringComparison.Ordinal) >= 0,
                added.Success ? "cell=" + cell : (added.Error ?? "无真空格 " + Snip(ReadBody(s, id))));
            s.CloseDocByPath(path);
        }

        private static void RunAddImage(TestRun run, WdSession s)
        {
            string path, id;
            if (!TryOpen(run, s, out path, out id, "wn-add-image"))
            {
                return;
            }

            string png = FixturePng.Ensure();
            ToolResult ins = s.InsertImage(new Dictionary<string, object>
            {
                ["channel_id"] = id,
                ["path"] = png
            });
            string imageId = WdSession.Str(ins, "image_id");
            ToolResult added = s.ProcessComments(id, Arr(AddItem(NoteA, new Dictionary<string, object>
            {
                ["image"] = imageId
            })));
            JArray rows = Comments(s.ListComments(id, null));
            string body = ReadBody(s, id);
            Ok(run, "wn-add-image",
                ins.Success
                    && added.Success
                    && Anchor(Obj(added), 0) == "image"
                    && rows.Count >= 1
                    && rows[rows.Count - 1].Value<bool?>("point") == true
                    && body.IndexOf("<image", StringComparison.Ordinal) >= 0
                    && body.IndexOf(NoteA, StringComparison.Ordinal) < 0,
                !ins.Success ? (ins.Error ?? "插图失败") : (added.Error ?? "anchor=" + Anchor(Obj(added), 0)));
            s.CloseDocByPath(path);
        }

        private static void RunPartial(TestRun run, WdSession s)
        {
            string path, id;
            if (!TryOpen(run, s, out path, out id, "wn-miss-only"))
            {
                return;
            }

            string before = Ops(s, id);
            ToolResult miss = s.ProcessComments(id, Arr(AddItem(NoteA, TargetCodes("S_99999"))));
            Ok(run, "wn-miss-only",
                !miss.Success
                    && Comments(s.ListComments(id, null)).Count == 0
                    && CopyCount(Ops(s, id)) == CopyCount(before),
                miss.Success ? "不存在的码却成功" : (miss.Error ?? ""));
            s.CloseDocByPath(path);

            if (!TryOpen(run, s, out path, out id, "wn-partial", "wn-partial-stays"))
            {
                return;
            }

            if (!s.TypePlain(path, SentA + "\r" + SentB))
            {
                FailMany(run, "COM 写句失败", "wn-partial", "wn-partial-stays");
                s.CloseDocByPath(path);
                return;
            }

            s.ActivateDoc(path);
            List<string> codes = ListCodes(ReadBody(s, id), "S");
            if (codes.Count < 2)
            {
                FailMany(run, "句子不足 " + codes.Count, "wn-partial", "wn-partial-stays");
                s.CloseDocByPath(path);
                return;
            }

            ToolResult added = s.ProcessComments(id, Arr(
                AddItem(NoteA, TargetCodes(codes[0])),
                AddItem(NoteB, TargetCodes(codes[1])),
                AddItem(NoteC, TargetCodes("S_99999"))));
            JObject data = Obj(added);
            JArray results = data["results"] as JArray ?? new JArray();
            JArray rows = Comments(s.ListComments(id, null));
            Ok(run, "wn-partial",
                !added.Success
                    && (added.Error ?? "") == "有批注未处理完成。"
                    && Num(data, "add_count") == 2
                    && results.Count == 3
                    && results[0].Value<bool?>("success") == true
                    && results[1].Value<bool?>("success") == true
                    && results[2].Value<bool?>("success") != true
                    && rows.Count == 2
                    && TextOf(rows, NoteA)
                    && TextOf(rows, NoteB),
                added.Error ?? "部分成功断言失败");

            string ops = Ops(s, id);
            string body = ReadBody(s, id);
            Ok(run, "wn-partial-stays",
                CopyCount(ops) >= 1
                    && !string.IsNullOrEmpty(LastCopyTs(ops))
                    && body.IndexOf(SentA, StringComparison.Ordinal) >= 0
                    && body.IndexOf(SentB, StringComparison.Ordinal) >= 0,
                "ops=" + Snip(ops));
            s.CloseDocByPath(path);
        }

        private static void RunRestore(TestRun run, WdSession s)
        {
            string path, id;
            if (!TryOpen(run, s, out path, out id, "wn-restore-add"))
            {
                return;
            }

            if (!s.TypePlain(path, SentA))
            {
                FailMany(run, "COM 写甲失败", "wn-restore-add");
                s.CloseDocByPath(path);
                return;
            }

            s.ActivateDoc(path);
            string code = FirstCode(ReadBody(s, id), "S");
            ToolResult added = s.ProcessComments(id, Arr(AddItem(NoteA, TargetCodes(code))));
            string ts = LastCopyTs(Ops(s, id));
            ToolResult rst = string.IsNullOrEmpty(ts) ? null : s.RestoreCheckpoint(id, ts);
            JArray rows = rst != null && rst.Success ? Comments(s.ListComments(id, null)) : new JArray();
            string body = ReadBody(s, id);
            Ok(run, "wn-restore-add",
                added.Success && rst != null && rst.Success && rows.Count == 0
                    && body.IndexOf(SentA, StringComparison.Ordinal) >= 0,
                rst == null ? "无副本" : (rst.Error ?? "恢复后仍有批注 " + rows.Count));
            s.CloseDocByPath(path);

            if (!TryOpen(run, s, out path, out id, "wn-restore-del"))
            {
                return;
            }

            if (!s.TypePlain(path, SentA))
            {
                FailMany(run, "COM 写甲失败", "wn-restore-del");
                s.CloseDocByPath(path);
                return;
            }

            s.ActivateDoc(path);
            code = FirstCode(ReadBody(s, id), "S");
            added = s.ProcessComments(id, Arr(AddItem(NoteA, TargetCodes(code))));
            ToolResult deleted = s.ProcessComments(id, Arr(DelItem(1)));
            ts = LastCopyTs(Ops(s, id));
            rst = string.IsNullOrEmpty(ts) ? null : s.RestoreCheckpoint(id, ts);
            rows = rst != null && rst.Success ? Comments(s.ListComments(id, null)) : new JArray();
            body = ReadBody(s, id);
            Ok(run, "wn-restore-del",
                added.Success && deleted.Success && rst != null && rst.Success
                    && TextOf(rows, NoteA)
                    && body.IndexOf(SentA, StringComparison.Ordinal) >= 0,
                rst == null ? "无副本" : (rst.Error ?? deleted.Error ?? "批注没回来"));
            s.CloseDocByPath(path);
        }

        private static void RunDeletes(TestRun run, WdSession s)
        {
            RunDeleteOne(run, s);
            RunDeleteTwoAndFour(run, s);
            RunDeleteTwice(run, s);
            RunDeleteOutOfRange(run, s);
            RunDeleteSameCall(run, s);
        }

        private static void RunDeleteOne(TestRun run, WdSession s)
        {
            string path, id;
            if (!TryOpenFour(run, s, out path, out id, "wn-del-one"))
            {
                return;
            }

            ToolResult deleted = s.ProcessComments(id, Arr(DelItem(2)));
            JArray rows = Comments(s.ListComments(id, null));
            Ok(run, "wn-del-one",
                deleted.Success
                    && Num(Obj(deleted), "delete_count") == 1
                    && rows.Count == 3
                    && TextOf(rows, NoteA) && TextOf(rows, NoteC) && TextOf(rows, NoteD)
                    && !TextOf(rows, NoteB),
                deleted.Success ? "留下 " + Texts(rows) : (deleted.Error ?? ""));
            s.CloseDocByPath(path);
        }

        private static void RunDeleteTwoAndFour(TestRun run, WdSession s)
        {
            string path, id;
            if (!TryOpenFour(run, s, out path, out id, "wn-del-24"))
            {
                return;
            }

            ToolResult deleted = s.ProcessComments(id, Arr(DelItem(2), DelItem(4)));
            JArray rows = Comments(s.ListComments(id, null));
            Ok(run, "wn-del-24",
                deleted.Success
                    && rows.Count == 2
                    && TextOf(rows, NoteA) && TextOf(rows, NoteC)
                    && !TextOf(rows, NoteB) && !TextOf(rows, NoteD),
                deleted.Success ? "留下 " + Texts(rows) : (deleted.Error ?? ""));
            s.CloseDocByPath(path);
        }

        private static void RunDeleteTwice(TestRun run, WdSession s)
        {
            string path, id;
            if (!TryOpen(run, s, out path, out id, "wn-del-twice"))
            {
                return;
            }

            if (!AddNotes(s, path, id, SentA + "\r" + SentB + "\r" + SentC, NoteA, NoteB, NoteC))
            {
                FailMany(run, "铺三条失败", "wn-del-twice");
                s.CloseDocByPath(path);
                return;
            }

            ToolResult deleted = s.ProcessComments(id, Arr(DelItem(2), DelItem(2)));
            JObject data = Obj(deleted);
            JArray results = data["results"] as JArray ?? new JArray();
            JArray rows = Comments(s.ListComments(id, null));
            Ok(run, "wn-del-twice",
                !deleted.Success
                    && Num(data, "delete_count") == 1
                    && results.Count == 2
                    && results[1].Value<bool?>("success") != true
                    && rows.Count == 2
                    && TextOf(rows, NoteA) && TextOf(rows, NoteC)
                    && !TextOf(rows, NoteB),
                deleted.Error ?? Texts(rows));
            s.CloseDocByPath(path);
        }

        private static void RunDeleteOutOfRange(TestRun run, WdSession s)
        {
            string path, id;
            if (!TryOpen(run, s, out path, out id, "wn-del-oob"))
            {
                return;
            }

            if (!AddNotes(s, path, id, SentA + "\r" + SentB, NoteA, NoteB))
            {
                FailMany(run, "铺两条失败", "wn-del-oob");
                s.CloseDocByPath(path);
                return;
            }

            string before = Ops(s, id);
            ToolResult deleted = s.ProcessComments(id, Arr(DelItem(5)));
            JArray rows = Comments(s.ListComments(id, null));
            Ok(run, "wn-del-oob",
                !deleted.Success
                    && rows.Count == 2
                    && CopyCount(Ops(s, id)) == CopyCount(before),
                deleted.Success ? "超出序号却成功" : "留下 " + rows.Count + " 副本 " + Snip(Ops(s, id)));
            s.CloseDocByPath(path);
        }

        private static void RunDeleteSameCall(TestRun run, WdSession s)
        {
            string path, id;
            if (!TryOpen(run, s, out path, out id, "wn-del-same-call"))
            {
                return;
            }

            s.ActivateDoc(path);
            ToolResult added = s.ProcessComments(id, Arr(
                AddItem(NoteA, new Dictionary<string, object> { ["position"] = "start" }),
                DelItem(1)));
            JObject data = Obj(added);
            JArray rows = Comments(s.ListComments(id, null));
            Ok(run, "wn-del-same-call",
                !added.Success
                    && Num(data, "add_count") == 1
                    && Num(data, "delete_count") == 0
                    && TextOf(rows, NoteA)
                    && CopyCount(Ops(s, id)) >= 1,
                added.Error ?? "同一次加删不符合");
            s.CloseDocByPath(path);
        }

        private static void RunList(TestRun run, WdSession s)
        {
            string path, id;
            if (!TryOpen(run, s, out path, out id, "wn-list-empty"))
            {
                return;
            }

            ToolResult empty = s.ListComments(id, null);
            Ok(run, "wn-list-empty",
                empty.Success
                    && Msg(Obj(empty)) == "当前稿没有批注。"
                    && Comments(empty).Count == 0
                    && Ops(s, id).IndexOf("尚无操作记录", StringComparison.Ordinal) >= 0,
                empty.Success ? Msg(Obj(empty)) + " / " + Snip(Ops(s, id)) : (empty.Error ?? ""));
            s.CloseDocByPath(path);

            if (!TryOpen(run, s, out path, out id, "wn-list-all"))
            {
                return;
            }

            if (!AddNotes(s, path, id, SentA + "\r" + SentB + "\r" + SentC, NoteA, NoteB, NoteC))
            {
                FailMany(run, "铺三条失败", "wn-list-all");
                s.CloseDocByPath(path);
                return;
            }

            ToolResult all = s.ListComments(id, null);
            JArray rows = Comments(all);
            bool noCodes = true;
            bool fields = rows.Count == 3;
            for (int i = 0; i < rows.Count; i++)
            {
                JObject row = (JObject)rows[i];
                if (row["codes"] != null)
                {
                    noCodes = false;
                }

                fields = fields
                    && row.Value<int?>("index") == i + 1
                    && row["author"] != null
                    && row["scope_text"] != null
                    && row["point"] != null
                    && row["replies"] != null;
            }

            Ok(run, "wn-list-all",
                all.Success
                    && Msg(Obj(all)) == "已列出 3 条批注。"
                    && fields
                    && noCodes
                    && TextAt(rows, 0, NoteA)
                    && TextAt(rows, 1, NoteB)
                    && TextAt(rows, 2, NoteC),
                all.Success ? Texts(rows) : (all.Error ?? ""));
            s.CloseDocByPath(path);

            if (!TryOpen(run, s, out path, out id, "wn-list-filter"))
            {
                return;
            }

            if (!s.TypePlain(path, SentA + SentB + "\r" + SentC))
            {
                FailMany(run, "COM 写段失败", "wn-list-filter");
                s.CloseDocByPath(path);
                return;
            }

            s.ActivateDoc(path);
            string body = ReadBody(s, id);
            string paraBody = ReadBody(s, id, "paragraph");
            List<string> sentences = ListCodes(body, "S");
            List<string> paras = Paragraphs(paraBody);
            ToolResult added = null;
            ToolResult filtered = null;
            if (sentences.Count < 3 || paras.Count < 2)
            {
                FailMany(run, "段/句不足 s=" + sentences.Count + " p=" + paras.Count + " " + Snip(paraBody), "wn-list-filter");
            }
            else
            {
            added = s.ProcessComments(id, Arr(
                AddItem(NoteA, TargetCodes(sentences[0])),
                AddItem(NoteB, TargetCodes(sentences[1])),
                AddItem(NoteC, TargetCodes(sentences[2]))));
            filtered = s.ListComments(id, new Dictionary<string, object>
            {
                ["paragraph_codes"] = paras[1]
            });
            rows = Comments(filtered);
            Ok(run, "wn-list-filter",
                added.Success
                    && filtered.Success
                    && rows.Count == 1
                    && rows[0].Value<int?>("index") == 3
                    && TextAt(rows, 0, NoteC)
                    && Msg(Obj(filtered)) == "已列出 1 条批注。",
                filtered.Success ? "index=" + (rows.Count == 0 ? "?" : rows[0]["index"] + "") + " " + Texts(rows) : (filtered.Error ?? added.Error ?? ""));
            }

            s.CloseDocByPath(path);

            if (!TryOpen(run, s, out path, out id, "wn-list-filter-empty"))
            {
                return;
            }

            if (!s.TypePlain(path, SentA + "\r" + SentB))
            {
                FailMany(run, "COM 写段失败", "wn-list-filter-empty");
                s.CloseDocByPath(path);
                return;
            }

            s.ActivateDoc(path);
            body = ReadBody(s, id);
            paraBody = ReadBody(s, id, "paragraph");
            sentences = ListCodes(body, "S");
            paras = Paragraphs(paraBody);
            if (sentences.Count < 1 || paras.Count < 2)
            {
                FailMany(run, "段不足 s=" + sentences.Count + " p=" + paras.Count + " " + Snip(paraBody), "wn-list-filter-empty");
            }
            else
            {
            added = s.ProcessComments(id, Arr(AddItem(NoteA, TargetCodes(sentences[0]))));
            filtered = s.ListComments(id, new Dictionary<string, object>
            {
                ["paragraph_codes"] = paras[1]
            });
            ToolResult again = s.ListComments(id, null);
            Ok(run, "wn-list-filter-empty",
                added.Success
                    && filtered.Success
                    && Comments(filtered).Count == 0
                    && Msg(Obj(filtered)) == "该范围里没有批注。"
                    && Comments(again).Count == 1,
                filtered.Success ? Msg(Obj(filtered)) : (filtered.Error ?? ""));
            }

            s.CloseDocByPath(path);

            if (!TryOpen(run, s, out path, out id, "wn-list-point"))
            {
                return;
            }

            if (!s.TypePlain(path, SentA) || !s.ActivateDoc(path) || !s.SelectPlain(path, SentA, true))
            {
                FailMany(run, "铺光标失败", "wn-list-point");
                s.CloseDocByPath(path);
                return;
            }

            added = s.ProcessComments(id, Arr(AddItem(NoteA, null)));
            rows = Comments(s.ListComments(id, null));
            string scope = rows.Count == 0 ? "x" : (rows[0].Value<string>("scope_text") ?? "");
            string pointBody = ReadBody(s, id);
            bool point = rows.Count > 0 && rows[0].Value<bool?>("point") == true;
            Ok(run, "wn-list-point",
                added.Success
                    && Anchor(Obj(added), 0) == "cursor"
                    && rows.Count == 1
                    && pointBody.IndexOf(SentA, StringComparison.Ordinal) >= 0
                    && pointBody.IndexOf(NoteA, StringComparison.Ordinal) < 0
                    && (point
                        ? string.IsNullOrWhiteSpace(scope)
                        : scope.IndexOf("甲", StringComparison.Ordinal) >= 0),
                added.Success ? "anchor=" + Anchor(Obj(added), 0) + " scope=" + Snip(scope) : (added.Error ?? ""));
            s.CloseDocByPath(path);

            if (!TryOpen(run, s, out path, out id, "wn-list-no-cp"))
            {
                return;
            }

            if (!s.TypePlain(path, SentA))
            {
                FailMany(run, "COM 写甲失败", "wn-list-no-cp");
                s.CloseDocByPath(path);
                return;
            }

            s.ActivateDoc(path);
            string opsBefore = Ops(s, id);
            if (!s.AddComComment(path, SentA, NoteA))
            {
                FailMany(run, "COM 批注失败", "wn-list-no-cp");
                s.CloseDocByPath(path);
                return;
            }

            ToolResult listed = s.ListComments(id, null);
            Ok(run, "wn-list-no-cp",
                listed.Success
                    && Comments(listed).Count == 1
                    && opsBefore.IndexOf("尚无操作记录", StringComparison.Ordinal) >= 0
                    && Ops(s, id).IndexOf("尚无操作记录", StringComparison.Ordinal) >= 0,
                listed.Success ? Snip(Ops(s, id)) : (listed.Error ?? ""));
            s.CloseDocByPath(path);
        }

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
                run.CaseOk("wn-no-quit");
            }
            else
            {
                run.CaseFail("wn-no-quit", why);
            }
        }

        private static bool TryOpenFour(TestRun run, WdSession s, out string path, out string id, string caseId)
        {
            if (!TryOpen(run, s, out path, out id, caseId))
            {
                return false;
            }

            if (!AddNotes(s, path, id, SentA + "\r" + SentB + "\r" + SentC + "\r" + SentD, NoteA, NoteB, NoteC, NoteD))
            {
                FailMany(run, "铺四条失败", caseId);
                s.CloseDocByPath(path);
                return false;
            }

            return true;
        }

        private static bool AddNotes(WdSession s, string path, string id, string plain, params string[] notes)
        {
            if (!s.TypePlain(path, plain))
            {
                return false;
            }

            s.ActivateDoc(path);
            List<string> codes = ListCodes(ReadBody(s, id), "S");
            if (codes.Count < notes.Length)
            {
                return false;
            }

            var items = new List<object>();
            for (int i = 0; i < notes.Length; i++)
            {
                items.Add(AddItem(notes[i], TargetCodes(codes[i])));
            }

            ToolResult added = s.ProcessComments(id, items);
            return added.Success && Comments(s.ListComments(id, null)).Count == notes.Length;
        }

        private static Dictionary<string, object> AddItem(string text, Dictionary<string, object> target)
        {
            var item = new Dictionary<string, object>
            {
                ["type"] = "add",
                ["text"] = text
            };
            if (target != null)
            {
                item["target"] = target;
            }

            return item;
        }

        private static Dictionary<string, object> DelItem(int index)
        {
            return new Dictionary<string, object>
            {
                ["type"] = "delete",
                ["index"] = index
            };
        }

        private static Dictionary<string, object> TargetCodes(string codes)
        {
            return new Dictionary<string, object> { ["codes"] = codes ?? "" };
        }

        private static List<object> Arr(params object[] items)
        {
            return new List<object>(items ?? new object[0]);
        }

        private static bool TryOpen(TestRun run, WdSession s, out string path, out string id, params string[] ids)
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

        private static string ReadBody(WdSession s, string id, string codeLevel = null)
        {
            ToolResult got = s.GetContent(id, codeLevel);
            return got.Success ? WdSession.Str(got, "document_content") : "";
        }

        private static string Ops(WdSession s, string id)
        {
            return DataText(s.ListOperations(id));
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

        private static JObject Obj(ToolResult r)
        {
            if (r == null || r.Data == null)
            {
                return new JObject();
            }

            try
            {
                return JObject.Parse(JsonConvert.SerializeObject(r.Data));
            }
            catch
            {
                return new JObject();
            }
        }

        private static JArray Comments(ToolResult r)
        {
            JToken token = Obj(r)["comments"];
            return token as JArray ?? new JArray();
        }

        private static string Msg(JObject data)
        {
            return data == null ? "" : (data.Value<string>("message") ?? "");
        }

        private static int Num(JObject data, string key)
        {
            if (data == null || data[key] == null)
            {
                return -1;
            }

            return data.Value<int?>(key) ?? -1;
        }

        private static string Anchor(JObject data, int index)
        {
            JArray results = data == null ? null : data["results"] as JArray;
            if (results == null || index < 0 || index >= results.Count)
            {
                return "";
            }

            return results[index].Value<string>("anchor") ?? "";
        }

        private static bool TextOf(JArray rows, string note)
        {
            if (rows == null)
            {
                return false;
            }

            foreach (JToken row in rows)
            {
                string text = row.Value<string>("text") ?? "";
                if (text.IndexOf(note, StringComparison.Ordinal) >= 0)
                {
                    return true;
                }
            }

            return false;
        }

        private static bool TextAt(JArray rows, int index, string note)
        {
            if (rows == null || index < 0 || index >= rows.Count)
            {
                return false;
            }

            string text = rows[index].Value<string>("text") ?? "";
            return text.IndexOf(note, StringComparison.Ordinal) >= 0;
        }

        private static string Texts(JArray rows)
        {
            if (rows == null)
            {
                return "";
            }

            var parts = new List<string>();
            foreach (JToken row in rows)
            {
                parts.Add(row.Value<string>("text") ?? "");
            }

            return string.Join(" | ", parts);
        }

        private static int CopyCount(string listText)
        {
            if (string.IsNullOrEmpty(listText))
            {
                return 0;
            }

            return Regex.Matches(listText, "副本").Count;
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

        private static List<string> Paragraphs(string text)
        {
            var list = new List<string>();
            if (string.IsNullOrEmpty(text))
            {
                return list;
            }

            foreach (Match m in Regex.Matches(text, @"\[(P_[0-9A-Za-z]+) start\]"))
            {
                list.Add(m.Groups[1].Value);
            }

            return list;
        }

        private static string FirstParagraph(string text)
        {
            List<string> list = Paragraphs(text);
            return list.Count == 0 ? "" : list[0];
        }

        private static string FirstTableId(string body)
        {
            if (string.IsNullOrEmpty(body))
            {
                return "";
            }

            Match m = Regex.Match(body, @"<table\s+id=""(T_[^""]+)""", RegexOptions.IgnoreCase);
            if (m.Success)
            {
                return m.Groups[1].Value;
            }

            m = Regex.Match(body, @"T_[A-Za-z0-9]+");
            return m.Success ? m.Value : "";
        }

        private static string FirstEmptyCell(string body)
        {
            if (string.IsNullOrEmpty(body))
            {
                return "";
            }

            Match m = Regex.Match(
                body,
                @"\[(T_[0-9A-Za-z]+_r[0-9]+_c[0-9]+)-start\](?:\[EMPTY\])?\[T_[0-9A-Za-z]+_r[0-9]+_c[0-9]+-end\]");
            return m.Success ? m.Groups[1].Value : "";
        }

        private static void ExpectNoChannel(TestRun run, string id, ToolResult r)
        {
            if (!r.Success
                && (Contains(r.Error, "打开")
                    || Contains(r.Error, "unsupported")
                    || Contains(r.Error, "browser")
                    || Contains(r.Error, "渠道")
                    || Contains(r.Error, "PPT")
                    || Contains(r.Error, "Excel")))
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
            if (r != null && !r.Success && Contains(r.Error, needle))
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

        private static string Snip(string body)
        {
            if (string.IsNullOrEmpty(body))
            {
                return "";
            }

            return body.Length <= 120 ? body : body.Substring(0, 120);
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

                string path = Path.Combine(Path.GetTempPath(), "ew-wn-" + Guid.NewGuid().ToString("N") + ".png");
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
