using System;
using WordAddIn1;

namespace WordDocumentRoundtripTest
{
    /// <summary>Excel / ET 检查点。清单见 docs/测试/excel测试/检查点/用例清单.md。</summary>
    internal static class XlkCases
    {
        public const int Target = 19;
        private const string SheetA = "易写检查点甲表";
        private const string SheetB = "易写检查点乙表";
        private const string SheetC = "易写检查点丙表";
        private const string FarCell = "A80";

        public static void Run(TestRun run, string host, bool winBefore)
        {
            var s = new OfficeCkptSession(OfficeCkptSession.NormalizeExcelHost(host));
            RunNeg(run, s);
            RunEmpty(run, s);
            RunPos(run, s);
            RunTwo(run, s);
            RunCaret(run, s);
            RunContract(run, s);
            CkptCaseUtil.RunNoQuit(run, "xlk-no-quit", s.ProcName, winBefore);
        }

        private static void RunNeg(TestRun run, OfficeCkptSession s)
        {
            CkptCaseUtil.ExpectNoChannel(run, "xlk-neg-no-channel", s.ListOperations(null));
            CkptCaseUtil.ExpectFail(run, "xlk-neg-unknown-id", s.ListOperations("x999"), "未知");
            CkptCaseUtil.ExpectFail(run, "xlk-rst-unknown-id", s.RestoreCheckpoint("x999", "1"), "未知");

            string[] ids = { "xlk-rst-no-ts", "xlk-rst-unk-ts" };
            string path, id;
            if (!CkptCaseUtil.TryOpen(run, s, out path, out id, ids))
            {
                return;
            }

            CkptCaseUtil.ExpectFail(run, "xlk-rst-no-ts", s.RestoreCheckpoint(id, null), "timestamp");
            CkptCaseUtil.ExpectFail(run, "xlk-rst-unk-ts", s.RestoreCheckpoint(id, "not-a-copy"), "不存在");
            s.CloseByPath(path, id);
        }

        private static void RunEmpty(TestRun run, OfficeCkptSession s)
        {
            string[] ids = { "xlk-list-empty", "xlk-list-com-only" };
            string path, id;
            if (!CkptCaseUtil.TryOpen(run, s, out path, out id, ids))
            {
                return;
            }

            ToolResult empty = s.ListOperations(id);
            CkptCaseUtil.Ok(
                run,
                "xlk-list-empty",
                empty.Success && CkptCaseUtil.Contains(CkptCaseUtil.DataText(empty), "尚无操作记录"),
                empty.Success ? "list=" + CkptCaseUtil.Snip(CkptCaseUtil.DataText(empty)) : (empty.Error ?? ""));

            if (!s.ComAddSheet(path, SheetA))
            {
                CkptCaseUtil.FailMany(run, "COM 加甲表失败", "xlk-list-com-only");
                s.CloseByPath(path, id);
                return;
            }

            ToolResult comOnly = s.ListOperations(id);
            CkptCaseUtil.Ok(
                run,
                "xlk-list-com-only",
                comOnly.Success && CkptCaseUtil.Contains(CkptCaseUtil.DataText(comOnly), "尚无操作记录"),
                comOnly.Success
                    ? "COM 改后不应有副本: " + CkptCaseUtil.Snip(CkptCaseUtil.DataText(comOnly))
                    : (comOnly.Error ?? ""));
            s.CloseByPath(path, id);
        }

        private static void RunPos(TestRun run, OfficeCkptSession s)
        {
            string[] ids =
            {
                "xlk-write-ok", "xlk-list-copy", "xlk-list-ts", "xlk-after-new",
                "xlk-rst-ok", "xlk-rst-old-back", "xlk-rst-new-gone"
            };
            string path, id;
            if (!CkptCaseUtil.TryOpen(run, s, out path, out id, ids))
            {
                return;
            }

            ToolResult wrote = s.AddSheet(id, SheetB);
            CkptCaseUtil.Ok(run, "xlk-write-ok", wrote.Success, wrote.Error ?? "加乙表失败");

            ToolResult listed = s.ListOperations(id);
            string listText = CkptCaseUtil.DataText(listed);
            string ts = CkptCaseUtil.FirstCopyTs(listText);
            CkptCaseUtil.Ok(
                run,
                "xlk-list-copy",
                listed.Success && CkptCaseUtil.Contains(listText, "副本"),
                listed.Success ? "list=" + CkptCaseUtil.Snip(listText) : (listed.Error ?? ""));
            CkptCaseUtil.Ok(
                run,
                "xlk-list-ts",
                listed.Success && !string.IsNullOrEmpty(ts),
                listed.Success ? "抄不出 timestamp: " + CkptCaseUtil.Snip(listText) : (listed.Error ?? ""));
            CkptCaseUtil.Ok(
                run,
                "xlk-after-new",
                wrote.Success && s.OverviewHas(id, SheetB),
                wrote.Success ? "加后概览没有乙表" : (wrote.Error ?? ""));

            if (string.IsNullOrEmpty(ts))
            {
                CkptCaseUtil.FailMany(run, "无 timestamp，不能 restore", "xlk-rst-ok", "xlk-rst-old-back", "xlk-rst-new-gone");
                s.CloseByPath(path, id);
                return;
            }

            ToolResult rst = s.RestoreCheckpoint(id, ts);
            CkptCaseUtil.Ok(
                run,
                "xlk-rst-ok",
                rst.Success && CkptCaseUtil.Contains(CkptCaseUtil.DataText(rst), "已恢复"),
                rst.Success ? "回包=" + CkptCaseUtil.Snip(CkptCaseUtil.DataText(rst)) : (rst.Error ?? ""));
            CkptCaseUtil.Ok(
                run,
                "xlk-rst-old-back",
                rst.Success && !s.OverviewHas(id, SheetB),
                rst.Success ? "乙表还在" : (rst.Error ?? ""));
            CkptCaseUtil.Ok(
                run,
                "xlk-rst-new-gone",
                rst.Success && !s.OverviewHas(id, SheetB),
                rst.Success ? "乙表没撤掉" : (rst.Error ?? ""));
            s.CloseByPath(path, id);
        }

        private static void RunTwo(TestRun run, OfficeCkptSession s)
        {
            string[] ids = { "xlk-two-rst-first" };
            string path, id;
            if (!CkptCaseUtil.TryOpen(run, s, out path, out id, ids))
            {
                return;
            }

            ToolResult first = s.AddSheet(id, SheetB);
            string ts1 = CkptCaseUtil.FirstCopyTs(CkptCaseUtil.DataText(s.ListOperations(id)));
            s.AddSheet(id, SheetC);
            ToolResult rst = string.IsNullOrEmpty(ts1)
                ? new ToolResult { Success = false, Error = "第一次 list 无副本" }
                : s.RestoreCheckpoint(id, ts1);
            CkptCaseUtil.Ok(
                run,
                "xlk-two-rst-first",
                first.Success && rst.Success
                    && !s.OverviewHas(id, SheetB)
                    && !s.OverviewHas(id, SheetC),
                rst.Success ? "两次加表 restore 第一次不对" : (rst.Error ?? first.Error ?? ""));
            s.CloseByPath(path, id);
        }

        private static void RunCaret(TestRun run, OfficeCkptSession s)
        {
            string[] ids = { "xlk-rst-caret-cell", "xlk-rst-caret-scroll" };
            string path, id;
            if (!CkptCaseUtil.TryOpen(run, s, out path, out id, ids))
            {
                return;
            }

            if (!s.ComSelectCell(path, FarCell))
            {
                CkptCaseUtil.FailMany(run, "选中 " + FarCell + " 失败", ids);
                s.CloseByPath(path, id);
                return;
            }

            string before;
            if (!s.TryReadActiveCell(path, out before)
                || !string.Equals(
                    OfficeCkptSession.NormalizeCell(before),
                    FarCell,
                    StringComparison.OrdinalIgnoreCase))
            {
                CkptCaseUtil.FailMany(run, "建副本前活动格不是 " + FarCell + "（" + before + "）", ids);
                s.CloseByPath(path, id);
                return;
            }

            ToolResult wrote = s.AddSheet(id, SheetB);
            if (!wrote.Success)
            {
                CkptCaseUtil.FailMany(run, "加乙表失败: " + (wrote.Error ?? ""), ids);
                s.CloseByPath(path, id);
                return;
            }

            string ts = CkptCaseUtil.FirstCopyTs(CkptCaseUtil.DataText(s.ListOperations(id)));
            if (string.IsNullOrEmpty(ts))
            {
                CkptCaseUtil.FailMany(run, "无 timestamp，不能 restore", ids);
                s.CloseByPath(path, id);
                return;
            }

            s.ComSelectCell(path, "A1");
            ToolResult rst = s.RestoreCheckpoint(id, ts);
            string after;
            bool read = s.TryReadActiveCell(path, out after);
            string cell = OfficeCkptSession.NormalizeCell(after);
            CkptCaseUtil.Ok(
                run,
                "xlk-rst-caret-cell",
                rst.Success && read && string.Equals(cell, FarCell, StringComparison.OrdinalIgnoreCase)
                    && !s.OverviewHas(id, SheetB),
                rst.Success
                    ? "回滚后活动格=" + after + " 期望 " + FarCell
                    : (rst.Error ?? ""));
            CkptCaseUtil.Ok(
                run,
                "xlk-rst-caret-scroll",
                rst.Success && read && string.Equals(cell, FarCell, StringComparison.OrdinalIgnoreCase),
                rst.Success ? "回滚后没回到远格 " + FarCell + "（" + after + "）" : (rst.Error ?? ""));
            s.CloseByPath(path, id);
        }

        private static void RunContract(TestRun run, OfficeCkptSession s)
        {
            string[] ids = { "xlk-contract-no-ckpt" };
            string path, id;
            if (!CkptCaseUtil.TryOpen(run, s, out path, out id, ids))
            {
                return;
            }

            ToolResult bad = s.MutatingContractFail(id);
            ToolResult listed = s.ListOperations(id);
            CkptCaseUtil.Ok(
                run,
                "xlk-contract-no-ckpt",
                !bad.Success && listed.Success
                    && CkptCaseUtil.Contains(CkptCaseUtil.DataText(listed), "尚无操作记录"),
                bad.Success
                    ? "缺 action 应失败却成功"
                    : (listed.Success
                        ? "契约失败后不应有副本: " + CkptCaseUtil.Snip(CkptCaseUtil.DataText(listed))
                        : (listed.Error ?? bad.Error ?? "")));
            s.CloseByPath(path, id);
        }
    }
}
