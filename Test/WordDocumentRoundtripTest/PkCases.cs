using System;
using WordAddIn1;

namespace WordDocumentRoundtripTest
{
    /// <summary>PPT / WPP 检查点。清单见 docs/测试/ppt测试/检查点/用例清单.md。</summary>
    internal static class PkCases
    {
        public const int Target = 19;

        public static void Run(TestRun run, string host, bool winBefore)
        {
            var s = new OfficeCkptSession(OfficeCkptSession.NormalizePptHost(host));
            RunNeg(run, s);
            RunEmpty(run, s);
            RunPos(run, s);
            RunTwo(run, s);
            RunCaret(run, s);
            RunContract(run, s);
            CkptCaseUtil.RunNoQuit(run, "pk-no-quit", s.ProcName, winBefore);
        }

        private static void RunNeg(TestRun run, OfficeCkptSession s)
        {
            CkptCaseUtil.ExpectNoChannel(run, "pk-neg-no-channel", s.ListOperations(null));
            CkptCaseUtil.ExpectFail(run, "pk-neg-unknown-id", s.ListOperations("p999"), "未知");
            CkptCaseUtil.ExpectFail(run, "pk-rst-unknown-id", s.RestoreCheckpoint("p999", "1"), "未知");

            string[] ids = { "pk-rst-no-ts", "pk-rst-unk-ts" };
            string path, id;
            if (!CkptCaseUtil.TryOpen(run, s, out path, out id, ids))
            {
                return;
            }

            CkptCaseUtil.ExpectFail(run, "pk-rst-no-ts", s.RestoreCheckpoint(id, null), "timestamp");
            CkptCaseUtil.ExpectFail(run, "pk-rst-unk-ts", s.RestoreCheckpoint(id, "not-a-copy"), "不存在");
            s.CloseByPath(path, id);
        }

        private static void RunEmpty(TestRun run, OfficeCkptSession s)
        {
            string[] ids = { "pk-list-empty", "pk-list-com-only" };
            string path, id;
            if (!CkptCaseUtil.TryOpen(run, s, out path, out id, ids))
            {
                return;
            }

            ToolResult empty = s.ListOperations(id);
            CkptCaseUtil.Ok(
                run,
                "pk-list-empty",
                empty.Success && CkptCaseUtil.Contains(CkptCaseUtil.DataText(empty), "尚无操作记录"),
                empty.Success ? "list=" + CkptCaseUtil.Snip(CkptCaseUtil.DataText(empty)) : (empty.Error ?? ""));

            int before = s.SlideCount(id);
            if (!s.ComAddSlide(path))
            {
                CkptCaseUtil.FailMany(run, "COM 加页失败", "pk-list-com-only");
                s.CloseByPath(path, id);
                return;
            }

            ToolResult comOnly = s.ListOperations(id);
            CkptCaseUtil.Ok(
                run,
                "pk-list-com-only",
                comOnly.Success
                    && CkptCaseUtil.Contains(CkptCaseUtil.DataText(comOnly), "尚无操作记录")
                    && s.SlideCount(id) == before + 1,
                comOnly.Success
                    ? "COM 改后不应有副本: " + CkptCaseUtil.Snip(CkptCaseUtil.DataText(comOnly))
                    : (comOnly.Error ?? ""));
            s.CloseByPath(path, id);
        }

        private static void RunPos(TestRun run, OfficeCkptSession s)
        {
            string[] ids =
            {
                "pk-write-ok", "pk-list-copy", "pk-list-ts", "pk-after-new",
                "pk-rst-ok", "pk-rst-old-back", "pk-rst-new-gone"
            };
            string path, id;
            if (!CkptCaseUtil.TryOpen(run, s, out path, out id, ids))
            {
                return;
            }

            int baseline = Math.Max(1, s.SlideCount(id));
            ToolResult wrote = s.AddSlide(id);
            CkptCaseUtil.Ok(run, "pk-write-ok", wrote.Success, wrote.Error ?? "加页失败");

            ToolResult listed = s.ListOperations(id);
            string listText = CkptCaseUtil.DataText(listed);
            string ts = CkptCaseUtil.FirstCopyTs(listText);
            CkptCaseUtil.Ok(
                run,
                "pk-list-copy",
                listed.Success && CkptCaseUtil.Contains(listText, "副本"),
                listed.Success ? "list=" + CkptCaseUtil.Snip(listText) : (listed.Error ?? ""));
            CkptCaseUtil.Ok(
                run,
                "pk-list-ts",
                listed.Success && !string.IsNullOrEmpty(ts),
                listed.Success ? "抄不出 timestamp: " + CkptCaseUtil.Snip(listText) : (listed.Error ?? ""));
            int after = Math.Max(s.SlideCount(id), WdSession.IntOf(wrote, "slide_count"));
            CkptCaseUtil.Ok(
                run,
                "pk-after-new",
                wrote.Success && after == baseline + 1,
                wrote.Success
                    ? "加后页数=" + after + " 期望 " + (baseline + 1)
                      + " 回包=" + CkptCaseUtil.Snip(CkptCaseUtil.DataText(wrote))
                    : (wrote.Error ?? ""));

            if (string.IsNullOrEmpty(ts))
            {
                CkptCaseUtil.FailMany(run, "无 timestamp，不能 restore", "pk-rst-ok", "pk-rst-old-back", "pk-rst-new-gone");
                s.CloseByPath(path, id);
                return;
            }

            ToolResult rst = s.RestoreCheckpoint(id, ts);
            int back = s.SlideCount(id);
            CkptCaseUtil.Ok(
                run,
                "pk-rst-ok",
                rst.Success && CkptCaseUtil.Contains(CkptCaseUtil.DataText(rst), "已恢复"),
                rst.Success ? "回包=" + CkptCaseUtil.Snip(CkptCaseUtil.DataText(rst)) : (rst.Error ?? ""));
            CkptCaseUtil.Ok(
                run,
                "pk-rst-old-back",
                rst.Success && back == baseline,
                rst.Success ? "页数没回到 " + baseline + "，现在 " + back : (rst.Error ?? ""));
            CkptCaseUtil.Ok(
                run,
                "pk-rst-new-gone",
                rst.Success && back == baseline,
                rst.Success ? "新页还在，页数=" + back : (rst.Error ?? ""));
            s.CloseByPath(path, id);
        }

        private static void RunTwo(TestRun run, OfficeCkptSession s)
        {
            string[] ids = { "pk-two-rst-first" };
            string path, id;
            if (!CkptCaseUtil.TryOpen(run, s, out path, out id, ids))
            {
                return;
            }

            int baseline = Math.Max(1, s.SlideCount(id));
            ToolResult first = s.AddSlide(id);
            string ts1 = CkptCaseUtil.FirstCopyTs(CkptCaseUtil.DataText(s.ListOperations(id)));
            s.AddSlide(id);
            ToolResult rst = string.IsNullOrEmpty(ts1)
                ? new ToolResult { Success = false, Error = "第一次 list 无副本" }
                : s.RestoreCheckpoint(id, ts1);
            int back = s.SlideCount(id);
            CkptCaseUtil.Ok(
                run,
                "pk-two-rst-first",
                first.Success && rst.Success && back == baseline,
                rst.Success ? "两次加页 restore 第一次后页数=" + back : (rst.Error ?? first.Error ?? ""));
            s.CloseByPath(path, id);
        }

        private static void RunCaret(TestRun run, OfficeCkptSession s)
        {
            string[] ids = { "pk-rst-caret-slide2", "pk-rst-caret-scroll" };
            string path, id;
            if (!CkptCaseUtil.TryOpen(run, s, out path, out id, ids))
            {
                return;
            }

            if (!s.ComAddSlide(path) || !s.ComGotoSlide(path, 2))
            {
                CkptCaseUtil.FailMany(run, "COM 加第2页或落到第2页失败", ids);
                s.CloseByPath(path, id);
                return;
            }

            int pageBefore;
            if (!s.TryReadActiveSlide(path, out pageBefore) || pageBefore < 2)
            {
                CkptCaseUtil.FailMany(run, "建副本前不在第2页（" + pageBefore + "）", ids);
                s.CloseByPath(path, id);
                return;
            }

            int baseline = Math.Max(1, s.SlideCount(id));
            ToolResult wrote = s.AddSlide(id);
            if (!wrote.Success)
            {
                CkptCaseUtil.FailMany(run, "加页失败: " + (wrote.Error ?? ""), ids);
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

            s.ComGotoSlide(path, 1);
            ToolResult rst = s.RestoreCheckpoint(id, ts);
            int pageAfter;
            bool read = s.TryReadActiveSlide(path, out pageAfter);
            int back = s.SlideCount(id);
            CkptCaseUtil.Ok(
                run,
                "pk-rst-caret-slide2",
                rst.Success && read && pageAfter >= 2 && back == baseline,
                rst.Success
                    ? "回滚后页=" + pageAfter + " 页数=" + back + " 期望页≥2 页数=" + baseline
                    : (rst.Error ?? ""));
            CkptCaseUtil.Ok(
                run,
                "pk-rst-caret-scroll",
                rst.Success && read && pageAfter >= 2,
                rst.Success ? "回滚后视口页=" + pageAfter : (rst.Error ?? ""));
            s.CloseByPath(path, id);
        }

        private static void RunContract(TestRun run, OfficeCkptSession s)
        {
            string[] ids = { "pk-contract-no-ckpt" };
            string path, id;
            if (!CkptCaseUtil.TryOpen(run, s, out path, out id, ids))
            {
                return;
            }

            ToolResult bad = s.MutatingContractFail(id);
            ToolResult listed = s.ListOperations(id);
            CkptCaseUtil.Ok(
                run,
                "pk-contract-no-ckpt",
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
