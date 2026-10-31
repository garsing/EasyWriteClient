using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using WordAddIn1;
using WordAddIn1.DocumentMapping;

namespace WordDocumentRoundtripTest
{
    /// <summary>
    /// 读文列表编号进 display。对照 docs/测试/word测试/读文列表编号进display/用例清单.md。
    /// </summary>
    internal static class AnCases
    {
        public const int Target = 127;

        public static string IdFilter;

        private static readonly Regex SpanRe = new Regex(
            @"\[(S|P)_([0-9a-zA-Z]{5}(?:_[1-9][0-9]*)?)[^\]]*\](.*?)\[\1_\2[^\]]*\]",
            RegexOptions.Singleline);

        public static void Run(
            TestRun run,
            string host,
            int wordBefore,
            int wpsBefore,
            bool wordWinBefore,
            bool wpsWinBefore)
        {
            var s = new WdSession(host);
            string warm = WarmWps(s);
            try
            {
                RunRead(run, s);
                RunMerge(run, s);
                RunRepCopy(run, s);
                RunRepCount(run, s);
                RunRepStory(run, s);
                RunInsert(run, s);
                RunD39(run, s);
                RunReread(run, s);
                RunMore(run, s);
                RunNoQuit(run, wordBefore, wpsBefore, wordWinBefore, wpsWinBefore);
            }
            finally
            {
                if (!string.IsNullOrEmpty(warm))
                {
                    s.CloseDocByPath(warm);
                }
            }
        }

        private static string WarmWps(WdSession s)
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
            return path;
        }

        private static void RunRead(TestRun run, WdSession s)
        {
            ReadDecimal(run, s, "an-read-decimal-2", 2);
            ReadDecimal(run, s, "an-read-decimal-3", 3);
            ReadDecimal(run, s, "an-read-decimal-10", 10);
            ReadBullet(run, s, "an-read-bullet-2", 2);
            ReadBullet(run, s, "an-read-bullet-3", 3);

            Plant p;
            if (TryPlant(run, s, "1. 不是列表\r", null, out p, "an-read-hand"))
            {
                string inner = FirstInner(p.Body, "S");
                Ok(run, "an-read-hand",
                    NoTags(p.Body)
                    && inner.IndexOf("1. 不是列表", StringComparison.Ordinal) >= 0
                    && StoryHas(p, "1. 不是列表")
                    && !AnyNumbered(p),
                    Snip(p.Body));
                s.CloseDocByPath(p.Path);
            }

            if (TryPlant(run, s, "1. aa\r2. bb\r", null, out p, "an-read-hand-2"))
            {
                Ok(run, "an-read-hand-2",
                    NoTags(p.Body)
                    && p.Body.IndexOf("1. aa", StringComparison.Ordinal) >= 0
                    && p.Body.IndexOf("2. bb", StringComparison.Ordinal) >= 0
                    && StoryHas(p, "1. aa")
                    && StoryHas(p, "2. bb")
                    && !AnyNumbered(p),
                    Snip(p.Body));
                s.CloseDocByPath(p.Path);
            }

            if (TryPlant(run, s, "a\rx\rb\r", (ss, path) =>
                ss.ApplyNumberDefault(path, 1, 3) && ss.ClearParaKeepList(path, 2),
                out p, "an-read-empty-num"))
            {
                Ok(run, "an-read-empty-num",
                    NoTags(p.Body) && CountNumbered(p) >= 2,
                    "空编号段应仍有 ListFormat。 " + Snip(p.Body));
                s.CloseDocByPath(p.Path);
            }

            if (TryPlant(run, s, "a\r\rb\r", (ss, path) => ss.ApplyNumberDefault(path, 1, 1)
                && ss.ApplyNumberDefault(path, 3, 3),
                out p, "an-read-empty-plain"))
            {
                List<ParaSnap> paras = ContentParas(s, p.Path);
                bool midPlain = paras.Count >= 3 && !paras[1].Numbered;
                Ok(run, "an-read-empty-plain",
                    NoTags(p.Body) && midPlain && CountNumbered(p) >= 2,
                    Snip(p.Body));
                s.CloseDocByPath(p.Path);
            }

            if (TryPlant(run, s, JoinN(2), (ss, path) => ss.ApplyNumberDefault(path, 1, 2),
                out p, "an-read-para"))
            {
                ToolResult got = s.GetContent(p.Id, "paragraph");
                string body = got.Success ? WdSession.Str(got, "document_content") : "";
                string inner = FirstInner(body, "P");
                Ok(run, "an-read-para",
                    got.Success && NoTags(body) && HasDecimalPrefix(inner),
                    got.Success ? Snip(body) : (got.Error ?? "段级读失败"));
                s.CloseDocByPath(p.Path);
            }

            if (TryPlant(run, s, "标题段\r", (ss, path) =>
                {
                    ss.TryApplyHeading1(path, 1);
                    return true;
                },
                out p, "an-read-style"))
            {
                string ls = "";
                List<string> lists = s.ParaListStrings(p.Path);
                if (lists != null && lists.Count > 0)
                {
                    ls = (lists[0] ?? "").Trim();
                }

                bool windowNum = ls.Length > 0;
                bool displayNum = AnyNumbered(p) || HasDecimalPrefix(FirstInner(p.Body, "S"));
                Ok(run, "an-read-style",
                    NoTags(p.Body) && displayNum == windowNum,
                    windowNum
                        ? (displayNum ? Snip(p.Body) : "窗口有编号但 display 无前缀 " + Snip(p.Body))
                        : (displayNum ? "窗口无编号但 display 有前缀 " + Snip(p.Body) : "窗口无编号，display 亦无"));

                s.CloseDocByPath(p.Path);
            }

            if (TryPlant(run, s, "a\rb\rhello\rx\r", (ss, path) =>
                ss.ApplyNumberDefault(path, 1, 2) && ss.ApplyNumberDefault(path, 4, 4),
                out p, "an-read-list-break"))
            {
                bool helloPlain = InnerOfStory(p, "hello").IndexOf("hello", StringComparison.Ordinal) >= 0
                    && !HasDecimalPrefix(InnerOfStory(p, "hello"));
                bool xNum = HasDecimalPrefix(InnerOfStory(p, "x")) || ParaOfStory(p, "x").Numbered;
                Ok(run, "an-read-list-break",
                    NoTags(p.Body) && helloPlain && xNum,
                    Snip(p.Body));
                s.CloseDocByPath(p.Path);
            }

            if (TryPlant(run, s, "甲！乙！\r", (ss, path) => ss.ApplyNumberDefault(path, 1, 1),
                out p, "an-read-two-sent"))
            {
                List<string> inners = Inners(p.Body, "S");
                bool firstNum = inners.Count >= 1 && HasDecimalPrefix(inners[0]);
                bool secondPlain = inners.Count >= 2 && !HasDecimalPrefix(inners[1])
                    && inners[1].IndexOf("乙！", StringComparison.Ordinal) >= 0;
                Ok(run, "an-read-two-sent",
                    NoTags(p.Body) && firstNum && secondPlain && !StoryHas(p, "1. 甲"),
                    Snip(p.Body));
                s.CloseDocByPath(p.Path);
            }

            if (TryPlantTable(run, s, out p, "an-read-table"))
            {
                string tid = TableId(p.Body);
                string inner = CellInner(p.Body, "格甲");
                Ok(run, "an-read-table",
                    NoTags(p.Body) && !string.IsNullOrEmpty(tid) && HasDecimalPrefix(StripInTable(inner)),
                    Snip(p.Body));
                s.CloseDocByPath(p.Path);
            }

            if (TryPlant(run, s, "a\rb\rhello\rx\ry\r", (ss, path) =>
                ss.ApplyNumberDefault(path, 1, 2) && ss.ApplyBulletDefault(path, 4, 5),
                out p, "an-read-mixed-lists"))
            {
                bool helloPlain = !HasAutoPrefix(InnerOfStory(p, "hello"));
                bool aNum = HasDecimalPrefix(InnerOfStory(p, "a"));
                bool xBlt = HasBulletPrefix(InnerOfStory(p, "x")) || ParaOfStory(p, "x").Numbered;
                Ok(run, "an-read-mixed-lists",
                    NoTags(p.Body) && helloPlain && aNum && xBlt,
                    Snip(p.Body));
                s.CloseDocByPath(p.Path);
            }

            if (TryPlant(run, s, "a\r", (ss, path) =>
                ss.ApplyNumberDefault(path, 1, 1)
                && ss.InsertPageBreak(path)
                && ss.AppendText(path, "b\r")
                && ss.ApplyNumberDefault(path, ss.ParaStories(path).Count > 2 ? ss.ParaStories(path).Count - 1 : 2, ss.ParaStories(path).Count > 2 ? ss.ParaStories(path).Count - 1 : 2),
                out p, "an-read-after-page"))
            {
                Ok(run, "an-read-after-page",
                    NoTags(p.Body)
                    && HasDecimalPrefix(InnerOfStory(p, "a"))
                    && (HasDecimalPrefix(InnerOfStory(p, "b")) || p.Body.IndexOf("b", StringComparison.Ordinal) >= 0),
                    Snip(p.Body));
                s.CloseDocByPath(p.Path);
            }
        }

        private static void ReadDecimal(TestRun run, WdSession s, string id, int n)
        {
            Plant p;
            if (!TryPlant(run, s, JoinN(n), (ss, path) => ss.ApplyNumberDefault(path, 1, n), out p, id))
            {
                return;
            }

            int seen = 0;
            foreach (string inner in Inners(p.Body, "S"))
            {
                if (HasDecimalPrefix(inner))
                {
                    seen++;
                }
            }

            bool storyClean = true;
            foreach (ParaSnap para in ContentParas(s, p.Path))
            {
                if (para.Numbered && HasDecimalPrefix(para.Story))
                {
                    storyClean = false;
                }
            }

            Ok(run, id,
                NoTags(p.Body) && seen >= n && storyClean && CountNumbered(p) >= n,
                "display 编号 " + seen + "/" + n + " " + Snip(p.Body));
            s.CloseDocByPath(p.Path);
        }

        private static void ReadBullet(TestRun run, WdSession s, string id, int n)
        {
            Plant p;
            if (!TryPlant(run, s, JoinN(n), (ss, path) => ss.ApplyBulletDefault(path, 1, n), out p, id))
            {
                return;
            }

            int seen = 0;
            foreach (string inner in Inners(p.Body, "S"))
            {
                if (HasBulletPrefix(inner) || HasAutoPrefix(inner))
                {
                    seen++;
                }
            }

            Ok(run, id,
                NoTags(p.Body) && seen >= n && CountNumbered(p) >= n,
                "display 符号 " + seen + "/" + n + " " + Snip(p.Body));
            s.CloseDocByPath(p.Path);
        }

        private static void RunMerge(TestRun run, WdSession s)
        {
            MergeSame(run, s, "an-merge-two", 2, true, true);
            MergeSame(run, s, "an-merge-three", 3, true, true);

            Plant p;
            if (TryPlant(run, s, "同句。\r同句。\r", (ss, path) => ss.ApplyNumberDefault(path, 1, 1),
                out p, "an-merge-mixed-plain"))
            {
                List<string> inners = Inners(p.Body, "S");
                bool first = inners.Count >= 1 && HasDecimalPrefix(inners[0]);
                bool restHasPlain = p.Body.IndexOf("同句。", StringComparison.Ordinal) >= 0;
                Ok(run, "an-merge-mixed-plain", NoTags(p.Body) && first && restHasPlain, Snip(p.Body));
                s.CloseDocByPath(p.Path);
            }

            if (TryPlant(run, s, "同句。\r同句。\r", (ss, path) =>
                ss.ApplyNumberDefault(path, 1, 1) && ss.ApplyBulletDefault(path, 2, 2),
                out p, "an-merge-mixed-types"))
            {
                string joined = string.Join("\n", Inners(p.Body, "S"));
                Ok(run, "an-merge-mixed-types",
                    NoTags(p.Body)
                    && (HasDecimalPrefix(joined) || joined.IndexOf("1.", StringComparison.Ordinal) >= 0)
                    && (HasBulletPrefix(joined) || CountNumbered(p) >= 2),
                    Snip(p.Body));
                s.CloseDocByPath(p.Path);
            }

            if (TryPlant(run, s, "甲。\r甲。\r", (ss, path) =>
                {
                    if (!ss.InsertTableBeforePara(path, 2, "隔"))
                    {
                        return false;
                    }

                    List<string> stories = ss.ParaStories(path);
                    int last = stories == null ? 0 : stories.Count;
                    return ss.ApplyNumberDefault(path, 1, 1)
                        && last >= 2
                        && ss.ApplyNumberDefault(path, last, last);
                },
                out p, "an-merge-no-cross"))
            {
                List<string> codes = Codes(p.Body, "S");
                bool joinedAcrossTable = false;
                foreach (string inner in Inners(p.Body, "S"))
                {
                    if (inner.IndexOf("甲", StringComparison.Ordinal) >= 0
                        && (inner.IndexOf("隔", StringComparison.Ordinal) >= 0
                            || inner.IndexOf("<table", StringComparison.OrdinalIgnoreCase) >= 0))
                    {
                        joinedAcrossTable = true;
                    }
                }

                Ok(run, "an-merge-no-cross",
                    NoTags(p.Body)
                    && !string.IsNullOrEmpty(TableId(p.Body))
                    && codes.Count >= 2
                    && !joinedAcrossTable,
                    "跨表 seg 码数=" + codes.Count + " " + Snip(p.Body));
                s.CloseDocByPath(p.Path);
            }

            if (TryPlant(run, s, "第一条。\r第一条。\r", (ss, path) => ss.ApplyNumberDefault(path, 1, 2),
                out p, "an-merge-reread"))
            {
                string first = p.Body;
                ToolResult again = s.GetContent(p.Id);
                string second = again.Success ? WdSession.Str(again, "document_content") : "";
                Ok(run, "an-merge-reread",
                    again.Success && NoTags(second)
                    && Inners(first, "S").Count == Inners(second, "S").Count
                    && HasDecimalPrefix(string.Join("\n", Inners(second, "S"))),
                    Snip(second));
                s.CloseDocByPath(p.Path);
            }
        }

        private static void MergeSame(TestRun run, WdSession s, string id, int n, bool number, bool expectPrefix)
        {
            var sb = new System.Text.StringBuilder();
            for (int i = 0; i < n; i++)
            {
                sb.Append("第一条。\r");
            }

            Plant p;
            if (!TryPlant(run, s, sb.ToString(), number ? (ss, path) => ss.ApplyNumberDefault(path, 1, n) : (Func<WdSession, string, bool>)null,
                out p, id))
            {
                return;
            }

            string innerAll = string.Join("\n", Inners(p.Body, "S"));
            bool hasBody = p.Body.IndexOf("第一条。", StringComparison.Ordinal) >= 0
                || p.Body.IndexOf("第一条．", StringComparison.Ordinal) >= 0;
            Ok(run, id,
                NoTags(p.Body)
                && (!expectPrefix || HasDecimalPrefix(innerAll) || innerAll.IndexOf("1.", StringComparison.Ordinal) >= 0)
                && hasBody,
                Snip(p.Body));
            s.CloseDocByPath(p.Path);
        }

        private static void RunRepCopy(TestRun run, WdSession s)
        {
            RepNum(run, s, "an-rep-copy-2", JoinN(2), 2, "1. c\r2. d", WantNum(true, true), "c", "d");
            RepNum(run, s, "an-rep-copy-3", JoinN(3), 3, "1. x\r2. y\r3. z", WantNum(true, true, true), "x", "y", "z");
            RepNum(run, s, "an-rep-one", "a\r", 1, "1. x", WantNum(true), "x");

            Plant p;
            if (TryPlant(run, s, JoinN(2), (ss, path) => ss.ApplyBulletDefault(path, 1, 2),
                out p, "an-rep-bullet-copy"))
            {
                ToolResult wr = s.ProcessActions(p.Id, Replace(JoinNumberedS(p), "• p\r• q"));
                List<ParaSnap> after = ContentParas(s, p.Path);
                Ok(run, "an-rep-bullet-copy",
                    wr.Success && after.Count >= 2 && after[0].Numbered && after[1].Numbered
                    && Strip(after[0].Story) == "p" && Strip(after[1].Story) == "q",
                    wr.Success ? Dump(after) : (wr.Error ?? ""));
                s.CloseDocByPath(p.Path);
            }

            if (TryPlant(run, s, JoinN(2), (ss, path) => ss.ApplyNumberDefault(path, 1, 2),
                out p, "an-rep-para"))
            {
                ToolResult paraGot = s.GetContent(p.Id, "paragraph");
                string paraBody = paraGot.Success ? WdSession.Str(paraGot, "document_content") : "";
                List<string> pcodes = Codes(paraBody, "P");
                if (!paraGot.Success || pcodes.Count == 0)
                {
                    run.CaseFail("an-rep-para", paraGot.Error ?? "无 P_");
                }
                else
                {
                    ToolResult wr = s.ProcessActions(p.Id, ReplacePara(pcodes[0], "1. x"));
                    List<ParaSnap> after = ContentParas(s, p.Path);
                    Ok(run, "an-rep-para",
                        wr.Success && after.Count >= 1 && after[0].Numbered && Strip(after[0].Story) == "x",
                        wr.Success ? Dump(after) : (wr.Error ?? ""));
                }

                s.CloseDocByPath(p.Path);
            }

            RepNth(run, s, "an-rep-second-only", 3, 1, "1. m", true, "m");
            RepNth(run, s, "an-rep-last-only", 3, 2, "1. z", true, "z");
            RepNth(run, s, "an-rep-middle-3", 3, 1, "1. mid", true, "mid");

            if (TryPlantTable(run, s, out p, "an-rep-table"))
            {
                string tid = TableId(p.Body);
                string cellCode = CodeOfInner(p.Body, "格甲");
                if (string.IsNullOrEmpty(tid) || string.IsNullOrEmpty(cellCode))
                {
                    run.CaseFail("an-rep-table", "无 T_/格甲 " + Snip(p.Body));
                }
                else
                {
                    ToolResult wr = s.ProcessActions(p.Id, ReplaceInTable(cellCode, tid, "1. g"));
                    ToolResult got = s.GetContent(p.Id);
                    string body = got.Success ? WdSession.Str(got, "document_content") : "";
                    Ok(run, "an-rep-table",
                        wr.Success && body.IndexOf("g", StringComparison.Ordinal) >= 0
                        && HasDecimalPrefix(StripInTable(FirstInner(body, "S"))),
                        wr.Success ? Snip(body) : (wr.Error ?? ""));
                }

                s.CloseDocByPath(p.Path);
            }

            if (TryPlant(run, s, "a\rb\rhello\r", (ss, path) => ss.ApplyNumberDefault(path, 1, 2),
                out p, "an-rep-trailing-cr"))
            {
                ToolResult wr = s.ProcessActions(p.Id, Replace(JoinNumberedS(p), "1. c\r2. d\r"));
                List<ParaSnap> after = ContentParas(s, p.Path);
                ParaSnap hello = FindStory(after, "hello");
                Ok(run, "an-rep-trailing-cr",
                    wr.Success && after.Count >= 2 && after[0].Numbered && after[1].Numbered
                    && (hello.Story == null || !hello.Numbered),
                    wr.Success ? Dump(after) : (wr.Error ?? ""));
                s.CloseDocByPath(p.Path);
            }

            if (TryPlant(run, s, "甲！乙！\r", (ss, path) => ss.ApplyNumberDefault(path, 1, 1),
                out p, "an-rep-two-sent-para"))
            {
                List<string> codes = Codes(p.Body, "S");
                ToolResult wr = s.ProcessActions(p.Id, Replace(codes.Count > 0 ? codes[0] : "", "1. 丙。"));
                ToolResult got = s.GetContent(p.Id);
                string body = got.Success ? WdSession.Str(got, "document_content") : "";
                bool yiPlain = body.IndexOf("乙！", StringComparison.Ordinal) >= 0;
                Ok(run, "an-rep-two-sent-para",
                    wr.Success && CountNumberedPath(s, p.Path) >= 1 && yiPlain,
                    wr.Success ? Snip(body) : (wr.Error ?? ""));
                s.CloseDocByPath(p.Path);
            }

            if (TryPlant(run, s, "第一条。\r第一条。\r", (ss, path) => ss.ApplyNumberDefault(path, 1, 2),
                out p, "an-rep-merge-code"))
            {
                ToolResult wr = s.ProcessActions(p.Id, Replace(JoinNumberedS(p), "1. 甲。\r2. 乙。"));
                List<ParaSnap> after = ContentParas(s, p.Path);
                Ok(run, "an-rep-merge-code",
                    wr.Success && after.Count >= 2 && after[0].Numbered && after[1].Numbered
                    && Strip(after[0].Story).IndexOf("甲", StringComparison.Ordinal) >= 0
                    && Strip(after[1].Story).IndexOf("乙", StringComparison.Ordinal) >= 0,
                    wr.Success ? Dump(after) : (wr.Error ?? ""));
                s.CloseDocByPath(p.Path);
            }
        }

        private static void RunRepCount(TestRun run, WdSession s)
        {
            RepNum(run, s, "an-rep-2to3-allnum", JoinN(2), 2, "1. c\r2. d\r3. e", WantNum(true, true, true), "c", "d", "e");
            RepNum(run, s, "an-rep-2to4-allnum", JoinN(2), 2, "1. a\r2. b\r3. c\r4. d", WantNum(true, true, true, true), "a", "b", "c", "d");
            RepNum(run, s, "an-rep-3to2-allnum", JoinN(3), 3, "1. c\r2. d", WantNum(true, true), "c", "d");
            RepNum(run, s, "an-rep-2to1", JoinN(2), 2, "1. only", WantNum(true), "only");
            RepNum(run, s, "an-rep-1to3", "a\r", 1, "1. a\r2. b\r3. c", WantNum(true, true, true), "a", "b", "c");
            RepNum(run, s, "an-rep-2to3-third-plain", JoinN(2), 2, "1. c\r2. d\re", WantNum(true, true, false), "c", "d", "e");
            RepNum(run, s, "an-rep-2to3-first-plain", JoinN(2), 2, "c\r2. d\r3. e", WantNum(false, true, true), "c", "d", "e");
            RepNum(run, s, "an-rep-2to3-mid-plain", JoinN(2), 2, "1. c\rd\r3. e", WantNum(true, false, true), "c", "d", "e");
        }

        private static void RunRepStory(TestRun run, WdSession s)
        {
            RepNum(run, s, "an-rep-story-2", JoinN(2), 2, "c\rd", WantNum(false, false), "c", "d");
            RepNum(run, s, "an-rep-story-1", "a\r", 1, "x", WantNum(false), "x");
            RepNum(run, s, "an-rep-story-3", JoinN(3), 3, "x\ry\rz", WantNum(false, false, false), "x", "y", "z");
            RepNum(run, s, "an-rep-mixed-prefix", JoinN(2), 2, "1. c\rd", WantNum(true, false), "c", "d");

            Plant p;
            if (TryPlant(run, s, JoinN(2), (ss, path) => ss.ApplyNumberDefault(path, 1, 2),
                out p, "an-rep-other-bullet"))
            {
                ToolResult wr = s.ProcessActions(p.Id, Replace(JoinNumberedS(p), "• c\r2. d"));
                List<ParaSnap> after = ContentParas(s, p.Path);
                Ok(run, "an-rep-other-bullet",
                    wr.Success && after.Count >= 2 && !after[0].Numbered && after[1].Numbered
                    && Strip(after[0].Story).IndexOf("•", StringComparison.Ordinal) >= 0
                    && Strip(after[1].Story) == "d",
                    wr.Success ? Dump(after) : (wr.Error ?? ""));
                s.CloseDocByPath(p.Path);
            }

            if (TryPlant(run, s, JoinN(2), (ss, path) => ss.ApplyNumberDefault(path, 1, 2),
                out p, "an-rep-other-letter"))
            {
                ToolResult wr = s.ProcessActions(p.Id, Replace(JoinNumberedS(p), "a) c\r2. d"));
                List<ParaSnap> after = ContentParas(s, p.Path);
                Ok(run, "an-rep-other-letter",
                    wr.Success && after.Count >= 2 && !after[0].Numbered && after[1].Numbered
                    && Strip(after[0].Story).IndexOf("a)", StringComparison.Ordinal) >= 0
                    && Strip(after[1].Story) == "d",
                    wr.Success ? Dump(after) : (wr.Error ?? ""));
                s.CloseDocByPath(p.Path);
            }

            if (TryPlant(run, s, "a\rb\r", null, out p, "an-rep-plain-orig"))
            {
                ToolResult wr = s.ProcessActions(p.Id, Replace(JoinCodes(p, "S"), "1. c\r2. d"));
                List<ParaSnap> after = ContentParas(s, p.Path);
                Ok(run, "an-rep-plain-orig",
                    wr.Success && after.Count >= 2 && !after[0].Numbered && !after[1].Numbered
                    && Strip(after[0].Story).IndexOf("1. c", StringComparison.Ordinal) >= 0,
                    wr.Success ? Dump(after) : (wr.Error ?? ""));
                s.CloseDocByPath(p.Path);
            }

            if (TryPlant(run, s, "1. 不是列表\r", null, out p, "an-rep-hand-keep"))
            {
                ToolResult wr = s.ProcessActions(p.Id, Replace(JoinCodes(p, "S"), "1. 改了"));
                List<ParaSnap> after = ContentParas(s, p.Path);
                Ok(run, "an-rep-hand-keep",
                    wr.Success && after.Count >= 1 && !after[0].Numbered
                    && Strip(after[0].Story).IndexOf("1. 改了", StringComparison.Ordinal) >= 0,
                    wr.Success ? Dump(after) : (wr.Error ?? ""));
                s.CloseDocByPath(p.Path);
            }

            if (TryPlant(run, s, JoinN(2), (ss, path) => ss.ApplyBulletDefault(path, 1, 2),
                out p, "an-rep-bullet-story"))
            {
                ToolResult wr = s.ProcessActions(p.Id, Replace(JoinNumberedS(p), "p\rq"));
                List<ParaSnap> after = ContentParas(s, p.Path);
                Ok(run, "an-rep-bullet-story",
                    wr.Success && after.Count >= 2 && !after[0].Numbered && !after[1].Numbered,
                    wr.Success ? Dump(after) : (wr.Error ?? ""));
                s.CloseDocByPath(p.Path);
            }

            if (TryPlant(run, s, JoinN(2), (ss, path) => ss.ApplyBulletDefault(path, 1, 2),
                out p, "an-rep-type-keep-bullet"))
            {
                ToolResult wr = s.ProcessActions(p.Id, Replace(JoinNumberedS(p), "• p\rq"));
                List<ParaSnap> after = ContentParas(s, p.Path);
                Ok(run, "an-rep-type-keep-bullet",
                    wr.Success && after.Count >= 2 && after[0].Numbered && !after[1].Numbered
                    && Strip(after[1].Story) == "q",
                    wr.Success ? Dump(after) : (wr.Error ?? ""));
                s.CloseDocByPath(p.Path);
            }
        }

        private static void RunInsert(TestRun run, WdSession s)
        {
            InsAfter(run, s, "an-ins-plain-after-1", JoinN(2), 2, 0, "hello", false, "hello");
            InsAfter(run, s, "an-ins-plain-after-last", JoinN(2), 2, 1, "hello", false, "hello");
            InsAfter(run, s, "an-ins-copy", JoinN(2), 2, 0, "2. hello", true, "hello");
            InsAfter(run, s, "an-ins-two-num", JoinN(2), 2, 0, "1. x\r2. y", null, "x", "y");
            InsAfter(run, s, "an-ins-three-num", JoinN(2), 2, 0, "1. x\r2. y\r3. z", null, "x", "y", "z");
            InsAfter(run, s, "an-ins-mixed", JoinN(2), 2, 0, "1. x\ry", null, "x", "y");
            InsAfter(run, s, "an-ins-three-plain", JoinN(2), 2, 0, "x\ry\rz", null, "x", "y", "z");
            InsAfter(run, s, "an-ins-other-type", JoinN(2), 2, 0, "• z", false, "• z");

            Plant p;
            if (TryPlant(run, s, JoinN(2), (ss, path) => ss.ApplyNumberDefault(path, 1, 2),
                out p, "an-ins-start"))
            {
                ToolResult wr = s.ProcessActions(p.Id, InsertStart("head"));
                ParaSnap head = FindStory(ContentParas(s, p.Path), "head");
                Ok(run, "an-ins-start",
                    wr.Success && head.Story != null && head.Numbered && Strip(head.Story) == "heada",
                    wr.Success ? Dump(ContentParas(s, p.Path)) : (wr.Error ?? ""));
                s.CloseDocByPath(p.Path);
            }

            if (TryPlant(run, s, "a\rb\rhello\r", (ss, path) => ss.ApplyNumberDefault(path, 1, 2),
                out p, "an-ins-after-plain"))
            {
                string helloCode = CodeOfInner(p.Body, "hello");
                ToolResult wr = s.ProcessActions(p.Id, InsertAfter(helloCode, "tail"));
                List<ParaSnap> after = ContentParas(s, p.Path);
                ParaSnap hello = FindStory(after, "hello");
                ParaSnap tail = FindStory(after, "tail");
                Ok(run, "an-ins-after-plain",
                    wr.Success && hello.Story != null && !hello.Numbered && tail.Story != null && !tail.Numbered,
                    wr.Success ? Dump(after) : (wr.Error ?? ""));
                s.CloseDocByPath(p.Path);
            }

            InsAfter(run, s, "an-ins-between-2-3", JoinN(3), 3, 1, "1. n", true, "n");

            if (TryPlantTable(run, s, out p, "an-ins-table"))
            {
                string tid = TableId(p.Body);
                string cellCode = CodeOfInner(p.Body, "格甲");
                if (string.IsNullOrEmpty(tid) || string.IsNullOrEmpty(cellCode))
                {
                    run.CaseFail("an-ins-table", "无 T_/格甲 " + Snip(p.Body));
                }
                else
                {
                    var detail = new Dictionary<string, object>
                    {
                        ["target"] = new Dictionary<string, object>
                        {
                            ["codes"] = cellCode,
                            ["in_table"] = tid
                        },
                        ["insert"] = "1. in",
                        ["format"] = DefaultFormat()
                    };
                    ToolResult wr = s.ProcessActions(p.Id, One("insert", detail));
                    ToolResult got = s.GetContent(p.Id);
                    string body = got.Success ? WdSession.Str(got, "document_content") : "";
                    Ok(run, "an-ins-table",
                        wr.Success && body.IndexOf("in", StringComparison.Ordinal) >= 0,
                        wr.Success ? Snip(body) : (wr.Error ?? ""));
                }

                s.CloseDocByPath(p.Path);
            }
        }

        private static void RunD39(TestRun run, WdSession s)
        {
            Plant p;
            if (TryPlant(run, s, "a\rb\rhello\rtail\r", (ss, path) => ss.ApplyNumberDefault(path, 1, 2),
                out p, "an-plain-stay-rep", "an-plain-stay-ins", "an-plain-stay-story"))
            {
                ToolResult wr = s.ProcessActions(p.Id, Replace(JoinNumberedS(p), "1. x\r2. y\r3. z"));
                List<ParaSnap> after = ContentParas(s, p.Path);
                Ok(run, "an-plain-stay-rep",
                    wr.Success && PlainStay(after, "hello", "tail") && FindStory(after, "z").Numbered,
                    wr.Success ? Dump(after) : (wr.Error ?? ""));

                s.CloseDocByPath(p.Path);
            }

            if (TryPlant(run, s, "a\rb\rhello\rtail\r", (ss, path) => ss.ApplyNumberDefault(path, 1, 2),
                out p, "an-plain-stay-ins"))
            {
                ToolResult wr = s.ProcessActions(p.Id, InsertAfter(Codes(p.Body, "S")[0], "1. n"));
                List<ParaSnap> after = ContentParas(s, p.Path);
                Ok(run, "an-plain-stay-ins",
                    wr.Success && PlainStay(after, "hello", "tail"),
                    wr.Success ? Dump(after) : (wr.Error ?? ""));
                s.CloseDocByPath(p.Path);
            }

            if (TryPlant(run, s, "a\rb\rhello\rtail\r", (ss, path) => ss.ApplyNumberDefault(path, 1, 2),
                out p, "an-plain-stay-story"))
            {
                ToolResult wr = s.ProcessActions(p.Id, Replace(JoinNumberedS(p), "x\ry"));
                List<ParaSnap> after = ContentParas(s, p.Path);
                Ok(run, "an-plain-stay-story",
                    wr.Success && PlainStay(after, "hello", "tail"),
                    wr.Success ? Dump(after) : (wr.Error ?? ""));
                s.CloseDocByPath(p.Path);
            }

            if (TryPlant(run, s, "lead\ra\rb\r", (ss, path) => ss.ApplyNumberDefault(path, 2, 3),
                out p, "an-plain-before"))
            {
                ToolResult wr = s.ProcessActions(p.Id, InsertAfter(Codes(p.Body, "S")[1], "1. n"));
                ParaSnap lead = FindStory(ContentParas(s, p.Path), "lead");
                Ok(run, "an-plain-before",
                    wr.Success && lead.Story != null && !lead.Numbered,
                    wr.Success ? Dump(ContentParas(s, p.Path)) : (wr.Error ?? ""));
                s.CloseDocByPath(p.Path);
            }

            if (TryPlant(run, s, JoinN(2), (ss, path) => ss.ApplyNumberDefault(path, 1, 2),
                out p, "an-shift-ok-ins"))
            {
                ToolResult wr = s.ProcessActions(p.Id, InsertAfter(Codes(p.Body, "S")[0], "2. n"));
                ParaSnap b = FindStory(ContentParas(s, p.Path), "b");
                Ok(run, "an-shift-ok-ins",
                    wr.Success && b.Story != null && Strip(b.Story) == "nb" && b.Numbered,
                    wr.Success ? Dump(ContentParas(s, p.Path)) : (wr.Error ?? ""));
                s.CloseDocByPath(p.Path);
            }

            if (TryPlant(run, s, "a\rb\rc\r", (ss, path) => ss.ApplyNumberDefault(path, 1, 3),
                out p, "an-shift-ok-rep"))
            {
                List<string> codes = Codes(p.Body, "S");
                string firstTwo = codes.Count >= 2 ? codes[0] + "," + codes[1] : JoinCodes(p, "S");
                ToolResult wr = s.ProcessActions(p.Id, Replace(firstTwo, "1. x\r2. y\r3. z"));
                ParaSnap c = FindStory(ContentParas(s, p.Path), "c");
                Ok(run, "an-shift-ok-rep",
                    wr.Success && c.Story != null && Strip(c.Story) == "c" && c.Numbered,
                    wr.Success ? Dump(ContentParas(s, p.Path)) : (wr.Error ?? ""));
                s.CloseDocByPath(p.Path);
            }

            if (TryPlant(run, s, JoinN(3), (ss, path) => ss.ApplyNumberDefault(path, 1, 3),
                out p, "an-shift-ok-clear-mid"))
            {
                List<string> codes = Codes(p.Body, "S");
                string mid = codes.Count >= 2 ? codes[1] : JoinCodes(p, "S");
                ToolResult wr = s.ProcessActions(p.Id, Replace(mid, "m"));
                List<ParaSnap> after = ContentParas(s, p.Path);
                ParaSnap m = FindStory(after, "m");
                Ok(run, "an-shift-ok-clear-mid",
                    wr.Success && m.Story != null && !m.Numbered
                    && after[0].Numbered && FindStory(after, "c").Numbered,
                    wr.Success ? Dump(after) : (wr.Error ?? ""));
                s.CloseDocByPath(p.Path);
            }

            if (TryPlant(run, s, "a\rb\rhello\r", (ss, path) => ss.ApplyNumberDefault(path, 1, 2),
                out p, "an-plain-stay-2to1"))
            {
                ToolResult wr = s.ProcessActions(p.Id, Replace(JoinNumberedS(p), "1. only"));
                ParaSnap hello = FindStory(ContentParas(s, p.Path), "hello");
                Ok(run, "an-plain-stay-2to1",
                    wr.Success && hello.Story != null && !hello.Numbered,
                    wr.Success ? Dump(ContentParas(s, p.Path)) : (wr.Error ?? ""));
                s.CloseDocByPath(p.Path);
            }
        }

        private static void RunReread(TestRun run, WdSession s)
        {
            Plant p;
            if (TryPlant(run, s, JoinN(2), (ss, path) => ss.ApplyNumberDefault(path, 1, 2),
                out p, "an-reread-copy", "an-rep-then-ins"))
            {
                ToolResult wr = s.ProcessActions(p.Id, Replace(JoinNumberedS(p), "1. c\r2. d"));
                ToolResult got = s.GetContent(p.Id);
                string body = got.Success ? WdSession.Str(got, "document_content") : "";
                Ok(run, "an-reread-copy",
                    wr.Success && got.Success && NoTags(body)
                    && HasDecimalPrefix(string.Join("\n", Inners(body, "S"))),
                    wr.Success ? Snip(body) : (wr.Error ?? ""));

                List<string> codes = Codes(body, "S");
                ToolResult ins = s.ProcessActions(p.Id, InsertAfter(codes.Count > 0 ? codes[0] : "", "hello"));
                ParaSnap hello = FindStory(ContentParas(s, p.Path), "hello");
                ParaSnap c = FindStory(ContentParas(s, p.Path), "c");
                Ok(run, "an-rep-then-ins",
                    ins.Success && hello.Story != null && hello.Numbered
                    && Strip(hello.Story) == "hellod" && c.Numbered,
                    ins.Success ? Dump(ContentParas(s, p.Path)) : (ins.Error ?? ""));
                s.CloseDocByPath(p.Path);
            }

            if (TryPlant(run, s, JoinN(2), (ss, path) => ss.ApplyNumberDefault(path, 1, 2),
                out p, "an-reread-story"))
            {
                ToolResult wr = s.ProcessActions(p.Id, Replace(JoinNumberedS(p), "c\rd"));
                ToolResult got = s.GetContent(p.Id);
                string body = got.Success ? WdSession.Str(got, "document_content") : "";
                bool noPrefix = true;
                foreach (string inner in Inners(body, "S"))
                {
                    if (HasDecimalPrefix(inner))
                    {
                        noPrefix = false;
                    }
                }

                Ok(run, "an-reread-story",
                    wr.Success && got.Success && NoTags(body) && noPrefix,
                    wr.Success ? Snip(body) : (wr.Error ?? ""));
                s.CloseDocByPath(p.Path);
            }

            if (TryPlant(run, s, JoinN(2), (ss, path) => ss.ApplyNumberDefault(path, 1, 2),
                out p, "an-reread-ins"))
            {
                ToolResult wr = s.ProcessActions(p.Id, InsertAfter(Codes(p.Body, "S")[0], "hello"));
                ToolResult got = s.GetContent(p.Id);
                string body = got.Success ? WdSession.Str(got, "document_content") : "";
                string inner = InnerContaining(body, "hello");
                ParaSnap glued = FindStory(ContentParas(s, p.Path), "hello");
                Ok(run, "an-reread-ins",
                    wr.Success && got.Success && NoTags(body)
                    && HasDecimalPrefix(inner)
                    && inner.IndexOf("hellob", StringComparison.Ordinal) >= 0
                    && glued.Numbered && Strip(glued.Story) == "hellob",
                    wr.Success ? Snip(body) : (wr.Error ?? ""));
                s.CloseDocByPath(p.Path);
            }
        }

        private static void RunMore(TestRun run, WdSession s)
        {
            ReadDecimal(run, s, "an-read-decimal-1", 1);
            ReadDecimal(run, s, "an-read-decimal-4", 4);
            ReadDecimal(run, s, "an-read-decimal-6", 6);
            ReadBullet(run, s, "an-read-bullet-1", 1);
            ReadBullet(run, s, "an-read-bullet-4", 4);

            Plant p;
            if (TryPlant(run, s, "plain\ronly\r", null, out p, "an-read-plain-only"))
            {
                Ok(run, "an-read-plain-only",
                    NoTags(p.Body) && !AnyNumbered(p) && !HasDecimalPrefix(FirstInner(p.Body, "S")),
                    Snip(p.Body));
                s.CloseDocByPath(p.Path);
            }

            if (TryPlant(run, s, "lead\rx\ry\r", (ss, path) => ss.ApplyNumberDefault(path, 2, 3),
                out p, "an-read-lead-plain"))
            {
                Ok(run, "an-read-lead-plain",
                    NoTags(p.Body)
                    && !HasDecimalPrefix(InnerOfStory(p, "lead"))
                    && HasDecimalPrefix(InnerOfStory(p, "x")),
                    Snip(p.Body));
                s.CloseDocByPath(p.Path);
            }

            if (TryPlant(run, s, "第一条。\r第二条。\r", (ss, path) => ss.ApplyNumberDefault(path, 1, 2),
                out p, "an-read-cn-body"))
            {
                Ok(run, "an-read-cn-body",
                    NoTags(p.Body)
                    && HasDecimalPrefix(InnerOfStory(p, "第一条"))
                    && StoryHas(p, "第一条")
                    && !HasDecimalPrefix(ParaOfStory(p, "第一条").Story ?? ""),
                    Snip(p.Body));
                s.CloseDocByPath(p.Path);
            }

            if (TryPlant(run, s, "a\rb\rhello\r", (ss, path) => ss.ApplyNumberDefault(path, 1, 2),
                out p, "an-read-two-num-plain"))
            {
                Ok(run, "an-read-two-num-plain",
                    NoTags(p.Body)
                    && HasDecimalPrefix(InnerOfStory(p, "a"))
                    && !HasDecimalPrefix(InnerOfStory(p, "hello"))
                    && !ParaOfStory(p, "hello").Numbered,
                    Snip(p.Body));
                s.CloseDocByPath(p.Path);
            }

            if (TryPlant(run, s, JoinN(2), (ss, path) => ss.ApplyBulletDefault(path, 1, 2),
                out p, "an-read-para-bullet"))
            {
                ToolResult paraGot = s.GetContent(p.Id, "paragraph");
                string paraBody = paraGot.Success ? WdSession.Str(paraGot, "document_content") : "";
                Ok(run, "an-read-para-bullet",
                    paraGot.Success && NoTags(paraBody)
                    && Codes(paraBody, "P").Count >= 2
                    && (HasBulletPrefix(FirstInner(paraBody, "P")) || HasAutoPrefix(FirstInner(paraBody, "P"))),
                    Snip(paraBody));
                s.CloseDocByPath(p.Path);
            }

            MergeSame(run, s, "an-merge-four", 4, true, true);

            if (TryPlant(run, s, "同句。\r同句。\r", (ss, path) => ss.ApplyBulletDefault(path, 1, 2),
                out p, "an-merge-bullet-two"))
            {
                string joined = string.Join("\n", Inners(p.Body, "S"));
                Ok(run, "an-merge-bullet-two",
                    NoTags(p.Body) && CountNumbered(p) >= 2
                    && (HasBulletPrefix(joined) || HasAutoPrefix(joined)),
                    Snip(p.Body));
                s.CloseDocByPath(p.Path);
            }

            if (TryPlant(run, s, "同句。\r同句。\r", null, out p, "an-merge-plain-two"))
            {
                Ok(run, "an-merge-plain-two",
                    NoTags(p.Body) && !AnyNumbered(p)
                    && p.Body.IndexOf("同句。", StringComparison.Ordinal) >= 0,
                    Snip(p.Body));
                s.CloseDocByPath(p.Path);
            }

            if (TryPlant(run, s, "第一条。\r第一条。\r", (ss, path) => ss.ApplyNumberDefault(path, 1, 2),
                out p, "an-merge-two-para"))
            {
                ToolResult paraGot = s.GetContent(p.Id, "paragraph");
                string paraBody = paraGot.Success ? WdSession.Str(paraGot, "document_content") : "";
                Ok(run, "an-merge-two-para",
                    paraGot.Success && NoTags(paraBody)
                    && Codes(paraBody, "P").Count >= 2
                    && HasDecimalPrefix(string.Join("\n", Inners(paraBody, "P"))),
                    Snip(paraBody));
                s.CloseDocByPath(p.Path);
            }

            RepNum(run, s, "an-rep-copy-4", JoinN(4), 4, "1. w\r2. x\r3. y\r4. z",
                WantNum(true, true, true, true), "w", "x", "y", "z");
            RepNum(run, s, "an-rep-copy-5", JoinN(5), 5, "1. a\r2. b\r3. c\r4. d\r5. e",
                WantNum(true, true, true, true, true), "a", "b", "c", "d", "e");
            RepNth(run, s, "an-rep-first-only", 3, 0, "1. first", true, "first");
            RepNum(run, s, "an-rep-copy-cn", "甲。\r乙。\r", 2, "1. 丙。\r2. 丁。",
                WantNum(true, true), "丙", "丁");

            if (TryPlant(run, s, "a\r", (ss, path) => ss.ApplyBulletDefault(path, 1, 1),
                out p, "an-rep-bullet-one"))
            {
                ToolResult wr = s.ProcessActions(p.Id, Replace(JoinNumberedS(p), "• p"));
                List<ParaSnap> after = ContentParas(s, p.Path);
                Ok(run, "an-rep-bullet-one",
                    wr.Success && after.Count >= 1 && after[0].Numbered && Strip(after[0].Story) == "p",
                    wr.Success ? Dump(after) : (wr.Error ?? ""));
                s.CloseDocByPath(p.Path);
            }

            if (TryPlant(run, s, "a\rb\rc\rhello\r", (ss, path) => ss.ApplyNumberDefault(path, 1, 3),
                out p, "an-rep-trailing-cr-3"))
            {
                ToolResult wr = s.ProcessActions(p.Id, Replace(JoinNumberedS(p), "1. x\r2. y\r3. z\r"));
                List<ParaSnap> after = ContentParas(s, p.Path);
                Ok(run, "an-rep-trailing-cr-3",
                    wr.Success && after[0].Numbered && after[1].Numbered && after[2].Numbered
                    && PlainStay(after, "hello"),
                    wr.Success ? Dump(after) : (wr.Error ?? ""));
                s.CloseDocByPath(p.Path);
            }

            if (TryPlant(run, s, "a\rb\rkeep\r", (ss, path) => ss.ApplyNumberDefault(path, 1, 2),
                out p, "an-rep-keep-tail"))
            {
                ToolResult wr = s.ProcessActions(p.Id, Replace(JoinNumberedS(p), "1. c\r2. d"));
                List<ParaSnap> after = ContentParas(s, p.Path);
                Ok(run, "an-rep-keep-tail",
                    wr.Success && after[0].Numbered && after[1].Numbered
                    && Strip(after[0].Story) == "c" && Strip(after[1].Story) == "d"
                    && PlainStay(after, "keep"),
                    wr.Success ? Dump(after) : (wr.Error ?? ""));
                s.CloseDocByPath(p.Path);
            }

            if (TryPlant(run, s, "a\rb\rc\rkeep\r", (ss, path) => ss.ApplyNumberDefault(path, 1, 3),
                out p, "an-rep-keep-tail-3to2"))
            {
                ToolResult wr = s.ProcessActions(p.Id, Replace(JoinNumberedS(p), "1. x\r2. y"));
                List<ParaSnap> after = ContentParas(s, p.Path);
                Ok(run, "an-rep-keep-tail-3to2",
                    wr.Success && after[0].Numbered && after[1].Numbered
                    && Strip(after[0].Story) == "x" && Strip(after[1].Story) == "y"
                    && PlainStay(after, "keep"),
                    wr.Success ? Dump(after) : (wr.Error ?? ""));
                s.CloseDocByPath(p.Path);
            }

            RepNum(run, s, "an-rep-2to5-allnum", JoinN(2), 2,
                "1. a\r2. b\r3. c\r4. d\r5. e",
                WantNum(true, true, true, true, true), "a", "b", "c", "d", "e");
            RepNum(run, s, "an-rep-4to2-allnum", JoinN(4), 4, "1. p\r2. q",
                WantNum(true, true), "p", "q");
            RepNum(run, s, "an-rep-3to1", JoinN(3), 3, "1. only", WantNum(true), "only");
            RepNum(run, s, "an-rep-1to2", "a\r", 1, "1. m\r2. n", WantNum(true, true), "m", "n");

            if (TryPlant(run, s, JoinN(3), (ss, path) => ss.ApplyNumberDefault(path, 1, 3),
                out p, "an-rep-3to4-last-plain"))
            {
                ToolResult wr = s.ProcessActions(p.Id, Replace(JoinNumberedS(p), "1. u\r2. v\r3. w\rplain"));
                List<ParaSnap> after = ContentParas(s, p.Path);
                Ok(run, "an-rep-3to4-last-plain",
                    wr.Success && after.Count >= 4
                    && after[0].Numbered && after[1].Numbered && after[2].Numbered
                    && !after[3].Numbered && Strip(after[3].Story) == "plain"
                    && Strip(after[0].Story) == "u",
                    wr.Success ? Dump(after) : (wr.Error ?? ""));
                s.CloseDocByPath(p.Path);
            }

            RepNum(run, s, "an-rep-copy-6", JoinN(6), 6,
                "1. a\r2. b\r3. c\r4. d\r5. e\r6. f",
                WantNum(true, true, true, true, true, true), "a", "b", "c", "d", "e", "f");

            RepNum(run, s, "an-rep-story-4", JoinN(4), 4, "w\rx\ry\rz",
                WantNum(false, false, false, false), "w", "x", "y", "z");

            if (TryPlant(run, s, JoinN(2), (ss, path) => ss.ApplyNumberDefault(path, 1, 2),
                out p, "an-rep-other-wrap"))
            {
                ToolResult wr = s.ProcessActions(p.Id, Replace(JoinNumberedS(p), "(1) c\r2. d"));
                List<ParaSnap> after = ContentParas(s, p.Path);
                Ok(run, "an-rep-other-wrap",
                    wr.Success && after.Count >= 2 && !after[0].Numbered && after[1].Numbered
                    && Strip(after[0].Story).IndexOf("(1)", StringComparison.Ordinal) >= 0
                    && Strip(after[1].Story) == "d",
                    wr.Success ? Dump(after) : (wr.Error ?? ""));
                s.CloseDocByPath(p.Path);
            }

            if (TryPlant(run, s, "a\r", null, out p, "an-rep-plain-orig-one"))
            {
                ToolResult wr = s.ProcessActions(p.Id, Replace(JoinCodes(p, "S"), "1. c"));
                List<ParaSnap> after = ContentParas(s, p.Path);
                Ok(run, "an-rep-plain-orig-one",
                    wr.Success && after.Count >= 1 && !after[0].Numbered
                    && Strip(after[0].Story).IndexOf("1. c", StringComparison.Ordinal) >= 0,
                    wr.Success ? Dump(after) : (wr.Error ?? ""));
                s.CloseDocByPath(p.Path);
            }

            if (TryPlant(run, s, "a\r", (ss, path) => ss.ApplyBulletDefault(path, 1, 1),
                out p, "an-rep-bullet-one-story"))
            {
                ToolResult wr = s.ProcessActions(p.Id, Replace(JoinNumberedS(p), "p"));
                List<ParaSnap> after = ContentParas(s, p.Path);
                Ok(run, "an-rep-bullet-one-story",
                    wr.Success && after.Count >= 1 && !after[0].Numbered && Strip(after[0].Story) == "p",
                    wr.Success ? Dump(after) : (wr.Error ?? ""));
                s.CloseDocByPath(p.Path);
            }

            if (TryPlant(run, s, "1. aa\r2. bb\r", null, out p, "an-rep-hand-2-keep"))
            {
                ToolResult wr = s.ProcessActions(p.Id, Replace(JoinCodes(p, "S"), "1. cc\r2. dd"));
                List<ParaSnap> after = ContentParas(s, p.Path);
                Ok(run, "an-rep-hand-2-keep",
                    wr.Success && after.Count >= 2 && !after[0].Numbered && !after[1].Numbered
                    && Strip(after[0].Story).IndexOf("1. cc", StringComparison.Ordinal) >= 0,
                    wr.Success ? Dump(after) : (wr.Error ?? ""));
                s.CloseDocByPath(p.Path);
            }

            if (TryPlant(run, s, JoinN(2), (ss, path) => ss.ApplyNumberDefault(path, 1, 2),
                out p, "an-rep-mixed-tail-num"))
            {
                ToolResult wr = s.ProcessActions(p.Id, Replace(JoinNumberedS(p), "c\rd\r3. e"));
                List<ParaSnap> after = ContentParas(s, p.Path);
                Ok(run, "an-rep-mixed-tail-num",
                    wr.Success && after.Count >= 3
                    && !after[0].Numbered && !after[1].Numbered && after[2].Numbered
                    && Strip(after[2].Story) == "e",
                    wr.Success ? Dump(after) : (wr.Error ?? ""));
                s.CloseDocByPath(p.Path);
            }

            InsAfter(run, s, "an-ins-copy-last", JoinN(2), 2, 1, "2. n", false, "2. n");
            InsAfter(run, s, "an-ins-glue-x", JoinN(2), 2, 0, "x", true, "x");
            InsAfter(run, s, "an-ins-two-plain", JoinN(2), 2, 0, "x\ry", null, "x", "y");
            InsAfter(run, s, "an-ins-hello-cr", JoinN(2), 2, 0, "hello\r", false, "hello");
            InsAfter(run, s, "an-ins-after-1-two-num", JoinN(2), 2, 0, "1. x\r2. y", null, "x", "y");
            InsAfter(run, s, "an-ins-copy-between", JoinN(3), 3, 0, "1. n", true, "n");

            if (TryPlant(run, s, JoinN(2), (ss, path) => ss.ApplyNumberDefault(path, 1, 2),
                out p, "an-ins-start-num"))
            {
                ToolResult wr = s.ProcessActions(p.Id, InsertStart("1. head"));
                ParaSnap head = FindStory(ContentParas(s, p.Path), "head");
                Ok(run, "an-ins-start-num",
                    wr.Success && head.Story != null && head.Numbered
                    && Strip(head.Story).IndexOf("head", StringComparison.Ordinal) >= 0,
                    wr.Success ? Dump(ContentParas(s, p.Path)) : (wr.Error ?? ""));
                s.CloseDocByPath(p.Path);
            }

            if (TryPlant(run, s, "a\rb\rkeep\r", (ss, path) => ss.ApplyNumberDefault(path, 1, 2),
                out p, "an-ins-keep-tail"))
            {
                List<string> codes = Codes(p.Body, "S");
                ToolResult wr = s.ProcessActions(p.Id, InsertAfter(codes[0], "1. n"));
                List<ParaSnap> after = ContentParas(s, p.Path);
                Ok(run, "an-ins-keep-tail",
                    wr.Success && FindStory(after, "n").Numbered && PlainStay(after, "keep"),
                    wr.Success ? Dump(after) : (wr.Error ?? ""));
                s.CloseDocByPath(p.Path);
            }

            if (TryPlant(run, s, "a\rb\rkeep\rtail\r", (ss, path) => ss.ApplyNumberDefault(path, 1, 2),
                out p, "an-plain-stay-rep-3"))
            {
                ToolResult wr = s.ProcessActions(p.Id, Replace(JoinNumberedS(p), "1. x\r2. y\r3. z"));
                List<ParaSnap> after = ContentParas(s, p.Path);
                Ok(run, "an-plain-stay-rep-3",
                    wr.Success && FindStory(after, "z").Numbered && PlainStay(after, "keep", "tail"),
                    wr.Success ? Dump(after) : (wr.Error ?? ""));
                s.CloseDocByPath(p.Path);
            }

            if (TryPlant(run, s, "a\rb\rkeep\r", (ss, path) => ss.ApplyBulletDefault(path, 1, 2),
                out p, "an-plain-stay-bullet"))
            {
                ToolResult wr = s.ProcessActions(p.Id, Replace(JoinNumberedS(p), "• p\r• q"));
                List<ParaSnap> after = ContentParas(s, p.Path);
                Ok(run, "an-plain-stay-bullet",
                    wr.Success && after[0].Numbered && after[1].Numbered && PlainStay(after, "keep"),
                    wr.Success ? Dump(after) : (wr.Error ?? ""));
                s.CloseDocByPath(p.Path);
            }

            if (TryPlant(run, s, "a\rkeep\rb\r", (ss, path) =>
                ss.ApplyNumberDefault(path, 1, 1) && ss.ApplyNumberDefault(path, 3, 3),
                out p, "an-plain-mid-rep"))
            {
                ToolResult wr = s.ProcessActions(p.Id, Replace(CodeOfInner(p.Body, "a"), "1. x"));
                List<ParaSnap> after = ContentParas(s, p.Path);
                Ok(run, "an-plain-mid-rep",
                    wr.Success && FindStory(after, "x").Numbered
                    && FindStory(after, "b").Numbered && PlainStay(after, "keep"),
                    wr.Success ? Dump(after) : (wr.Error ?? ""));
                s.CloseDocByPath(p.Path);
            }

            if (TryPlant(run, s, "a\rb\rkeep\r", (ss, path) => ss.ApplyNumberDefault(path, 1, 2),
                out p, "an-plain-stay-2to3"))
            {
                ToolResult wr = s.ProcessActions(p.Id, Replace(JoinNumberedS(p), "1. x\r2. y\r3. z"));
                List<ParaSnap> after = ContentParas(s, p.Path);
                Ok(run, "an-plain-stay-2to3",
                    wr.Success && FindStory(after, "z").Numbered && PlainStay(after, "keep"),
                    wr.Success ? Dump(after) : (wr.Error ?? ""));
                s.CloseDocByPath(p.Path);
            }

            if (TryPlant(run, s, "a\rb\rc\rd\r", (ss, path) => ss.ApplyNumberDefault(path, 1, 4),
                out p, "an-shift-ok-rep-4"))
            {
                List<string> codes = Codes(p.Body, "S");
                string firstTwo = codes.Count >= 2 ? codes[0] + "," + codes[1] : JoinNumberedS(p);
                ToolResult wr = s.ProcessActions(p.Id, Replace(firstTwo, "1. x\r2. y\r3. z"));
                List<ParaSnap> after = ContentParas(s, p.Path);
                Ok(run, "an-shift-ok-rep-4",
                    wr.Success && FindStory(after, "c").Numbered && FindStory(after, "d").Numbered
                    && Strip(FindStory(after, "c").Story) == "c",
                    wr.Success ? Dump(after) : (wr.Error ?? ""));
                s.CloseDocByPath(p.Path);
            }

            if (TryPlant(run, s, JoinN(3), (ss, path) => ss.ApplyNumberDefault(path, 1, 3),
                out p, "an-reread-copy-3"))
            {
                ToolResult wr = s.ProcessActions(p.Id, Replace(JoinNumberedS(p), "1. x\r2. y\r3. z"));
                ToolResult got = s.GetContent(p.Id);
                string body = got.Success ? WdSession.Str(got, "document_content") : "";
                Ok(run, "an-reread-copy-3",
                    wr.Success && got.Success && NoTags(body)
                    && HasDecimalPrefix(InnerContaining(body, "x"))
                    && HasDecimalPrefix(InnerContaining(body, "z")),
                    wr.Success ? Snip(body) : (wr.Error ?? ""));
                s.CloseDocByPath(p.Path);
            }

            if (TryPlant(run, s, "a\rb\rkeep\r", (ss, path) => ss.ApplyNumberDefault(path, 1, 2),
                out p, "an-reread-keep-tail"))
            {
                ToolResult wr = s.ProcessActions(p.Id, Replace(JoinNumberedS(p), "1. c\r2. d"));
                ToolResult got = s.GetContent(p.Id);
                string body = got.Success ? WdSession.Str(got, "document_content") : "";
                List<ParaSnap> after = ContentParas(s, p.Path);
                Ok(run, "an-reread-keep-tail",
                    wr.Success && got.Success && NoTags(body)
                    && HasDecimalPrefix(InnerContaining(body, "c"))
                    && !HasDecimalPrefix(InnerContaining(body, "keep"))
                    && PlainStay(after, "keep"),
                    wr.Success ? Snip(body) : (wr.Error ?? ""));
                s.CloseDocByPath(p.Path);
            }

            if (TryPlant(run, s, JoinN(2), (ss, path) => ss.ApplyBulletDefault(path, 1, 2),
                out p, "an-reread-bullet"))
            {
                ToolResult wr = s.ProcessActions(p.Id, Replace(JoinNumberedS(p), "• p\r• q"));
                ToolResult got = s.GetContent(p.Id);
                string body = got.Success ? WdSession.Str(got, "document_content") : "";
                Ok(run, "an-reread-bullet",
                    wr.Success && got.Success && NoTags(body)
                    && (HasBulletPrefix(InnerContaining(body, "p")) || HasAutoPrefix(InnerContaining(body, "p"))),
                    wr.Success ? Snip(body) : (wr.Error ?? ""));
                s.CloseDocByPath(p.Path);
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
            if (!WantAny("an-no-quit"))
            {
                return;
            }

            bool wordOk = OfficeAlive.NotQuit("WINWORD", wordBefore, wordWinBefore, out wordDetail);
            bool wpsOk = OfficeAlive.NotQuit("wps", wpsBefore, wpsWinBefore, out wpsDetail);
            Ok(run, "an-no-quit", wordOk && wpsOk, wordDetail + " " + wpsDetail);
        }

        private static bool WantAny(params string[] ids)
        {
            string f = IdFilter;
            if (string.IsNullOrEmpty(f)
                || string.Equals(f, "an", StringComparison.OrdinalIgnoreCase)
                || string.Equals(f, "an-", StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            if (ids == null || ids.Length == 0)
            {
                return true;
            }

            foreach (string id in ids)
            {
                if (string.IsNullOrEmpty(id))
                {
                    continue;
                }

                if (id.IndexOf(f, StringComparison.OrdinalIgnoreCase) >= 0
                    || f.IndexOf(id, StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    return true;
                }
            }

            return false;
        }

        private static void RepNum(
            TestRun run,
            WdSession s,
            string id,
            string plain,
            int numberTo,
            string neu,
            bool[] wantNum,
            params string[] stories)
        {
            Plant p;
            if (!TryPlant(run, s, plain, (ss, path) => ss.ApplyNumberDefault(path, 1, numberTo), out p, id))
            {
                return;
            }

            string codes = JoinNumberedS(p);
            ToolResult wr = s.ProcessActions(p.Id, Replace(codes, neu));
            List<ParaSnap> after = ContentParas(s, p.Path);
            bool ok = wr.Success && MatchWant(after, wantNum, stories);
            Ok(run, id, ok, (wr.Error ?? "") + " codes=" + codes + " exec=" + WdSession.IntOf(wr, "executed_actions") + " " + Dump(after));
            s.CloseDocByPath(p.Path);
        }

        private static void RepNth(TestRun run, WdSession s, string id, int n, int index, string neu, bool wantNum, string story)
        {
            Plant p;
            if (!TryPlant(run, s, JoinN(n), (ss, path) => ss.ApplyNumberDefault(path, 1, n), out p, id))
            {
                return;
            }

            List<string> codes = Codes(p.Body, "S");
            if (codes.Count <= index)
            {
                run.CaseFail(id, "码不足 " + codes.Count);
                s.CloseDocByPath(p.Path);
                return;
            }

            ToolResult wr = s.ProcessActions(p.Id, Replace(codes[index], neu));
            List<ParaSnap> after = ContentParas(s, p.Path);
            ParaSnap hit = FindStory(after, story);
            bool sides = after.Count >= n && after[0].Numbered && after[after.Count - 1].Numbered;
            Ok(run, id,
                wr.Success && hit.Story != null && hit.Numbered == wantNum && sides,
                wr.Success ? Dump(after) : (wr.Error ?? ""));
            s.CloseDocByPath(p.Path);
        }

        private static void InsAfter(
            TestRun run,
            WdSession s,
            string id,
            string plain,
            int numberTo,
            int afterIndex,
            string insert,
            bool? newNumbered,
            params string[] expectStories)
        {
            Plant p;
            if (!TryPlant(run, s, plain, (ss, path) => ss.ApplyNumberDefault(path, 1, numberTo), out p, id))
            {
                return;
            }

            List<string> codes = Codes(p.Body, "S");
            if (codes.Count <= afterIndex)
            {
                run.CaseFail(id, "码不足");
                s.CloseDocByPath(p.Path);
                return;
            }

            ToolResult wr = s.ProcessActions(p.Id, InsertAfter(codes[afterIndex], insert));
            List<ParaSnap> after = ContentParas(s, p.Path);
            bool storiesOk = true;
            foreach (string st in expectStories)
            {
                if (FindStory(after, st.Replace("• ", "").Replace("•", "")).Story == null
                    && FindStory(after, st).Story == null)
                {
                    storiesOk = false;
                }
            }

            bool numOk = true;
            if (newNumbered.HasValue && expectStories.Length >= 1)
            {
                ParaSnap n = FindStory(after, expectStories[0].Replace("• z", "z").Replace("• ", ""));
                if (n.Story == null)
                {
                    n = FindStory(after, expectStories[0]);
                }

                numOk = n.Story != null && n.Numbered == newNumbered.Value;
            }

            if (id == "an-ins-two-num" || id == "an-ins-three-num")
            {
                numOk = FindStory(after, "x").Numbered && FindStory(after, "y").Numbered;
                if (id == "an-ins-three-num")
                {
                    numOk = numOk && FindStory(after, "z").Numbered;
                }
            }
            else if (id == "an-ins-mixed")
            {
                ParaSnap y = FindStory(after, "y");
                numOk = FindStory(after, "x").Numbered
                    && y.Story != null
                    && Strip(y.Story) == "yb";
            }
            else if (id == "an-ins-three-plain")
            {
                ParaSnap z = FindStory(after, "z");
                numOk = !FindStory(after, "x").Numbered
                    && !FindStory(after, "y").Numbered
                    && z.Story != null
                    && Strip(z.Story) == "zb";
            }
            else if (id == "an-ins-plain-after-1")
            {
                ParaSnap glued = FindStory(after, "hello");
                numOk = FindStory(after, "a").Numbered
                    && glued.Story != null
                    && glued.Numbered
                    && Strip(glued.Story) == "hellob";
            }
            else if (id == "an-ins-plain-after-last")
            {
                numOk = !FindStory(after, "hello").Numbered
                    && FindStory(after, "a").Numbered
                    && FindStory(after, "b").Numbered;
            }
            else if (id == "an-ins-other-type")
            {
                ParaSnap z = FindStory(after, "z");
                numOk = z.Story != null
                    && z.Numbered
                    && Strip(z.Story).IndexOf("z", StringComparison.Ordinal) >= 0
                    && Strip(z.Story).IndexOf("b", StringComparison.Ordinal) >= 0;
            }
            else if (id == "an-ins-between-2-3")
            {
                numOk = FindStory(after, "n").Numbered && FindStory(after, "c").Numbered;
            }
            else if (id == "an-ins-glue-x")
            {
                ParaSnap glued = FindStory(after, "x");
                numOk = glued.Story != null
                    && glued.Numbered
                    && Strip(glued.Story) == "xb";
            }
            else if (id == "an-ins-two-plain")
            {
                ParaSnap y = FindStory(after, "y");
                numOk = !FindStory(after, "x").Numbered
                    && y.Story != null
                    && Strip(y.Story) == "yb";
            }
            else if (id == "an-ins-after-1-two-num")
            {
                numOk = FindStory(after, "x").Numbered && FindStory(after, "y").Numbered;
            }

            Ok(run, id, wr.Success && storiesOk && numOk, wr.Success ? Dump(after) : (wr.Error ?? ""));
            s.CloseDocByPath(p.Path);
        }

        private static bool TryPlantTable(TestRun run, WdSession s, out Plant p, params string[] failIds)
        {
            p = default(Plant);
            if (!WantAny(failIds))
            {
                return false;
            }

            if (!TryOpen(run, s, out p.Path, out p.Id, failIds))
            {
                return false;
            }

            if (!s.InsertTable(p.Path, 1, 2, new[] { "格甲", "格乙" }) || !s.NumberFirstTableCell(p.Path))
            {
                FailMany(run, "建表或套编号失败", failIds);
                s.CloseDocByPath(p.Path);
                return false;
            }

            return FillBody(run, s, ref p, failIds);
        }

        private static bool TryPlant(
            TestRun run,
            WdSession s,
            string plain,
            Func<WdSession, string, bool> setup,
            out Plant p,
            params string[] failIds)
        {
            p = default(Plant);
            if (!WantAny(failIds))
            {
                return false;
            }

            if (!TryOpen(run, s, out p.Path, out p.Id, failIds))
            {
                return false;
            }

            if (!s.TypePlain(p.Path, plain))
            {
                FailMany(run, "COM 写句失败", failIds);
                s.CloseDocByPath(p.Path);
                return false;
            }

            s.ActivateDoc(p.Path);
            if (setup != null && !setup(s, p.Path))
            {
                FailMany(run, "COM 套列表失败", failIds);
                s.CloseDocByPath(p.Path);
                return false;
            }

            return FillBody(run, s, ref p, failIds);
        }

        private static bool TryOpen(TestRun run, WdSession s, out string path, out string id, params string[] failIds)
        {
            path = s.NewDocxPath();
            id = null;
            ToolResult opened = s.Open(path, true, s.App);
            if (!opened.Success)
            {
                FailMany(run, opened.Error ?? "新建失败", failIds);
                return false;
            }

            id = WdSession.Str(opened, "channel_id");
            s.ActivateDoc(path);
            return true;
        }

        private static bool FillBody(TestRun run, WdSession s, ref Plant p, params string[] failIds)
        {
            s.ActivateDoc(p.Path);
            ToolResult got = s.GetContent(p.Id);
            if (!got.Success)
            {
                FailMany(run, got.Error ?? "读失败", failIds);
                s.CloseDocByPath(p.Path);
                return false;
            }

            p.Body = WdSession.Str(got, "document_content");
            p.Paras = ContentParas(s, p.Path);
            return true;
        }

        private struct Plant
        {
            public string Path;
            public string Id;
            public string Body;
            public List<ParaSnap> Paras;
        }

        private struct ParaSnap
        {
            public string Story;
            public string List;
            public bool Numbered;
        }

        private static List<ParaSnap> ContentParas(WdSession s, string path)
        {
            List<string> stories = s.ParaStories(path);
            List<string> lists = s.ParaListStrings(path);
            var result = new List<ParaSnap>();
            for (int i = 0; i < stories.Count; i++)
            {
                string st = Strip(stories[i]);
                string ls = i < lists.Count ? (lists[i] ?? "") : "";
                if (i == stories.Count - 1 && st.Length == 0 && result.Count > 0)
                {
                    continue;
                }

                result.Add(new ParaSnap
                {
                    Story = st,
                    List = ls,
                    Numbered = !string.IsNullOrEmpty(ls.Trim())
                });
            }

            return result;
        }

        private static bool AnyNumbered(Plant p)
        {
            return CountNumbered(p) > 0;
        }

        private static int CountNumbered(Plant p)
        {
            int n = 0;
            if (p.Paras == null)
            {
                return 0;
            }

            foreach (ParaSnap x in p.Paras)
            {
                if (x.Numbered)
                {
                    n++;
                }
            }

            return n;
        }

        private static int CountNumberedPath(WdSession s, string path)
        {
            int n = 0;
            foreach (ParaSnap x in ContentParas(s, path))
            {
                if (x.Numbered)
                {
                    n++;
                }
            }

            return n;
        }

        private static bool StoryHas(Plant p, string needle)
        {
            if (p.Paras == null)
            {
                return false;
            }

            foreach (ParaSnap x in p.Paras)
            {
                if ((x.Story ?? "").IndexOf(needle, StringComparison.Ordinal) >= 0)
                {
                    return true;
                }
            }

            return false;
        }

        private static ParaSnap ParaOfStory(Plant p, string needle)
        {
            return FindStory(p.Paras, needle);
        }

        private static ParaSnap FindStory(List<ParaSnap> paras, string needle)
        {
            if (paras == null || string.IsNullOrEmpty(needle))
            {
                return default(ParaSnap);
            }

            foreach (ParaSnap x in paras)
            {
                if ((x.Story ?? "").IndexOf(needle, StringComparison.Ordinal) >= 0)
                {
                    return x;
                }
            }

            return default(ParaSnap);
        }

        private static string InnerOfStory(Plant p, string story)
        {
            foreach (string inner in Inners(p.Body, "S"))
            {
                if (StripInTable(inner).IndexOf(story, StringComparison.Ordinal) >= 0)
                {
                    return StripInTable(inner);
                }
            }

            return "";
        }

        private static string InnerContaining(string body, string story)
        {
            foreach (string inner in Inners(body, "S"))
            {
                if (StripInTable(inner).IndexOf(story, StringComparison.Ordinal) >= 0)
                {
                    return StripInTable(inner);
                }
            }

            return "";
        }

        private static bool PlainStay(List<ParaSnap> after, params string[] names)
        {
            foreach (string n in names)
            {
                ParaSnap hit = FindStory(after, n);
                if (hit.Story == null || hit.Numbered)
                {
                    return false;
                }
            }

            return true;
        }

        private static bool MatchWant(List<ParaSnap> after, bool[] wantNum, string[] stories)
        {
            if (after == null || wantNum == null)
            {
                return false;
            }

            if (after.Count < wantNum.Length)
            {
                return false;
            }

            for (int i = 0; i < wantNum.Length; i++)
            {
                if (after[i].Numbered != wantNum[i])
                {
                    return false;
                }

                if (stories != null && i < stories.Length)
                {
                    if (Strip(after[i].Story) != stories[i]
                        && (after[i].Story ?? "").IndexOf(stories[i], StringComparison.Ordinal) < 0)
                    {
                        return false;
                    }
                }
            }

            return true;
        }

        private static bool[] WantNum(params bool[] v)
        {
            return v;
        }

        private static bool NoTags(string body)
        {
            return body != null
                && body.IndexOf("<autonum", StringComparison.OrdinalIgnoreCase) < 0
                && body.IndexOf("[list]", StringComparison.OrdinalIgnoreCase) < 0;
        }

        private static bool HasDecimalPrefix(string inner)
        {
            inner = StripInTable(inner ?? "").TrimStart();
            if (string.IsNullOrEmpty(inner))
            {
                return false;
            }

            // 与 AutonumConvert：1. / 1、 / 一、 都是十进制，不要求 Word 的点号。
            return AutonumConvert.StripNewByOriginalTypes(new[] { "1." }, inner) != inner;
        }

        private static bool HasBulletPrefix(string inner)
        {
            inner = StripInTable(inner ?? "").TrimStart();
            if (inner.Length == 0)
            {
                return false;
            }

            char c = inner[0];
            return c == '•' || c == '·' || c == '●' || c == '○' || c == '▪' || c == '◦' || c == ''
                || c >= 0xF000
                || (c == '-' && inner.StartsWith("- ", StringComparison.Ordinal));
        }

        private static bool HasAutoPrefix(string inner)
        {
            return HasDecimalPrefix(inner) || HasBulletPrefix(inner);
        }

        private static string StripInTable(string inner)
        {
            if (string.IsNullOrEmpty(inner))
            {
                return "";
            }

            return Regex.Replace(inner, @"^\[in_table=[^\]]+\]", "");
        }

        private static string FirstInner(string body, string prefix)
        {
            List<string> list = Inners(body, prefix);
            return list.Count == 0 ? "" : list[0];
        }

        private static List<string> Inners(string body, string prefix)
        {
            var list = new List<string>();
            if (string.IsNullOrEmpty(body))
            {
                return list;
            }

            foreach (Match m in SpanRe.Matches(body))
            {
                if (m.Groups[1].Value == prefix)
                {
                    list.Add(m.Groups[3].Value);
                }
            }

            return list;
        }

        private static List<string> Codes(string body, string prefix)
        {
            var list = new List<string>();
            if (string.IsNullOrEmpty(body))
            {
                return list;
            }

            foreach (Match m in SpanRe.Matches(body))
            {
                if (m.Groups[1].Value == prefix)
                {
                    list.Add(prefix + "_" + m.Groups[2].Value);
                }
            }

            return list;
        }

        private static string JoinCodes(Plant p, string prefix)
        {
            return string.Join(",", Codes(p.Body, prefix).ToArray());
        }

        /// <summary>
        /// display 对内以任意宿主 ListString 为前缀、后面跟 COM 存文即算编号句。
        /// 不依赖 1. / 1、 / • 某一种写法；对不上也不把全文 S_ 交出去。
        /// </summary>
        private static bool NumberedDisplayMatchesWant(string inner, HashSet<string> want)
        {
            if (want == null || want.Count == 0)
            {
                return false;
            }

            string core = Strip(inner);
            if (core.Length == 0)
            {
                return false;
            }

            if (want.Contains(core))
            {
                return true;
            }

            foreach (string w in want)
            {
                if (!string.IsNullOrEmpty(w)
                    && core.Length > w.Length
                    && core.EndsWith(w, StringComparison.Ordinal)
                    && core.Substring(0, core.Length - w.Length).Trim().Length > 0)
                {
                    return true;
                }
            }

            return false;
        }

        private static string JoinNumberedS(Plant p)
        {
            var want = new HashSet<string>(StringComparer.Ordinal);
            if (p.Paras != null)
            {
                foreach (ParaSnap x in p.Paras)
                {
                    if (x.Numbered && !string.IsNullOrEmpty(x.Story))
                    {
                        want.Add(Strip(x.Story));
                    }
                }
            }

            var codes = new List<string>();
            foreach (Match m in SpanRe.Matches(p.Body ?? ""))
            {
                if (m.Groups[1].Value != "S")
                {
                    continue;
                }

                string inner = Strip(StripInTable(m.Groups[3].Value));
                if (NumberedDisplayMatchesWant(inner, want))
                {
                    codes.Add("S_" + m.Groups[2].Value);
                }
            }

            if (codes.Count == 0)
            {
                Console.WriteLine("[JoinNumberedS] unmatched want=["
                    + string.Join("|", want)
                    + "] 不退化成全文 S_");
                return "";
            }

            return string.Join(",", codes.ToArray());
        }

        private static string TableId(string body)
        {
            Match m = Regex.Match(body ?? "", @"id=""(T_[^""]+)""");
            return m.Success ? m.Groups[1].Value : "";
        }

        private static string CellInner(string body, string needle)
        {
            return InnerContaining(body, needle);
        }

        private static string CodeOfInner(string body, string needle)
        {
            if (string.IsNullOrEmpty(body))
            {
                return "";
            }

            foreach (Match m in SpanRe.Matches(body))
            {
                if (m.Groups[3].Value.IndexOf(needle, StringComparison.Ordinal) >= 0)
                {
                    return m.Groups[1].Value + "_" + m.Groups[2].Value;
                }
            }

            return "";
        }

        private static string JoinN(int n)
        {
            var sb = new System.Text.StringBuilder();
            for (int i = 0; i < n; i++)
            {
                if (n <= 3)
                {
                    sb.Append((char)('a' + i));
                }
                else
                {
                    sb.Append("i");
                    sb.Append(i);
                }

                sb.Append('\r');
            }

            return sb.ToString();
        }

        private static string Strip(string text)
        {
            if (string.IsNullOrEmpty(text))
            {
                return "";
            }

            return text.TrimEnd('\r', '\a', '\n', '\u0007');
        }

        private static string Dump(List<ParaSnap> paras)
        {
            if (paras == null)
            {
                return "";
            }

            var parts = new List<string>();
            foreach (ParaSnap x in paras)
            {
                parts.Add((x.Numbered ? "[" + x.List + "]" : "[ ]") + Strip(x.Story));
            }

            return string.Join(" | ", parts.ToArray());
        }

        private static string Snip(string body)
        {
            if (string.IsNullOrEmpty(body))
            {
                return "";
            }

            return body.Length <= 160 ? body : body.Substring(0, 160);
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

        private static List<Dictionary<string, object>> Replace(string codes, string neu)
        {
            return One("replace", new Dictionary<string, object>
            {
                ["original"] = new Dictionary<string, object> { ["codes"] = codes },
                ["new"] = neu,
                ["format"] = DefaultFormat()
            });
        }

        private static List<Dictionary<string, object>> ReplacePara(string codes, string neu)
        {
            return One("replace", new Dictionary<string, object>
            {
                ["original"] = new Dictionary<string, object> { ["paragraph_codes"] = codes },
                ["new"] = neu,
                ["format"] = DefaultFormat()
            });
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

        private static List<Dictionary<string, object>> InsertAfter(string codes, string insert)
        {
            return One("insert", new Dictionary<string, object>
            {
                ["target"] = new Dictionary<string, object> { ["codes"] = codes },
                ["insert"] = insert,
                ["format"] = DefaultFormat()
            });
        }

        private static List<Dictionary<string, object>> InsertStart(string insert)
        {
            return One("insert", new Dictionary<string, object>
            {
                ["target"] = new Dictionary<string, object> { ["position"] = "start" },
                ["insert"] = insert,
                ["format"] = DefaultFormat()
            });
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
