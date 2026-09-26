using System;
using System.Collections.Generic;
using WordAddIn1;

namespace PptHtmlContentRoundtripTest
{
    /// <summary>管页加巩固。前缀 sld-e-，100 条。</summary>
    internal static partial class SlideManageRunner
    {
        private static void RunExtraParse(TestRun run, IList<string> filters)
        {
            if (!Want(filters, "sld-e-"))
            {
                return;
            }

            Console.WriteLine("--- I 参数（无 COM）---");
            ChannelRegistry.Register(new KindStub("sld-e-stub-excel", ChannelKind.Excel), false);

            if (Want(filters, "sld-e-action-blank-fail"))
            {
                var a = Args("sld-stub-ppt");
                a["action"] = "   ";
                ExpectFail(run, "sld-e-action-blank-fail", a, "action");
            }

            if (Want(filters, "sld-e-neg-excel"))
            {
                var a = Args("sld-e-stub-excel");
                a["action"] = "add";
                ExpectFail(run, "sld-e-neg-excel", a, "渠道");
            }
        }

        private static void RunExtra(TestRun run, IList<string> filters, Deck deck, Func<Deck> openSecond)
        {
            if (!Want(filters, "sld-e-"))
            {
                return;
            }

            Console.WriteLine("--- I–P 加巩固 COM ---");
            RunExtraArgs(run, filters, deck);
            RunExtraAdd(run, filters, deck);
            RunExtraDup(run, filters, deck);
            RunExtraDel(run, filters, deck);
            RunExtraMove(run, filters, deck);
            RunExtraReorder(run, filters, deck);
            RunExtraCross(run, filters, deck, openSecond);
            RunExtraHiddenChain(run, filters, deck);
        }

        private static void RunExtraArgs(TestRun run, IList<string> filters, Deck deck)
        {
            if (Want(filters, "sld-e-action-upper")
                || Want(filters, "sld-e-action-mixed")
                || Want(filters, "sld-e-action-spaces")
                || Want(filters, "sld-e-echo-kind")
                || Want(filters, "sld-e-echo-to-index"))
            {
                if (!deck.EnsureCount(1, out string err))
                {
                    FailNeeded(
                        run,
                        filters,
                        err,
                        "sld-e-action-upper",
                        "sld-e-action-mixed",
                        "sld-e-action-spaces",
                        "sld-e-echo-kind",
                        "sld-e-echo-to-index");
                }
                else
                {
                    if (Want(filters, "sld-e-action-upper"))
                    {
                        AssertAddAction(run, "sld-e-action-upper", deck, "ADD", "add");
                    }

                    if (Want(filters, "sld-e-action-mixed"))
                    {
                        AssertAddAction(run, "sld-e-action-mixed", deck, "Add", "add");
                    }

                    if (Want(filters, "sld-e-action-spaces"))
                    {
                        AssertAddAction(run, "sld-e-action-spaces", deck, " add ", "add");
                    }

                    if (Want(filters, "sld-e-echo-kind") || Want(filters, "sld-e-echo-to-index"))
                    {
                        var a = deck.Args();
                        a["action"] = "add";
                        a["to_index"] = 1;
                        ToolResult r = Invoke(a);
                        if (Want(filters, "sld-e-echo-kind"))
                        {
                            var fails = new List<string>();
                            if (!Ok(r, fails) || !Align(r, deck, fails))
                            {
                                Finish(run, "sld-e-echo-kind", fails);
                            }
                            else
                            {
                                string kind = GetString(r, "kind");
                                string want = deck.Wpp ? "wpp" : "ppt";
                                if (!string.Equals(kind, want, StringComparison.OrdinalIgnoreCase))
                                {
                                    fails.Add("kind=" + kind + " 期望 " + want);
                                }

                                Finish(run, "sld-e-echo-kind", fails);
                            }
                        }

                        if (Want(filters, "sld-e-echo-to-index"))
                        {
                            var fails = new List<string>();
                            if (!Ok(r, fails))
                            {
                                Finish(run, "sld-e-echo-to-index", fails);
                            }
                            else if (GetInt(r, "to_index") != 1)
                            {
                                fails.Add("to_index 应回显 1");
                                Finish(run, "sld-e-echo-to-index", fails);
                            }
                            else
                            {
                                Finish(run, "sld-e-echo-to-index", fails);
                            }
                        }
                    }
                }
            }

            if (Want(filters, "sld-e-dup-src-word-fail"))
            {
                if (!deck.EnsureCount(1, out string err))
                {
                    run.CaseFail("sld-e-dup-src-word-fail", err);
                }
                else
                {
                    int before = deck.Count();
                    var a = Dup(deck, deck.Ids()[0]);
                    a["source_channel_id"] = "sld-stub-word";
                    ExpectFailUnchanged(run, "sld-e-dup-src-word-fail", a, deck, before, "演示文稿");
                }
            }
        }

        private static void AssertAddAction(TestRun run, string id, Deck deck, string action, string echo)
        {
            var a = deck.Args();
            a["action"] = action;
            ToolResult r = Invoke(a);
            var fails = new List<string>();
            if (!Ok(r, fails) || !Align(r, deck, fails))
            {
                Finish(run, id, fails);
                return;
            }

            if (!string.Equals(GetString(r, "action"), echo, StringComparison.Ordinal))
            {
                fails.Add("action 应归一成 " + echo);
            }

            Finish(run, id, fails);
        }

        private static void RunExtraAdd(TestRun run, IList<string> filters, Deck deck)
        {
            if (!Want(filters, "sld-e-add-"))
            {
                return;
            }

            if (Want(filters, "sld-e-add-count-plus"))
            {
                if (!deck.EnsureCount(3, out string err))
                {
                    run.CaseFail("sld-e-add-count-plus", err);
                }
                else
                {
                    var a = deck.Args();
                    a["action"] = "add";
                    a["to_index"] = 4;
                    AssertNewAt(run, "sld-e-add-count-plus", Invoke(a), deck, 4);
                }
            }

            if (Want(filters, "sld-e-add-to-3-of-3"))
            {
                if (!deck.EnsureCount(3, out string err))
                {
                    run.CaseFail("sld-e-add-to-3-of-3", err);
                }
                else
                {
                    string last = deck.Ids()[2];
                    var a = deck.Args();
                    a["action"] = "add";
                    a["to_index"] = 3;
                    ToolResult r = Invoke(a);
                    var fails = new List<string>();
                    if (!Ok(r, fails) || !Align(r, deck, fails))
                    {
                        Finish(run, "sld-e-add-to-3-of-3", fails);
                    }
                    else if (IndexOf(Rows(r), GetString(r, "slide_id")) != 3 || IndexOf(Rows(r), last) != 4)
                    {
                        fails.Add("新页应在 3，原末页变 4");
                        Finish(run, "sld-e-add-to-3-of-3", fails);
                    }
                    else
                    {
                        Finish(run, "sld-e-add-to-3-of-3", fails);
                    }
                }
            }

            if (Want(filters, "sld-e-add-to-str"))
            {
                if (!deck.EnsureCount(2, out string err))
                {
                    run.CaseFail("sld-e-add-to-str", err);
                }
                else
                {
                    var a = deck.Args();
                    a["action"] = "add";
                    a["to_index"] = "2";
                    AssertNewAt(run, "sld-e-add-to-str", Invoke(a), deck, 2);
                }
            }

            if (Want(filters, "sld-e-add-layout-trim") || Want(filters, "sld-e-add-layout-empty"))
            {
                if (!deck.EnsureCount(1, out string err))
                {
                    FailNeeded(run, filters, err, "sld-e-add-layout-trim", "sld-e-add-layout-empty");
                }
                else
                {
                    var seed = deck.Args();
                    seed["action"] = "add";
                    ToolResult seedR = Invoke(seed);
                    string layout = "";
                    IList<SlideRow> seedRows = Rows(seedR);
                    if (seedRows.Count > 0)
                    {
                        layout = seedRows[seedRows.Count - 1].Layout ?? "";
                    }

                    if (Want(filters, "sld-e-add-layout-trim"))
                    {
                        var a = deck.Args();
                        a["action"] = "add";
                        if (!string.IsNullOrWhiteSpace(layout))
                        {
                            a["layout"] = "  " + layout + "  ";
                        }

                        ToolResult r = Invoke(a);
                        var fails = new List<string>();
                        if (!Ok(r, fails) || !Align(r, deck, fails))
                        {
                            Finish(run, "sld-e-add-layout-trim", fails);
                        }
                        else
                        {
                            Finish(run, "sld-e-add-layout-trim", fails);
                        }
                    }

                    if (Want(filters, "sld-e-add-layout-empty"))
                    {
                        var a = deck.Args();
                        a["action"] = "add";
                        a["layout"] = "";
                        ToolResult r = Invoke(a);
                        var fails = new List<string>();
                        if (!Ok(r, fails) || !Align(r, deck, fails))
                        {
                            Finish(run, "sld-e-add-layout-empty", fails);
                        }
                        else
                        {
                            Finish(run, "sld-e-add-layout-empty", fails);
                        }
                    }
                }
            }

            if (Want(filters, "sld-e-add-ids-stable"))
            {
                if (!deck.EnsureCount(2, out string err))
                {
                    run.CaseFail("sld-e-add-ids-stable", err);
                }
                else
                {
                    IList<string> old = deck.Ids();
                    var a = deck.Args();
                    a["action"] = "add";
                    ToolResult r = Invoke(a);
                    var fails = new List<string>();
                    if (!Ok(r, fails) || !Align(r, deck, fails))
                    {
                        Finish(run, "sld-e-add-ids-stable", fails);
                    }
                    else if (IndexOf(Rows(r), old[0]) < 1 || IndexOf(Rows(r), old[1]) < 1)
                    {
                        fails.Add("原 id 应仍在");
                        Finish(run, "sld-e-add-ids-stable", fails);
                    }
                    else
                    {
                        Finish(run, "sld-e-add-ids-stable", fails);
                    }
                }
            }

            if (Want(filters, "sld-e-add-five"))
            {
                if (!deck.EnsureCount(1, out string err))
                {
                    run.CaseFail("sld-e-add-five", err);
                }
                else
                {
                    var ids = new HashSet<string>(deck.Ids(), StringComparer.Ordinal);
                    var fails = new List<string>();
                    for (int i = 0; i < 5; i++)
                    {
                        var a = deck.Args();
                        a["action"] = "add";
                        ToolResult r = Invoke(a);
                        if (!Ok(r, fails))
                        {
                            break;
                        }

                        string nid = GetString(r, "slide_id");
                        if (string.IsNullOrEmpty(nid) || !ids.Add(nid))
                        {
                            fails.Add("第 " + (i + 1) + " 次 add id 不新");
                            break;
                        }
                    }

                    if (fails.Count == 0 && (deck.Count() != 6 || ids.Count != 6))
                    {
                        fails.Add("连加 5 次后应为 6 个不同 id");
                    }

                    Finish(run, "sld-e-add-five", fails);
                }
            }

            if (Want(filters, "sld-e-add-after-fail"))
            {
                if (!deck.EnsureCount(1, out string err))
                {
                    run.CaseFail("sld-e-add-after-fail", err);
                }
                else
                {
                    var bad = deck.Args();
                    bad["action"] = "add";
                    bad["to_index"] = 0;
                    Invoke(bad);
                    int before = deck.Count();
                    var a = deck.Args();
                    a["action"] = "add";
                    ToolResult r = Invoke(a);
                    var fails = new List<string>();
                    if (!Ok(r, fails) || GetInt(r, "slide_count") != before + 1)
                    {
                        if (fails.Count == 0)
                        {
                            fails.Add("失败后再 add 应成功 +1");
                        }

                        Finish(run, "sld-e-add-after-fail", fails);
                    }
                    else
                    {
                        Finish(run, "sld-e-add-after-fail", fails);
                    }
                }
            }

            if (Want(filters, "sld-e-add-slides-seq"))
            {
                if (!deck.EnsureCount(1, out string err))
                {
                    run.CaseFail("sld-e-add-slides-seq", err);
                }
                else
                {
                    var a = deck.Args();
                    a["action"] = "add";
                    ToolResult r = Invoke(a);
                    var fails = new List<string>();
                    if (!Ok(r, fails))
                    {
                        Finish(run, "sld-e-add-slides-seq", fails);
                    }
                    else
                    {
                        IList<SlideRow> rows = Rows(r);
                        for (int i = 0; i < rows.Count; i++)
                        {
                            if (rows[i].Index != i + 1)
                            {
                                fails.Add("index 须连续 1..n");
                                break;
                            }
                        }

                        Finish(run, "sld-e-add-slides-seq", fails);
                    }
                }
            }

            if (Want(filters, "sld-e-add-focus-to-index"))
            {
                if (!deck.EnsureCount(2, out string err))
                {
                    run.CaseFail("sld-e-add-focus-to-index", err);
                }
                else
                {
                    var a = deck.Args();
                    a["action"] = "add";
                    a["to_index"] = 2;
                    ToolResult r = Invoke(a);
                    var fails = new List<string>();
                    if (!Ok(r, fails) || GetInt(r, "to_index") != 2
                        || IndexOf(Rows(r), GetString(r, "slide_id")) != 2)
                    {
                        if (fails.Count == 0)
                        {
                            fails.Add("焦点 to_index 应=新页 index");
                        }

                        Finish(run, "sld-e-add-focus-to-index", fails);
                    }
                    else
                    {
                        Finish(run, "sld-e-add-focus-to-index", fails);
                    }
                }
            }

            if (Want(filters, "sld-e-add-old-shift"))
            {
                if (!deck.EnsureCount(2, out string err))
                {
                    run.CaseFail("sld-e-add-old-shift", err);
                }
                else
                {
                    IList<string> old = deck.Ids();
                    var a = deck.Args();
                    a["action"] = "add";
                    a["to_index"] = 1;
                    ToolResult r = Invoke(a);
                    var fails = new List<string>();
                    if (!Ok(r, fails) || IndexOf(Rows(r), old[0]) != 2 || IndexOf(Rows(r), old[1]) != 3)
                    {
                        if (fails.Count == 0)
                        {
                            fails.Add("插到 1 后原页应变 2、3");
                        }

                        Finish(run, "sld-e-add-old-shift", fails);
                    }
                    else
                    {
                        Finish(run, "sld-e-add-old-shift", fails);
                    }
                }
            }

            if (Want(filters, "sld-e-add-between"))
            {
                if (!deck.EnsureCount(3, out string err))
                {
                    run.CaseFail("sld-e-add-between", err);
                }
                else
                {
                    IList<string> old = deck.Ids();
                    var a = deck.Args();
                    a["action"] = "add";
                    a["to_index"] = 2;
                    ToolResult r = Invoke(a);
                    var fails = new List<string>();
                    if (!Ok(r, fails) || !Align(r, deck, fails))
                    {
                        Finish(run, "sld-e-add-between", fails);
                    }
                    else if (IndexOf(Rows(r), old[0]) != 1
                        || IndexOf(Rows(r), GetString(r, "slide_id")) != 2
                        || IndexOf(Rows(r), old[1]) != 3
                        || IndexOf(Rows(r), old[2]) != 4)
                    {
                        fails.Add("3 页插 2 应为 a,新,b,c");
                        Finish(run, "sld-e-add-between", fails);
                    }
                    else
                    {
                        Finish(run, "sld-e-add-between", fails);
                    }
                }
            }

        }

        private static void RunExtraDup(TestRun run, IList<string> filters, Deck deck)
        {
            if (!Want(filters, "sld-e-dup-"))
            {
                return;
            }

            if (Want(filters, "sld-e-dup-count-plus"))
            {
                if (!deck.EnsureCount(3, out string err))
                {
                    run.CaseFail("sld-e-dup-count-plus", err);
                }
                else
                {
                    var a = Dup(deck, deck.Ids()[0]);
                    a["to_index"] = 4;
                    AssertNewAt(run, "sld-e-dup-count-plus", Invoke(a), deck, 4);
                }
            }

            if (Want(filters, "sld-e-dup-to-last"))
            {
                if (!deck.EnsureCount(3, out string err))
                {
                    run.CaseFail("sld-e-dup-to-last", err);
                }
                else
                {
                    string last = deck.Ids()[2];
                    var a = Dup(deck, deck.Ids()[0]);
                    a["to_index"] = 3;
                    ToolResult r = Invoke(a);
                    var fails = new List<string>();
                    if (!Ok(r, fails) || !Align(r, deck, fails))
                    {
                        Finish(run, "sld-e-dup-to-last", fails);
                    }
                    else if (IndexOf(Rows(r), GetString(r, "slide_id")) != 3 || IndexOf(Rows(r), last) != 4)
                    {
                        fails.Add("新页应在 3，原末页变 4");
                        Finish(run, "sld-e-dup-to-last", fails);
                    }
                    else
                    {
                        Finish(run, "sld-e-dup-to-last", fails);
                    }
                }
            }

            if (Want(filters, "sld-e-dup-to-2"))
            {
                if (!deck.EnsureCount(3, out string err))
                {
                    run.CaseFail("sld-e-dup-to-2", err);
                }
                else
                {
                    var a = Dup(deck, deck.Ids()[2]);
                    a["to_index"] = 2;
                    AssertNewAt(run, "sld-e-dup-to-2", Invoke(a), deck, 2);
                }
            }

            if (Want(filters, "sld-e-dup-last-tail") || Want(filters, "sld-e-dup-mid-tail"))
            {
                if (!deck.EnsureCount(3, out string err))
                {
                    FailNeeded(run, filters, err, "sld-e-dup-last-tail", "sld-e-dup-mid-tail");
                }
                else
                {
                    if (Want(filters, "sld-e-dup-last-tail"))
                    {
                        AssertNewAt(run, "sld-e-dup-last-tail", Invoke(Dup(deck, deck.Ids()[2])), deck, 4);
                    }

                    if (Want(filters, "sld-e-dup-mid-tail"))
                    {
                        if (!deck.EnsureCount(3, out err))
                        {
                            run.CaseFail("sld-e-dup-mid-tail", err);
                        }
                        else
                        {
                            AssertNewAt(run, "sld-e-dup-mid-tail", Invoke(Dup(deck, deck.Ids()[1])), deck, 4);
                        }
                    }
                }
            }

            if (Want(filters, "sld-e-dup-ids-three"))
            {
                if (!deck.EnsureCount(3, out string err))
                {
                    run.CaseFail("sld-e-dup-ids-three", err);
                }
                else
                {
                    IList<string> abc = deck.Ids();
                    var a = deck.Args();
                    a["action"] = "duplicate";
                    a["slide_ids"] = new[] { abc[0], abc[1], abc[2] };
                    ToolResult r = Invoke(a);
                    var fails = new List<string>();
                    IList<CreatedRow> cr = Created(r);
                    if (!Ok(r, fails) || GetInt(r, "slide_count") != 6 || cr.Count != 3)
                    {
                        if (fails.Count == 0)
                        {
                            fails.Add("应 +3 且 created=3");
                        }

                        Finish(run, "sld-e-dup-ids-three", fails);
                    }
                    else
                    {
                        Finish(run, "sld-e-dup-ids-three", fails);
                    }
                }
            }

            if (Want(filters, "sld-e-dup-ids-rev"))
            {
                if (!deck.EnsureCount(3, out string err))
                {
                    run.CaseFail("sld-e-dup-ids-rev", err);
                }
                else
                {
                    IList<string> abc = deck.Ids();
                    var a = deck.Args();
                    a["action"] = "duplicate";
                    a["slide_ids"] = new[] { abc[2], abc[1], abc[0] };
                    ToolResult r = Invoke(a);
                    var fails = new List<string>();
                    IList<CreatedRow> cr = Created(r);
                    if (!Ok(r, fails) || cr.Count != 3
                        || cr[0].Source != abc[2] || cr[1].Source != abc[1] || cr[2].Source != abc[0])
                    {
                        if (fails.Count == 0)
                        {
                            fails.Add("created 源序应为 c,b,a");
                        }

                        Finish(run, "sld-e-dup-ids-rev", fails);
                    }
                    else
                    {
                        Finish(run, "sld-e-dup-ids-rev", fails);
                    }
                }
            }

            if (Want(filters, "sld-e-dup-ids-empty-item-fail"))
            {
                if (!deck.EnsureCount(3, out string err))
                {
                    run.CaseFail("sld-e-dup-ids-empty-item-fail", err);
                }
                else
                {
                    var a = deck.Args();
                    a["action"] = "duplicate";
                    a["slide_ids"] = new[] { deck.Ids()[0], "" };
                    ExpectFailUnchanged(run, "sld-e-dup-ids-empty-item-fail", a, deck, 3, "空");
                }
            }

            if (Want(filters, "sld-e-dup-no-id-fail"))
            {
                if (!deck.EnsureCount(2, out string err))
                {
                    run.CaseFail("sld-e-dup-no-id-fail", err);
                }
                else
                {
                    var a = deck.Args();
                    a["action"] = "duplicate";
                    ExpectFailUnchanged(run, "sld-e-dup-no-id-fail", a, deck, 2, "slide_id");
                }
            }

            if (Want(filters, "sld-e-dup-empty-id-fail"))
            {
                if (!deck.EnsureCount(2, out string err))
                {
                    run.CaseFail("sld-e-dup-empty-id-fail", err);
                }
                else
                {
                    var a = deck.Args();
                    a["action"] = "duplicate";
                    a["slide_id"] = "";
                    ExpectFailUnchanged(run, "sld-e-dup-empty-id-fail", a, deck, 2, "slide_id");
                }
            }

            if (Want(filters, "sld-e-dup-id-spaces"))
            {
                if (!deck.EnsureCount(2, out string err))
                {
                    run.CaseFail("sld-e-dup-id-spaces", err);
                }
                else
                {
                    string aId = deck.Ids()[0];
                    var a = deck.Args();
                    a["action"] = "duplicate";
                    a["slide_id"] = "  " + aId + "  ";
                    ToolResult r = Invoke(a);
                    var fails = new List<string>();
                    IList<CreatedRow> cr = Created(r);
                    if (!Ok(r, fails) || cr.Count != 1 || cr[0].Source != aId)
                    {
                        if (fails.Count == 0)
                        {
                            fails.Add("应 trim slide_id");
                        }

                        Finish(run, "sld-e-dup-id-spaces", fails);
                    }
                    else
                    {
                        Finish(run, "sld-e-dup-id-spaces", fails);
                    }
                }
            }

            if (Want(filters, "sld-e-dup-to-0-fail") || Want(filters, "sld-e-dup-to-neg-fail"))
            {
                if (!deck.EnsureCount(3, out string err))
                {
                    FailNeeded(run, filters, err, "sld-e-dup-to-0-fail", "sld-e-dup-to-neg-fail");
                }
                else
                {
                    string aId = deck.Ids()[0];
                    if (Want(filters, "sld-e-dup-to-0-fail"))
                    {
                        var z = Dup(deck, aId);
                        z["to_index"] = 0;
                        ExpectFailUnchanged(run, "sld-e-dup-to-0-fail", z, deck, 3, "越界");
                    }

                    if (Want(filters, "sld-e-dup-to-neg-fail"))
                    {
                        var n = Dup(deck, aId);
                        n["to_index"] = -1;
                        ExpectFailUnchanged(run, "sld-e-dup-to-neg-fail", n, deck, 3, "越界");
                    }
                }
            }

            if (Want(filters, "sld-e-dup-one-page"))
            {
                if (!deck.EnsureCount(1, out string err))
                {
                    run.CaseFail("sld-e-dup-one-page", err);
                }
                else
                {
                    ToolResult r = Invoke(Dup(deck, deck.Ids()[0]));
                    var fails = new List<string>();
                    if (!Ok(r, fails) || GetInt(r, "slide_count") != 2 || !Align(r, deck, fails))
                    {
                        if (fails.Count == 0)
                        {
                            fails.Add("单页拷应成 2 页");
                        }

                        Finish(run, "sld-e-dup-one-page", fails);
                    }
                    else
                    {
                        Finish(run, "sld-e-dup-one-page", fails);
                    }
                }
            }

            if (Want(filters, "sld-e-dup-ids-list"))
            {
                if (!deck.EnsureCount(3, out string err))
                {
                    run.CaseFail("sld-e-dup-ids-list", err);
                }
                else
                {
                    var a = deck.Args();
                    a["action"] = "duplicate";
                    a["slide_ids"] = new List<string> { deck.Ids()[0] };
                    ToolResult r = Invoke(a);
                    var fails = new List<string>();
                    if (!Ok(r, fails) || Created(r).Count != 1)
                    {
                        if (fails.Count == 0)
                        {
                            fails.Add("List slide_ids 应能解");
                        }

                        Finish(run, "sld-e-dup-ids-list", fails);
                    }
                    else
                    {
                        Finish(run, "sld-e-dup-ids-list", fails);
                    }
                }
            }

            if (Want(filters, "sld-e-dup-created-index") || Want(filters, "sld-e-dup-new-unique")
                || Want(filters, "sld-e-dup-src-index-stable") || Want(filters, "sld-e-dup-action-echo"))
            {
                if (!deck.EnsureCount(3, out string err))
                {
                    FailNeeded(
                        run,
                        filters,
                        err,
                        "sld-e-dup-created-index",
                        "sld-e-dup-new-unique",
                        "sld-e-dup-src-index-stable",
                        "sld-e-dup-action-echo");
                }
                else
                {
                    IList<string> old = deck.Ids();
                    string aId = old[0];
                    ToolResult r = Invoke(Dup(deck, aId));
                    var fails = new List<string>();
                    if (!Ok(r, fails))
                    {
                        FailNeeded(
                            run,
                            filters,
                            fails[0],
                            "sld-e-dup-created-index",
                            "sld-e-dup-new-unique",
                            "sld-e-dup-src-index-stable",
                            "sld-e-dup-action-echo");
                    }
                    else
                    {
                        IList<CreatedRow> cr = Created(r);
                        string nid = GetString(r, "slide_id");
                        if (Want(filters, "sld-e-dup-created-index"))
                        {
                            var f = new List<string>();
                            if (cr.Count != 1 || cr[0].Index != IndexOf(Rows(r), nid))
                            {
                                f.Add("created.index 应等于新页 index");
                            }

                            Finish(run, "sld-e-dup-created-index", f);
                        }

                        if (Want(filters, "sld-e-dup-new-unique"))
                        {
                            var f = new List<string>();
                            if (IndexOf(Rows(r), aId) < 1 || string.Equals(nid, aId, StringComparison.Ordinal))
                            {
                                f.Add("新 id 须不同于源且源仍在");
                            }

                            Finish(run, "sld-e-dup-new-unique", f);
                        }

                        if (Want(filters, "sld-e-dup-src-index-stable"))
                        {
                            var f = new List<string>();
                            if (IndexOf(Rows(r), aId) != 1)
                            {
                                f.Add("尾部拷后源页应仍在 index=1");
                            }

                            Finish(run, "sld-e-dup-src-index-stable", f);
                        }

                        if (Want(filters, "sld-e-dup-action-echo"))
                        {
                            var f = new List<string>();
                            if (!string.Equals(GetString(r, "action"), "duplicate", StringComparison.Ordinal))
                            {
                                f.Add("action 应为 duplicate");
                            }

                            Finish(run, "sld-e-dup-action-echo", f);
                        }
                    }
                }
            }

            if (Want(filters, "sld-e-dup-to-1-seq"))
            {
                if (!deck.EnsureCount(3, out string err))
                {
                    run.CaseFail("sld-e-dup-to-1-seq", err);
                }
                else
                {
                    IList<string> abc = deck.Ids();
                    var a = deck.Args();
                    a["action"] = "duplicate";
                    a["slide_ids"] = new[] { abc[0], abc[2] };
                    a["to_index"] = 1;
                    ToolResult r = Invoke(a);
                    var fails = new List<string>();
                    IList<CreatedRow> cr = Created(r);
                    if (!Ok(r, fails) || cr.Count != 2
                        || IndexOf(Rows(r), cr[0].NewId) != 1
                        || IndexOf(Rows(r), cr[1].NewId) != 2)
                    {
                        if (fails.Count == 0)
                        {
                            fails.Add("连续插 1 应落在 1、2");
                        }

                        Finish(run, "sld-e-dup-to-1-seq", fails);
                    }
                    else
                    {
                        Finish(run, "sld-e-dup-to-1-seq", fails);
                    }
                }
            }

            if (Want(filters, "sld-e-dup-missing-second-fail"))
            {
                if (!deck.EnsureCount(3, out string err))
                {
                    run.CaseFail("sld-e-dup-missing-second-fail", err);
                }
                else
                {
                    var a = deck.Args();
                    a["action"] = "duplicate";
                    a["slide_ids"] = new[] { deck.Ids()[0], "999999" };
                    ExpectFailUnchanged(run, "sld-e-dup-missing-second-fail", a, deck, 3, "不存在");
                }
            }

            if (Want(filters, "sld-e-dup-after-move"))
            {
                if (!deck.EnsureCount(3, out string err))
                {
                    run.CaseFail("sld-e-dup-after-move", err);
                }
                else
                {
                    string aId = deck.Ids()[0];
                    Invoke(Move(deck, aId, 3));
                    ToolResult r = Invoke(Dup(deck, aId));
                    var fails = new List<string>();
                    if (!Ok(r, fails) || IndexOf(Rows(r), aId) < 1)
                    {
                        if (fails.Count == 0)
                        {
                            fails.Add("挪页后仍应按 id 拷到");
                        }

                        Finish(run, "sld-e-dup-after-move", fails);
                    }
                    else
                    {
                        Finish(run, "sld-e-dup-after-move", fails);
                    }
                }
            }
        }

        private static void RunExtraDel(TestRun run, IList<string> filters, Deck deck)
        {
            if (!Want(filters, "sld-e-del-"))
            {
                return;
            }

            if (Want(filters, "sld-e-del-first"))
            {
                if (!deck.EnsureCount(3, out string err))
                {
                    run.CaseFail("sld-e-del-first", err);
                }
                else
                {
                    string first = deck.Ids()[0];
                    AssertDeleted(run, "sld-e-del-first", Invoke(Del(deck, first)), deck, first, 2);
                }
            }

            if (Want(filters, "sld-e-del-last-of-3"))
            {
                if (!deck.EnsureCount(3, out string err))
                {
                    run.CaseFail("sld-e-del-last-of-3", err);
                }
                else
                {
                    string last = deck.Ids()[2];
                    AssertDeleted(run, "sld-e-del-last-of-3", Invoke(Del(deck, last)), deck, last, 2);
                }
            }

            if (Want(filters, "sld-e-del-three-of-five"))
            {
                if (!deck.EnsureCount(5, out string err))
                {
                    run.CaseFail("sld-e-del-three-of-five", err);
                }
                else
                {
                    IList<string> ids = deck.Ids();
                    var a = deck.Args();
                    a["action"] = "delete";
                    a["confirm"] = true;
                    a["slide_ids"] = new[] { ids[1], ids[2], ids[3] };
                    ToolResult r = Invoke(a);
                    var fails = new List<string>();
                    if (!Ok(r, fails) || GetInt(r, "slide_count") != 2
                        || IndexOf(Rows(r), ids[0]) != 1 || IndexOf(Rows(r), ids[4]) != 2)
                    {
                        if (fails.Count == 0)
                        {
                            fails.Add("应剩首尾两页");
                        }

                        Finish(run, "sld-e-del-three-of-five", fails);
                    }
                    else
                    {
                        Finish(run, "sld-e-del-three-of-five", fails);
                    }
                }
            }

            if (Want(filters, "sld-e-del-ids-rev"))
            {
                if (!deck.EnsureCount(4, out string err))
                {
                    run.CaseFail("sld-e-del-ids-rev", err);
                }
                else
                {
                    IList<string> ids = deck.Ids();
                    var a = deck.Args();
                    a["action"] = "delete";
                    a["confirm"] = true;
                    a["slide_ids"] = new[] { ids[2], ids[0] };
                    ToolResult r = Invoke(a);
                    IList<string> del = DeletedList(r);
                    var fails = new List<string>();
                    if (!Ok(r, fails) || del.Count != 2 || del[0] != ids[2] || del[1] != ids[0])
                    {
                        if (fails.Count == 0)
                        {
                            fails.Add("deleted 序应跟 slide_ids");
                        }

                        Finish(run, "sld-e-del-ids-rev", fails);
                    }
                    else
                    {
                        Finish(run, "sld-e-del-ids-rev", fails);
                    }
                }
            }

            if (Want(filters, "sld-e-del-partial-missing-fail"))
            {
                if (!deck.EnsureCount(3, out string err))
                {
                    run.CaseFail("sld-e-del-partial-missing-fail", err);
                }
                else
                {
                    var a = deck.Args();
                    a["action"] = "delete";
                    a["confirm"] = true;
                    a["slide_ids"] = new[] { deck.Ids()[0], "999999" };
                    ExpectFailUnchanged(run, "sld-e-del-partial-missing-fail", a, deck, 3, "不存在");
                }
            }

            if (Want(filters, "sld-e-del-no-id-fail"))
            {
                if (!deck.EnsureCount(2, out string err))
                {
                    run.CaseFail("sld-e-del-no-id-fail", err);
                }
                else
                {
                    var a = deck.Args();
                    a["action"] = "delete";
                    a["confirm"] = true;
                    ExpectFailUnchanged(run, "sld-e-del-no-id-fail", a, deck, 2, "slide_id");
                }
            }

            if (Want(filters, "sld-e-del-empty-ids-fail"))
            {
                if (!deck.EnsureCount(2, out string err))
                {
                    run.CaseFail("sld-e-del-empty-ids-fail", err);
                }
                else
                {
                    var a = deck.Args();
                    a["action"] = "delete";
                    a["confirm"] = true;
                    a["slide_ids"] = new string[0];
                    ExpectFailUnchanged(run, "sld-e-del-empty-ids-fail", a, deck, 2, "不能为空");
                }
            }

            if (Want(filters, "sld-e-del-confirm-str"))
            {
                if (!deck.EnsureCount(3, out string err))
                {
                    run.CaseFail("sld-e-del-confirm-str", err);
                }
                else
                {
                    string b = deck.Ids()[1];
                    var a = deck.Args();
                    a["action"] = "delete";
                    a["slide_id"] = b;
                    a["confirm"] = "true";
                    AssertDeleted(run, "sld-e-del-confirm-str", Invoke(a), deck, b, 2);
                }
            }

            if (Want(filters, "sld-e-del-remain-ids"))
            {
                if (!deck.EnsureCount(3, out string err))
                {
                    run.CaseFail("sld-e-del-remain-ids", err);
                }
                else
                {
                    IList<string> old = deck.Ids();
                    ToolResult r = Invoke(Del(deck, old[1]));
                    var fails = new List<string>();
                    if (!Ok(r, fails) || IndexOf(Rows(r), old[0]) < 1 || IndexOf(Rows(r), old[2]) < 1
                        || IndexOf(Rows(r), old[1]) >= 0)
                    {
                        if (fails.Count == 0)
                        {
                            fails.Add("剩下两页 id 应不变");
                        }

                        Finish(run, "sld-e-del-remain-ids", fails);
                    }
                    else
                    {
                        Finish(run, "sld-e-del-remain-ids", fails);
                    }
                }
            }

            if (Want(filters, "sld-e-del-ids-wins"))
            {
                if (!deck.EnsureCount(3, out string err))
                {
                    run.CaseFail("sld-e-del-ids-wins", err);
                }
                else
                {
                    IList<string> abc = deck.Ids();
                    var a = deck.Args();
                    a["action"] = "delete";
                    a["confirm"] = true;
                    a["slide_id"] = abc[1];
                    a["slide_ids"] = new[] { abc[0] };
                    ToolResult r = Invoke(a);
                    var fails = new List<string>();
                    if (!Ok(r, fails) || IndexOf(Rows(r), abc[0]) >= 0 || IndexOf(Rows(r), abc[1]) < 1)
                    {
                        if (fails.Count == 0)
                        {
                            fails.Add("应以 slide_ids 为准删 a 留 b");
                        }

                        Finish(run, "sld-e-del-ids-wins", fails);
                    }
                    else
                    {
                        Finish(run, "sld-e-del-ids-wins", fails);
                    }
                }
            }

            if (Want(filters, "sld-e-del-one-of-two"))
            {
                if (!deck.EnsureCount(2, out string err))
                {
                    run.CaseFail("sld-e-del-one-of-two", err);
                }
                else
                {
                    AssertDeleted(run, "sld-e-del-one-of-two", Invoke(Del(deck, deck.Ids()[0])), deck, null, 1);
                }
            }

            if (Want(filters, "sld-e-del-id-spaces"))
            {
                if (!deck.EnsureCount(3, out string err))
                {
                    run.CaseFail("sld-e-del-id-spaces", err);
                }
                else
                {
                    string b = deck.Ids()[1];
                    var a = deck.Args();
                    a["action"] = "delete";
                    a["confirm"] = true;
                    a["slide_id"] = "  " + b + "  ";
                    AssertDeleted(run, "sld-e-del-id-spaces", Invoke(a), deck, b, 2);
                }
            }

            if (Want(filters, "sld-e-del-empty-item-fail"))
            {
                if (!deck.EnsureCount(3, out string err))
                {
                    run.CaseFail("sld-e-del-empty-item-fail", err);
                }
                else
                {
                    var a = deck.Args();
                    a["action"] = "delete";
                    a["confirm"] = true;
                    a["slide_ids"] = new[] { deck.Ids()[0], "" };
                    ExpectFailUnchanged(run, "sld-e-del-empty-item-fail", a, deck, 3, "空");
                }
            }

            if (Want(filters, "sld-e-del-after-fail"))
            {
                if (!deck.EnsureCount(3, out string err))
                {
                    run.CaseFail("sld-e-del-after-fail", err);
                }
                else
                {
                    Invoke(Del(deck, "999999"));
                    string b = deck.Ids()[1];
                    AssertDeleted(run, "sld-e-del-after-fail", Invoke(Del(deck, b)), deck, b, 2);
                }
            }

            if (Want(filters, "sld-e-del-ids-one-arr"))
            {
                if (!deck.EnsureCount(3, out string err))
                {
                    run.CaseFail("sld-e-del-ids-one-arr", err);
                }
                else
                {
                    string b = deck.Ids()[1];
                    var a = deck.Args();
                    a["action"] = "delete";
                    a["confirm"] = true;
                    a["slide_ids"] = new[] { b };
                    AssertDeleted(run, "sld-e-del-ids-one-arr", Invoke(a), deck, b, 2);
                }
            }

            if (Want(filters, "sld-e-del-display"))
            {
                if (!deck.EnsureCount(3, out string err))
                {
                    run.CaseFail("sld-e-del-display", err);
                }
                else
                {
                    string b = deck.Ids()[1];
                    ToolResult r = Invoke(Del(deck, b));
                    var fails = new List<string>();
                    string display = GetString(r, "display_contents");
                    if (!Ok(r, fails) || display.IndexOf("deleted=", StringComparison.Ordinal) < 0
                        || display.IndexOf(b, StringComparison.Ordinal) < 0)
                    {
                        if (fails.Count == 0)
                        {
                            fails.Add("display_contents 应含 deleted=");
                        }

                        Finish(run, "sld-e-del-display", fails);
                    }
                    else
                    {
                        Finish(run, "sld-e-del-display", fails);
                    }
                }
            }
        }

        private static void RunExtraMove(TestRun run, IList<string> filters, Deck deck)
        {
            if (!Want(filters, "sld-e-move-"))
            {
                return;
            }

            if (Want(filters, "sld-e-move-adj-down"))
            {
                AbcMove(run, "sld-e-move-adj-down", deck, ids => Move(deck, ids[0], 2), (r, ids) => OrderIs(Rows(r), ids[1], ids[0], ids[2]));
            }

            if (Want(filters, "sld-e-move-adj-up"))
            {
                AbcMove(run, "sld-e-move-adj-up", deck, ids => Move(deck, ids[2], 2), (r, ids) => OrderIs(Rows(r), ids[0], ids[2], ids[1]));
            }

            if (Want(filters, "sld-e-move-mid-end"))
            {
                AbcMove(run, "sld-e-move-mid-end", deck, ids => Move(deck, ids[1], 3), (r, ids) => OrderIs(Rows(r), ids[0], ids[2], ids[1]));
            }

            if (Want(filters, "sld-e-move-first-same"))
            {
                AbcMove(run, "sld-e-move-first-same", deck, ids => Move(deck, ids[0], 1), (r, ids) => OrderIs(Rows(r), ids[0], ids[1], ids[2]));
            }

            if (Want(filters, "sld-e-move-last-same"))
            {
                AbcMove(run, "sld-e-move-last-same", deck, ids => Move(deck, ids[2], 3), (r, ids) => OrderIs(Rows(r), ids[0], ids[1], ids[2]));
            }

            if (Want(filters, "sld-e-move-no-id-fail"))
            {
                if (!deck.EnsureCount(3, out string err))
                {
                    run.CaseFail("sld-e-move-no-id-fail", err);
                }
                else
                {
                    var a = deck.Args();
                    a["action"] = "move";
                    a["to_index"] = 1;
                    ExpectFailUnchanged(run, "sld-e-move-no-id-fail", a, deck, 3, "slide_id");
                }
            }

            if (Want(filters, "sld-e-move-empty-id-fail"))
            {
                if (!deck.EnsureCount(3, out string err))
                {
                    run.CaseFail("sld-e-move-empty-id-fail", err);
                }
                else
                {
                    ExpectFailUnchanged(run, "sld-e-move-empty-id-fail", Move(deck, "", 1), deck, 3, "slide_id");
                }
            }

            if (Want(filters, "sld-e-move-to-str"))
            {
                if (!deck.EnsureCount(3, out string err))
                {
                    run.CaseFail("sld-e-move-to-str", err);
                }
                else
                {
                    IList<string> abc = deck.Ids();
                    var a = deck.Args();
                    a["action"] = "move";
                    a["slide_id"] = abc[2];
                    a["to_index"] = "1";
                    ToolResult r = Invoke(a);
                    var fails = new List<string>();
                    if (!Ok(r, fails) || !OrderIs(Rows(r), abc[2], abc[0], abc[1]))
                    {
                        if (fails.Count == 0)
                        {
                            fails.Add("字符串 to_index 应能解");
                        }

                        Finish(run, "sld-e-move-to-str", fails);
                    }
                    else
                    {
                        Finish(run, "sld-e-move-to-str", fails);
                    }
                }
            }

            if (Want(filters, "sld-e-move-id-spaces"))
            {
                if (!deck.EnsureCount(3, out string err))
                {
                    run.CaseFail("sld-e-move-id-spaces", err);
                }
                else
                {
                    IList<string> abc = deck.Ids();
                    var a = deck.Args();
                    a["action"] = "move";
                    a["slide_id"] = "  " + abc[2] + "  ";
                    a["to_index"] = 1;
                    ToolResult r = Invoke(a);
                    var fails = new List<string>();
                    if (!Ok(r, fails) || !OrderIs(Rows(r), abc[2], abc[0], abc[1]))
                    {
                        if (fails.Count == 0)
                        {
                            fails.Add("应 trim slide_id");
                        }

                        Finish(run, "sld-e-move-id-spaces", fails);
                    }
                    else
                    {
                        Finish(run, "sld-e-move-id-spaces", fails);
                    }
                }
            }

            if (Want(filters, "sld-e-move-id-stable"))
            {
                if (!deck.EnsureCount(3, out string err))
                {
                    run.CaseFail("sld-e-move-id-stable", err);
                }
                else
                {
                    string aId = deck.Ids()[0];
                    ToolResult r = Invoke(Move(deck, aId, 3));
                    var fails = new List<string>();
                    if (!Ok(r, fails) || !string.Equals(GetString(r, "slide_id"), aId, StringComparison.Ordinal)
                        || IndexOf(Rows(r), aId) != 3)
                    {
                        if (fails.Count == 0)
                        {
                            fails.Add("move 后 id 应不变且在末");
                        }

                        Finish(run, "sld-e-move-id-stable", fails);
                    }
                    else
                    {
                        Finish(run, "sld-e-move-id-stable", fails);
                    }
                }
            }

            if (Want(filters, "sld-e-move-neg-fail"))
            {
                if (!deck.EnsureCount(3, out string err))
                {
                    run.CaseFail("sld-e-move-neg-fail", err);
                }
                else
                {
                    ExpectFailUnchanged(run, "sld-e-move-neg-fail", Move(deck, deck.Ids()[0], -1), deck, 3, "越界");
                }
            }

            if (Want(filters, "sld-e-move-4-to-end"))
            {
                if (!deck.EnsureCount(4, out string err))
                {
                    run.CaseFail("sld-e-move-4-to-end", err);
                }
                else
                {
                    IList<string> ids = deck.Ids();
                    ToolResult r = Invoke(Move(deck, ids[0], 4));
                    var fails = new List<string>();
                    if (!Ok(r, fails) || !OrderIs(Rows(r), ids[1], ids[2], ids[3], ids[0]))
                    {
                        if (fails.Count == 0)
                        {
                            fails.Add("4 页首页移到末应为 b,c,d,a");
                        }

                        Finish(run, "sld-e-move-4-to-end", fails);
                    }
                    else
                    {
                        Finish(run, "sld-e-move-4-to-end", fails);
                    }
                }
            }

            if (Want(filters, "sld-e-move-4-to-1"))
            {
                if (!deck.EnsureCount(4, out string err))
                {
                    run.CaseFail("sld-e-move-4-to-1", err);
                }
                else
                {
                    IList<string> ids = deck.Ids();
                    ToolResult r = Invoke(Move(deck, ids[3], 1));
                    var fails = new List<string>();
                    if (!Ok(r, fails) || !OrderIs(Rows(r), ids[3], ids[0], ids[1], ids[2]))
                    {
                        if (fails.Count == 0)
                        {
                            fails.Add("4 页末页移到 1 应为 d,a,b,c");
                        }

                        Finish(run, "sld-e-move-4-to-1", fails);
                    }
                    else
                    {
                        Finish(run, "sld-e-move-4-to-1", fails);
                    }
                }
            }

            if (Want(filters, "sld-e-move-action-echo"))
            {
                if (!deck.EnsureCount(3, out string err))
                {
                    run.CaseFail("sld-e-move-action-echo", err);
                }
                else
                {
                    ToolResult r = Invoke(Move(deck, deck.Ids()[0], 2));
                    var fails = new List<string>();
                    if (!Ok(r, fails) || !string.Equals(GetString(r, "action"), "move", StringComparison.Ordinal))
                    {
                        if (fails.Count == 0)
                        {
                            fails.Add("action 应为 move");
                        }

                        Finish(run, "sld-e-move-action-echo", fails);
                    }
                    else
                    {
                        Finish(run, "sld-e-move-action-echo", fails);
                    }
                }
            }
        }

        private static void AbcMove(
            TestRun run,
            string id,
            Deck deck,
            Func<IList<string>, Dictionary<string, object>> make,
            Func<ToolResult, IList<string>, bool> order)
        {
            if (!deck.EnsureCount(3, out string err))
            {
                run.CaseFail(id, err);
                return;
            }

            IList<string> abc = deck.Ids();
            ToolResult r = Invoke(make(abc));
            var fails = new List<string>();
            if (!Ok(r, fails) || !Align(r, deck, fails) || !order(r, abc))
            {
                if (fails.Count == 0)
                {
                    fails.Add("顺序不对");
                }

                Finish(run, id, fails);
                return;
            }

            Finish(run, id, fails);
        }

        private static void RunExtraReorder(TestRun run, IList<string> filters, Deck deck)
        {
            if (!Want(filters, "sld-e-reorder-"))
            {
                return;
            }

            if (Want(filters, "sld-e-reorder-rot-left"))
            {
                AbcReorder(run, "sld-e-reorder-rot-left", deck, ids => new[] { ids[1], ids[2], ids[0] });
            }

            if (Want(filters, "sld-e-reorder-rot-right"))
            {
                AbcReorder(run, "sld-e-reorder-rot-right", deck, ids => new[] { ids[2], ids[0], ids[1] });
            }

            if (Want(filters, "sld-e-reorder-4"))
            {
                if (!deck.EnsureCount(4, out string err))
                {
                    run.CaseFail("sld-e-reorder-4", err);
                }
                else
                {
                    IList<string> ids = deck.Ids();
                    string[] order = { ids[1], ids[3], ids[0], ids[2] };
                    var a = deck.Args();
                    a["action"] = "reorder";
                    a["order"] = order;
                    ToolResult r = Invoke(a);
                    var fails = new List<string>();
                    if (!Ok(r, fails) || !OrderIs(Rows(r), order))
                    {
                        if (fails.Count == 0)
                        {
                            fails.Add("4 页自定义序不对");
                        }

                        Finish(run, "sld-e-reorder-4", fails);
                    }
                    else
                    {
                        Finish(run, "sld-e-reorder-4", fails);
                    }
                }
            }

            if (Want(filters, "sld-e-reorder-empty-fail"))
            {
                if (!deck.EnsureCount(3, out string err))
                {
                    run.CaseFail("sld-e-reorder-empty-fail", err);
                }
                else
                {
                    var a = deck.Args();
                    a["action"] = "reorder";
                    a["order"] = new string[0];
                    ExpectFailUnchanged(run, "sld-e-reorder-empty-fail", a, deck, 3, "order");
                }
            }

            if (Want(filters, "sld-e-reorder-long-fail"))
            {
                if (!deck.EnsureCount(3, out string err))
                {
                    run.CaseFail("sld-e-reorder-long-fail", err);
                }
                else
                {
                    IList<string> abc = deck.Ids();
                    var a = deck.Args();
                    a["action"] = "reorder";
                    a["order"] = new[] { abc[0], abc[1], abc[2], abc[0] };
                    ExpectFailUnchanged(run, "sld-e-reorder-long-fail", a, deck, 3, "页数");
                }
            }

            if (Want(filters, "sld-e-reorder-spaces"))
            {
                if (!deck.EnsureCount(3, out string err))
                {
                    run.CaseFail("sld-e-reorder-spaces", err);
                }
                else
                {
                    IList<string> abc = deck.Ids();
                    var a = deck.Args();
                    a["action"] = "reorder";
                    a["order"] = new[] { "  " + abc[2] + "  ", abc[1], abc[0] };
                    ToolResult r = Invoke(a);
                    var fails = new List<string>();
                    if (!Ok(r, fails) || !OrderIs(Rows(r), abc[2], abc[1], abc[0]))
                    {
                        if (fails.Count == 0)
                        {
                            fails.Add("order 项应 trim");
                        }

                        Finish(run, "sld-e-reorder-spaces", fails);
                    }
                    else
                    {
                        Finish(run, "sld-e-reorder-spaces", fails);
                    }
                }
            }

            if (Want(filters, "sld-e-reorder-2-swap"))
            {
                if (!deck.EnsureCount(2, out string err))
                {
                    run.CaseFail("sld-e-reorder-2-swap", err);
                }
                else
                {
                    IList<string> ids = deck.Ids();
                    var a = deck.Args();
                    a["action"] = "reorder";
                    a["order"] = new[] { ids[1], ids[0] };
                    ToolResult r = Invoke(a);
                    var fails = new List<string>();
                    if (!Ok(r, fails) || !OrderIs(Rows(r), ids[1], ids[0]))
                    {
                        if (fails.Count == 0)
                        {
                            fails.Add("两页对调失败");
                        }

                        Finish(run, "sld-e-reorder-2-swap", fails);
                    }
                    else
                    {
                        Finish(run, "sld-e-reorder-2-swap", fails);
                    }
                }
            }

            if (Want(filters, "sld-e-reorder-5-rev"))
            {
                if (!deck.EnsureCount(5, out string err))
                {
                    run.CaseFail("sld-e-reorder-5-rev", err);
                }
                else
                {
                    IList<string> ids = deck.Ids();
                    string[] order = { ids[4], ids[3], ids[2], ids[1], ids[0] };
                    var a = deck.Args();
                    a["action"] = "reorder";
                    a["order"] = order;
                    ToolResult r = Invoke(a);
                    var fails = new List<string>();
                    if (!Ok(r, fails) || !OrderIs(Rows(r), order))
                    {
                        if (fails.Count == 0)
                        {
                            fails.Add("5 页倒序失败");
                        }

                        Finish(run, "sld-e-reorder-5-rev", fails);
                    }
                    else
                    {
                        Finish(run, "sld-e-reorder-5-rev", fails);
                    }
                }
            }

            if (Want(filters, "sld-e-reorder-ids-stable"))
            {
                if (!deck.EnsureCount(3, out string err))
                {
                    run.CaseFail("sld-e-reorder-ids-stable", err);
                }
                else
                {
                    IList<string> abc = deck.Ids();
                    var a = deck.Args();
                    a["action"] = "reorder";
                    a["order"] = new[] { abc[2], abc[1], abc[0] };
                    ToolResult r = Invoke(a);
                    var fails = new List<string>();
                    if (!Ok(r, fails)
                        || IndexOf(Rows(r), abc[0]) < 1
                        || IndexOf(Rows(r), abc[1]) < 1
                        || IndexOf(Rows(r), abc[2]) < 1)
                    {
                        if (fails.Count == 0)
                        {
                            fails.Add("reorder 不得改 id");
                        }

                        Finish(run, "sld-e-reorder-ids-stable", fails);
                    }
                    else
                    {
                        Finish(run, "sld-e-reorder-ids-stable", fails);
                    }
                }
            }

            if (Want(filters, "sld-e-reorder-list"))
            {
                if (!deck.EnsureCount(3, out string err))
                {
                    run.CaseFail("sld-e-reorder-list", err);
                }
                else
                {
                    IList<string> abc = deck.Ids();
                    var a = deck.Args();
                    a["action"] = "reorder";
                    a["order"] = new List<string> { abc[2], abc[1], abc[0] };
                    ToolResult r = Invoke(a);
                    var fails = new List<string>();
                    if (!Ok(r, fails) || !OrderIs(Rows(r), abc[2], abc[1], abc[0]))
                    {
                        if (fails.Count == 0)
                        {
                            fails.Add("List order 应能解");
                        }

                        Finish(run, "sld-e-reorder-list", fails);
                    }
                    else
                    {
                        Finish(run, "sld-e-reorder-list", fails);
                    }
                }
            }

            if (Want(filters, "sld-e-reorder-after-move"))
            {
                if (!deck.EnsureCount(3, out string err))
                {
                    run.CaseFail("sld-e-reorder-after-move", err);
                }
                else
                {
                    IList<string> abc = deck.Ids();
                    Invoke(Move(deck, abc[0], 3));
                    var a = deck.Args();
                    a["action"] = "reorder";
                    a["order"] = new[] { abc[0], abc[1], abc[2] };
                    ToolResult r = Invoke(a);
                    var fails = new List<string>();
                    if (!Ok(r, fails) || !OrderIs(Rows(r), abc[0], abc[1], abc[2]))
                    {
                        if (fails.Count == 0)
                        {
                            fails.Add("move 后再 reorder 回原序失败");
                        }

                        Finish(run, "sld-e-reorder-after-move", fails);
                    }
                    else
                    {
                        Finish(run, "sld-e-reorder-after-move", fails);
                    }
                }
            }

            if (Want(filters, "sld-e-reorder-then-ident"))
            {
                if (!deck.EnsureCount(3, out string err))
                {
                    run.CaseFail("sld-e-reorder-then-ident", err);
                }
                else
                {
                    IList<string> abc = deck.Ids();
                    var rev = deck.Args();
                    rev["action"] = "reorder";
                    rev["order"] = new[] { abc[2], abc[1], abc[0] };
                    Invoke(rev);
                    var ident = deck.Args();
                    ident["action"] = "reorder";
                    ident["order"] = new[] { abc[2], abc[1], abc[0] };
                    ToolResult r = Invoke(ident);
                    var fails = new List<string>();
                    if (!Ok(r, fails) || !OrderIs(Rows(r), abc[2], abc[1], abc[0]))
                    {
                        if (fails.Count == 0)
                        {
                            fails.Add("再 ident 当前序应不变");
                        }

                        Finish(run, "sld-e-reorder-then-ident", fails);
                    }
                    else
                    {
                        Finish(run, "sld-e-reorder-then-ident", fails);
                    }
                }
            }
        }

        private static void AbcReorder(TestRun run, string id, Deck deck, Func<IList<string>, string[]> orderOf)
        {
            if (!deck.EnsureCount(3, out string err))
            {
                run.CaseFail(id, err);
                return;
            }

            IList<string> abc = deck.Ids();
            string[] order = orderOf(abc);
            var a = deck.Args();
            a["action"] = "reorder";
            a["order"] = order;
            ToolResult r = Invoke(a);
            var fails = new List<string>();
            if (!Ok(r, fails) || !OrderIs(Rows(r), order))
            {
                if (fails.Count == 0)
                {
                    fails.Add("序不对");
                }

                Finish(run, id, fails);
                return;
            }

            Finish(run, id, fails);
        }

        private static void RunExtraCross(TestRun run, IList<string> filters, Deck dest, Func<Deck> openSecond)
        {
            bool need = Want(filters, "sld-e-cross-");
            if (!need)
            {
                return;
            }

            Deck src;
            try
            {
                src = openSecond();
            }
            catch (Exception ex)
            {
                FailNeeded(
                    run,
                    filters,
                    ex.Message,
                    "sld-e-cross-to-1",
                    "sld-e-cross-to-last",
                    "sld-e-cross-to-plus",
                    "sld-e-cross-repeat",
                    "sld-e-cross-missing-fail",
                    "sld-e-cross-to-0-fail",
                    "sld-e-cross-over-fail",
                    "sld-e-cross-same-ch",
                    "sld-e-cross-src-content",
                    "sld-e-cross-dest-first",
                    "sld-e-cross-created-idx",
                    "sld-e-cross-empty-ids-fail");
                return;
            }

            if (!src.EnsureCount(2, out string err) || !dest.EnsureCount(2, out err))
            {
                FailNeeded(run, filters, err, "sld-e-cross-to-1", "sld-e-cross-to-last", "sld-e-cross-to-plus");
            }

            IList<string> sids = src.Ids();

            if (Want(filters, "sld-e-cross-to-1"))
            {
                ResetCross(dest, src);
                sids = src.Ids();
                string first = dest.Ids()[0];
                var a = CrossDup(dest, src, sids[0]);
                a["to_index"] = 1;
                ToolResult r = Invoke(a);
                var fails = new List<string>();
                if (!Ok(r, fails) || IndexOf(Rows(r), GetString(r, "slide_id")) != 1
                    || IndexOf(Rows(r), first) != 2 || src.Count() != 2)
                {
                    if (fails.Count == 0)
                    {
                        fails.Add("跨稿插 1 失败或源被改");
                    }

                    Finish(run, "sld-e-cross-to-1", fails);
                }
                else
                {
                    Finish(run, "sld-e-cross-to-1", fails);
                }
            }

            if (Want(filters, "sld-e-cross-to-last"))
            {
                ResetCross(dest, src);
                sids = src.Ids();
                string last = dest.Ids()[1];
                var a = CrossDup(dest, src, sids[0]);
                a["to_index"] = 2;
                ToolResult r = Invoke(a);
                var fails = new List<string>();
                if (!Ok(r, fails) || IndexOf(Rows(r), GetString(r, "slide_id")) != 2
                    || IndexOf(Rows(r), last) != 3)
                {
                    if (fails.Count == 0)
                    {
                        fails.Add("跨稿 to_index=destCount 新页应在 2");
                    }

                    Finish(run, "sld-e-cross-to-last", fails);
                }
                else
                {
                    Finish(run, "sld-e-cross-to-last", fails);
                }
            }

            if (Want(filters, "sld-e-cross-to-plus"))
            {
                ResetCross(dest, src);
                sids = src.Ids();
                var a = CrossDup(dest, src, sids[1]);
                a["to_index"] = 3;
                AssertNewAt(run, "sld-e-cross-to-plus", Invoke(a), dest, 3);
            }

            if (Want(filters, "sld-e-cross-repeat"))
            {
                ResetCross(dest, src);
                sids = src.Ids();
                int dc = dest.Count();
                var a = dest.Args();
                a["action"] = "duplicate";
                a["source_channel_id"] = src.Channel.ChannelId;
                a["slide_ids"] = new[] { sids[0], sids[0] };
                ToolResult r = Invoke(a);
                var fails = new List<string>();
                IList<CreatedRow> cr = Created(r);
                if (!Ok(r, fails) || dest.Count() != dc + 2 || cr.Count != 2
                    || cr[0].NewId == cr[1].NewId)
                {
                    if (fails.Count == 0)
                    {
                        fails.Add("跨稿重复源应 +2 且新 id 不同");
                    }

                    Finish(run, "sld-e-cross-repeat", fails);
                }
                else
                {
                    Finish(run, "sld-e-cross-repeat", fails);
                }
            }

            if (Want(filters, "sld-e-cross-missing-fail"))
            {
                ResetCross(dest, src);
                int dc = dest.Count();
                int sc = src.Count();
                var a = CrossDup(dest, src, "999999");
                ToolResult r = Invoke(a);
                var fails = new List<string>();
                if (r != null && r.Success)
                {
                    fails.Add("应失败");
                }
                else if ((r == null ? "" : r.Error ?? "").IndexOf("不存在", StringComparison.Ordinal) < 0)
                {
                    fails.Add("文案应含不存在");
                }

                if (dest.Count() != dc || src.Count() != sc)
                {
                    fails.Add("页数被改了");
                }

                Finish(run, "sld-e-cross-missing-fail", fails);
            }

            if (Want(filters, "sld-e-cross-to-0-fail") || Want(filters, "sld-e-cross-over-fail")
                || Want(filters, "sld-e-cross-empty-ids-fail"))
            {
                ResetCross(dest, src);
                sids = src.Ids();
                int dc = dest.Count();
                if (Want(filters, "sld-e-cross-to-0-fail"))
                {
                    var a = CrossDup(dest, src, sids[0]);
                    a["to_index"] = 0;
                    ExpectFailUnchanged(run, "sld-e-cross-to-0-fail", a, dest, dc, "越界");
                }

                if (Want(filters, "sld-e-cross-over-fail"))
                {
                    var a = CrossDup(dest, src, sids[0]);
                    a["to_index"] = dest.Count() + 2;
                    ExpectFailUnchanged(run, "sld-e-cross-over-fail", a, dest, dest.Count(), "越界");
                }

                if (Want(filters, "sld-e-cross-empty-ids-fail"))
                {
                    var a = dest.Args();
                    a["action"] = "duplicate";
                    a["source_channel_id"] = src.Channel.ChannelId;
                    a["slide_ids"] = new string[0];
                    ExpectFailUnchanged(run, "sld-e-cross-empty-ids-fail", a, dest, dest.Count(), "不能为空");
                }
            }

            if (Want(filters, "sld-e-cross-same-ch"))
            {
                if (!dest.EnsureCount(3, out err))
                {
                    run.CaseFail("sld-e-cross-same-ch", err);
                }
                else
                {
                    string aId = dest.Ids()[0];
                    var a = Dup(dest, aId);
                    a["source_channel_id"] = dest.Channel.ChannelId;
                    ToolResult r = Invoke(a);
                    var fails = new List<string>();
                    if (!Ok(r, fails) || Created(r).Count != 1 || IndexOf(Rows(r), aId) < 1)
                    {
                        if (fails.Count == 0)
                        {
                            fails.Add("source=自己应等同同稿拷");
                        }

                        Finish(run, "sld-e-cross-same-ch", fails);
                    }
                    else
                    {
                        Finish(run, "sld-e-cross-same-ch", fails);
                    }
                }
            }

            if (Want(filters, "sld-e-cross-src-content"))
            {
                ResetCross(dest, src);
                sids = src.Ids();
                if (!src.TryApplyText(sids[0], "跨稿源", out err))
                {
                    run.CaseFail("sld-e-cross-src-content", err ?? "源页铺字失败");
                }
                else
                {
                    ToolResult r = Invoke(CrossDup(dest, src, sids[0]));
                    var fails = new List<string>();
                    string nid = GetString(r, "slide_id");
                    if (!Ok(r, fails) || !dest.PageHasText(nid, "跨稿源", out err)
                        || !src.PageHasText(sids[0], "跨稿源", out _))
                    {
                        if (fails.Count == 0)
                        {
                            fails.Add(err ?? "目标新页或源页丢了「跨稿源」");
                        }

                        Finish(run, "sld-e-cross-src-content", fails);
                    }
                    else
                    {
                        Finish(run, "sld-e-cross-src-content", fails);
                    }
                }
            }

            if (Want(filters, "sld-e-cross-dest-first"))
            {
                ResetCross(dest, src);
                sids = src.Ids();
                string first = dest.Ids()[0];
                ToolResult r = Invoke(CrossDup(dest, src, sids[0]));
                var fails = new List<string>();
                if (!Ok(r, fails) || IndexOf(Rows(r), first) != 1)
                {
                    if (fails.Count == 0)
                    {
                        fails.Add("默认尾部拷不应改目标首页 id");
                    }

                    Finish(run, "sld-e-cross-dest-first", fails);
                }
                else
                {
                    Finish(run, "sld-e-cross-dest-first", fails);
                }
            }

            if (Want(filters, "sld-e-cross-created-idx"))
            {
                ResetCross(dest, src);
                sids = src.Ids();
                ToolResult r = Invoke(CrossDup(dest, src, sids[0]));
                var fails = new List<string>();
                IList<CreatedRow> cr = Created(r);
                string nid = GetString(r, "slide_id");
                if (!Ok(r, fails) || cr.Count != 1 || cr[0].Index != IndexOf(Rows(r), nid))
                {
                    if (fails.Count == 0)
                    {
                        fails.Add("跨稿 created.index 应等于新页 index");
                    }

                    Finish(run, "sld-e-cross-created-idx", fails);
                }
                else
                {
                    Finish(run, "sld-e-cross-created-idx", fails);
                }
            }
        }

        private static void RunExtraHiddenChain(TestRun run, IList<string> filters, Deck deck)
        {
            if (Want(filters, "sld-e-hid-listed")
                || Want(filters, "sld-e-hid-dup-flag")
                || Want(filters, "sld-e-hid-move"))
            {
                if (!deck.EnsureCount(3, out string err) || !deck.TryHide(2, true, out err))
                {
                    FailNeeded(run, filters, err ?? "隐藏失败", "sld-e-hid-listed", "sld-e-hid-dup-flag", "sld-e-hid-move");
                }
                else
                {
                    IList<string> abc = deck.Ids();
                    string hid = abc[1];
                    if (Want(filters, "sld-e-hid-listed"))
                    {
                        ToolResult r = Invoke(Move(deck, abc[0], 1));
                        var fails = new List<string>();
                        bool found = false;
                        IList<SlideRow> rows = Rows(r);
                        for (int i = 0; i < rows.Count; i++)
                        {
                            if (string.Equals(rows[i].SlideId, hid, StringComparison.Ordinal))
                            {
                                found = rows[i].Hidden;
                                break;
                            }
                        }

                        if (!Ok(r, fails) || !found)
                        {
                            if (fails.Count == 0)
                            {
                                fails.Add("slides[].hidden 应为 true");
                            }

                            Finish(run, "sld-e-hid-listed", fails);
                        }
                        else
                        {
                            Finish(run, "sld-e-hid-listed", fails);
                        }
                    }

                    if (Want(filters, "sld-e-hid-dup-flag"))
                    {
                        if (!deck.EnsureCount(3, out err) || !deck.TryHide(2, true, out err))
                        {
                            run.CaseFail("sld-e-hid-dup-flag", err ?? "隐藏失败");
                        }
                        else
                        {
                            hid = deck.Ids()[1];
                            ToolResult r = Invoke(Dup(deck, hid));
                            var fails = new List<string>();
                            string nid = GetString(r, "slide_id");
                            bool hiddenNew = false;
                            IList<SlideRow> rows = Rows(r);
                            for (int i = 0; i < rows.Count; i++)
                            {
                                if (string.Equals(rows[i].SlideId, nid, StringComparison.Ordinal))
                                {
                                    hiddenNew = rows[i].Hidden;
                                    break;
                                }
                            }

                            if (!Ok(r, fails) || !hiddenNew)
                            {
                                if (fails.Count == 0)
                                {
                                    fails.Add("拷隐藏页后新页也应 hidden");
                                }

                                Finish(run, "sld-e-hid-dup-flag", fails);
                            }
                            else
                            {
                                Finish(run, "sld-e-hid-dup-flag", fails);
                            }
                        }
                    }

                    if (Want(filters, "sld-e-hid-move"))
                    {
                        if (!deck.EnsureCount(3, out err) || !deck.TryHide(2, true, out err))
                        {
                            run.CaseFail("sld-e-hid-move", err ?? "隐藏失败");
                        }
                        else
                        {
                            hid = deck.Ids()[1];
                            ToolResult r = Invoke(Move(deck, hid, 1));
                            var fails = new List<string>();
                            bool still = false;
                            IList<SlideRow> rows = Rows(r);
                            for (int i = 0; i < rows.Count; i++)
                            {
                                if (string.Equals(rows[i].SlideId, hid, StringComparison.Ordinal) && rows[i].Index == 1)
                                {
                                    still = rows[i].Hidden;
                                    break;
                                }
                            }

                            if (!Ok(r, fails) || !still)
                            {
                                if (fails.Count == 0)
                                {
                                    fails.Add("隐藏页移到 1 后仍应 hidden");
                                }

                                Finish(run, "sld-e-hid-move", fails);
                            }
                            else
                            {
                                Finish(run, "sld-e-hid-move", fails);
                            }
                        }
                    }
                }
            }

            if (Want(filters, "sld-e-chain-ops"))
            {
                if (!deck.EnsureCount(2, out string err))
                {
                    run.CaseFail("sld-e-chain-ops", err);
                }
                else
                {
                    var fails = new List<string>();
                    var add = deck.Args();
                    add["action"] = "add";
                    ToolResult r = Invoke(add);
                    if (!Ok(r, fails))
                    {
                        Finish(run, "sld-e-chain-ops", fails);
                        return;
                    }

                    r = Invoke(Dup(deck, deck.Ids()[0]));
                    if (!Ok(r, fails))
                    {
                        Finish(run, "sld-e-chain-ops", fails);
                        return;
                    }

                    r = Invoke(Move(deck, GetString(r, "slide_id"), 1));
                    if (!Ok(r, fails))
                    {
                        Finish(run, "sld-e-chain-ops", fails);
                        return;
                    }

                    IList<string> now = deck.Ids();
                    var ord = deck.Args();
                    ord["action"] = "reorder";
                    ord["order"] = new[] { now[now.Count - 1], now[0], now[1], now[2] };
                    if (now.Count != 4)
                    {
                        fails.Add("链式后应为 4 页，实际 " + now.Count);
                        Finish(run, "sld-e-chain-ops", fails);
                        return;
                    }

                    r = Invoke(ord);
                    if (!Ok(r, fails))
                    {
                        Finish(run, "sld-e-chain-ops", fails);
                        return;
                    }

                    r = Invoke(Del(deck, deck.Ids()[1]));
                    if (!Ok(r, fails) || !Align(r, deck, fails) || GetInt(r, "slide_count") != 3)
                    {
                        if (fails.Count == 0)
                        {
                            fails.Add("链式最后删一页应为 3");
                        }

                        Finish(run, "sld-e-chain-ops", fails);
                        return;
                    }

                    Finish(run, "sld-e-chain-ops", fails);
                }
            }
        }

        private static void AssertNewAt(TestRun run, string id, ToolResult r, Deck deck, int index)
        {
            var fails = new List<string>();
            if (!Ok(r, fails) || !Align(r, deck, fails))
            {
                Finish(run, id, fails);
                return;
            }

            if (IndexOf(Rows(r), GetString(r, "slide_id")) != index)
            {
                fails.Add("新页应在 index=" + index);
            }

            Finish(run, id, fails);
        }

        private static void AssertDeleted(TestRun run, string id, ToolResult r, Deck deck, string gone, int expectCount)
        {
            var fails = new List<string>();
            if (!Ok(r, fails) || !Align(r, deck, fails))
            {
                Finish(run, id, fails);
                return;
            }

            if (GetInt(r, "slide_count") != expectCount)
            {
                fails.Add("应为 " + expectCount + " 页");
            }

            if (gone != null && (IndexOf(Rows(r), gone) >= 0 || !DeletedHas(r, gone)))
            {
                fails.Add("列表应无 " + gone);
            }

            Finish(run, id, fails);
        }

        private static Dictionary<string, object> CrossDup(Deck dest, Deck src, string slideId)
        {
            var a = dest.Args();
            a["action"] = "duplicate";
            a["slide_id"] = slideId;
            a["source_channel_id"] = src.Channel.ChannelId;
            return a;
        }

        private static void ResetCross(Deck dest, Deck src)
        {
            dest.EnsureCount(2, out _);
            src.EnsureCount(2, out _);
        }
    }
}