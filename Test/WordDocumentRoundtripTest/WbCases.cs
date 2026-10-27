using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using WordAddIn1;

namespace WordDocumentRoundtripTest
{
    /// <summary>
    /// 一次调用改多个格式。铺稿用 COM，抄当次 S_ / P_。
    /// </summary>
    internal static class WbCases
    {
        public const int Target = 37;

        private const string Jia = "易写多格式甲句。";
        private const string Yi = "易写多格式乙句。";
        private const string Bing = "易写多格式丙句。";
        private const string Ding = "易写多格式丁句。";
        private const string Wu = "易写多格式戊句。";
        private const string Body = Jia + Yi + "\r" + Bing + "\r" + Ding;
        private const string Before = "易写多格式表前句。";
        private const string Left = "易写多格式左格。";
        private const string Right = "易写多格式右格。";
        private const string Dup = "易写多格式重复段。";
        private const string Mid = "易写多格式夹心句。";

        public static void Run(
            TestRun run,
            string host,
            int wordBefore,
            int wpsBefore,
            bool wordWinBefore,
            bool wpsWinBefore)
        {
            var s = new WdSession(host);
            RunSentence(run, s);
            RunParagraphChar(run, s);
            RunParagraphLayout(run, s);
            RunCross(run, s);
            RunContract(run, s);
            RunSnapshot(run, s);
            RunNoQuit(run, wordBefore, wpsBefore, wordWinBefore, wpsWinBefore);
        }

        private static void RunSentence(TestRun run, WdSession s)
        {
            using (Planted doc = Plant(s, Body))
            {
                if (doc == null || !doc.HasBody(run, "wb-s-one-flat"))
                {
                    FailRest(run, "铺稿甲失败", "wb-s-one-flat");
                    return;
                }

                string sa = doc.Sentence(Jia);
                string sb = doc.Sentence(Yi);
                string sc = doc.Sentence(Bing);
                string sd = doc.Sentence(Ding);
                ToolResult flat = s.ApplyDocumentFormat(CharFlat(doc.Id, sa, Bold()));
                Ok(run, "wb-s-one-flat",
                    flat.Success && LegacyHas(flat, sa) && BoldOf(s, doc, sa) && !BoldOf(s, doc, sb) && !BoldOf(s, doc, sc) && !BoldOf(s, doc, sd) && doc.TextsRemain(s),
                    flat.Success ? "甲未只加粗，或回包不是旧形状" : (flat.Error ?? ""));
            }

            using (Planted doc = Plant(s, Body))
            {
                if (doc == null || !doc.HasBody(run, "wb-s-one-item"))
                {
                    FailRest(run, "铺稿甲失败", "wb-s-one-item");
                    return;
                }

                string sa = doc.Sentence(Jia);
                ToolResult one = s.ApplyDocumentFormat(CharItems(doc.Id, Item("target_codes", sa, "char_format", Bold())));
                Ok(run, "wb-s-one-item",
                    one.Success && RunsAre(one, sa) && BoldOf(s, doc, sa) && !BoldOf(s, doc, doc.Sentence(Yi)) && !BoldOf(s, doc, doc.Sentence(Bing)) && !BoldOf(s, doc, doc.Sentence(Ding)) && doc.TextsRemain(s),
                    one.Success ? "单项未只加粗甲" : (one.Error ?? ""));
            }

            using (Planted doc = Plant(s, Body))
            {
                if (doc == null || !doc.HasBody(run, "wb-s-multi"))
                {
                    FailRest(run, "铺稿甲失败", "wb-s-multi");
                    return;
                }

                string sa = doc.Sentence(Jia);
                string sb = doc.Sentence(Yi);
                string sc = doc.Sentence(Bing);
                string sd = doc.Sentence(Ding);
                ToolResult multi = s.ApplyDocumentFormat(CharItems(
                    doc.Id,
                    Item("target_codes", sa, "char_format", Bold()),
                    Item("target_codes", sb, "char_format", Under()),
                    Item("target_codes", sc, "char_format", Strike())));
                Ok(run, "wb-s-multi",
                    multi.Success
                        && RunsAre(multi, sa, sb, sc)
                        && IndexesAre(multi, 0, 1, 2)
                        && BoldOf(s, doc, sa) && !UnderOf(s, doc, sa) && !StrikeOf(s, doc, sa)
                        && UnderOf(s, doc, sb) && !BoldOf(s, doc, sb) && !StrikeOf(s, doc, sb)
                        && StrikeOf(s, doc, sc) && !BoldOf(s, doc, sc) && !UnderOf(s, doc, sc)
                        && !BoldOf(s, doc, sd) && !UnderOf(s, doc, sd) && !StrikeOf(s, doc, sd)
                        && doc.TextsRemain(s),
                    multi.Success ? "三句格式串了" : (multi.Error ?? ""));
            }

            using (Planted doc = Plant(s, Body))
            {
                if (doc == null || !doc.HasBody(run, "wb-s-gap"))
                {
                    FailRest(run, "铺稿甲失败", "wb-s-gap", "wb-s-gap-flat", "wb-s-reverse", "wb-s-multi-gap");
                    return;
                }

                string sa = doc.Sentence(Jia);
                string sb = doc.Sentence(Yi);
                string sc = doc.Sentence(Bing);
                string sd = doc.Sentence(Ding);
                ToolResult gap = s.ApplyDocumentFormat(CharItems(doc.Id, Item("target_codes", sa + "," + sc, "char_format", Bold())));
                Ok(run, "wb-s-gap",
                    gap.Success && RunsAre(gap, sa, sc) && BoldOf(s, doc, sa) && BoldOf(s, doc, sc) && !BoldOf(s, doc, sb) && !BoldOf(s, doc, sd),
                    gap.Success ? "跳句未拆开或乙被加粗" : (gap.Error ?? ""));
            }

            using (Planted doc = Plant(s, Body))
            {
                if (doc == null || !doc.HasBody(run, "wb-s-gap-flat"))
                {
                    FailRest(run, "铺稿甲失败", "wb-s-gap-flat");
                    return;
                }

                string sa = doc.Sentence(Jia);
                string sb = doc.Sentence(Yi);
                string sc = doc.Sentence(Bing);
                ToolResult gap = s.ApplyDocumentFormat(CharFlat(doc.Id, sa + "," + sc, Bold()));
                Ok(run, "wb-s-gap-flat",
                    gap.Success && RunsAre(gap, sa, sc) && BoldOf(s, doc, sa) && BoldOf(s, doc, sc) && !BoldOf(s, doc, sb),
                    gap.Success ? "扁平跳句未拆开" : (gap.Error ?? ""));
            }

            using (Planted doc = Plant(s, Body))
            {
                if (doc == null || !doc.HasBody(run, "wb-s-reverse"))
                {
                    FailRest(run, "铺稿甲失败", "wb-s-reverse");
                    return;
                }

                string sa = doc.Sentence(Jia);
                string sb = doc.Sentence(Yi);
                string sc = doc.Sentence(Bing);
                ToolResult rev = s.ApplyDocumentFormat(CharItems(doc.Id, Item("target_codes", sb + "," + sa, "char_format", Bold())));
                Ok(run, "wb-s-reverse",
                    rev.Success && RunsAre(rev, sb, sa) && BoldOf(s, doc, sa) && BoldOf(s, doc, sb) && !BoldOf(s, doc, sc) && !BoldOf(s, doc, doc.Sentence(Ding)),
                    rev.Success ? "写反没有拆成两个单码" : (rev.Error ?? ""));
            }

            RunCrossSeg(run, s);

            using (Planted doc = Plant(s, Body))
            {
                if (doc == null || !doc.HasBody(run, "wb-s-multi-gap"))
                {
                    FailRest(run, "铺稿甲失败", "wb-s-multi-gap");
                    return;
                }

                string sa = doc.Sentence(Jia);
                string sb = doc.Sentence(Yi);
                string sc = doc.Sentence(Bing);
                string sd = doc.Sentence(Ding);
                ToolResult mix = s.ApplyDocumentFormat(CharItems(
                    doc.Id,
                    Item("target_codes", sa + "," + sd, "char_format", Bold()),
                    Item("target_codes", sb, "char_format", Under())));
                Ok(run, "wb-s-multi-gap",
                    mix.Success
                        && BoldOf(s, doc, sa) && !UnderOf(s, doc, sa)
                        && UnderOf(s, doc, sb) && !BoldOf(s, doc, sb)
                        && !BoldOf(s, doc, sc) && !UnderOf(s, doc, sc)
                        && BoldOf(s, doc, sd) && !UnderOf(s, doc, sd),
                    mix.Success ? "多项跳句格式不对" : (mix.Error ?? ""));
            }
        }

        private static void RunCrossSeg(TestRun run, WdSession s)
        {
            string path = s.NewDocxPath();
            ToolResult opened = s.Open(path, true, s.App);
            if (!opened.Success || !s.TypePlain(path, Before) || !s.InsertTable(path, 1, 2, new[] { Left, Right }))
            {
                Ok(run, "wb-s-cross-seg", false, opened.Error ?? "稿乙铺失败");
                if (opened.Success)
                {
                    s.CloseDocByPath(path);
                }

                return;
            }

            string id = WdSession.Str(opened, "channel_id");
            using (var doc = Planted.FromReads(s, path, id, Before, Left, Right))
            {
                string before = doc.Sentence(Before);
                string left = doc.Sentence(Left);
                string right = doc.Sentence(Right);
                if (string.IsNullOrEmpty(before) || string.IsNullOrEmpty(left) || string.IsNullOrEmpty(right))
                {
                    Ok(run, "wb-s-cross-seg", false, "稿乙抄不到三句");
                    return;
                }

                ToolResult apply = s.ApplyDocumentFormat(CharItems(
                    id,
                    Item("target_codes", before + "," + right, "char_format", Bold())));
                bool split = apply.Success && AppliedCount(apply) >= 2;
                Ok(run, "wb-s-cross-seg",
                    split && BoldOf(s, doc, before) && BoldOf(s, doc, right) && !BoldOf(s, doc, left),
                    apply.Success ? "跨 seg 未拆开或左格被加粗" : (apply.Error ?? ""));
            }
        }

        private static void RunParagraphChar(TestRun run, WdSession s)
        {
            using (Planted doc = Plant(s, Body))
            {
                if (doc == null || !doc.HasBody(run, "wb-p-one-flat"))
                {
                    FailRest(run, "铺稿甲失败", "wb-p-one-flat");
                    return;
                }

                string pa = doc.Paragraph(Jia);
                ToolResult flat = s.ApplyDocumentFormat(ParaCharFlat(doc.Id, pa, Under()));
                Ok(run, "wb-p-one-flat",
                    flat.Success && LegacyHas(flat, pa) && UnderOf(s, doc, doc.Sentence(Jia)) && UnderOf(s, doc, doc.Sentence(Yi)) && !UnderOf(s, doc, doc.Sentence(Bing)) && !UnderOf(s, doc, doc.Sentence(Ding)),
                    flat.Success ? "段甲下划线未按段套上" : (flat.Error ?? ""));
            }

            using (Planted doc = Plant(s, Body))
            {
                if (doc == null || !doc.HasBody(run, "wb-p-one-item"))
                {
                    FailRest(run, "铺稿甲失败", "wb-p-one-item");
                    return;
                }

                string pa = doc.Paragraph(Jia);
                ToolResult one = s.ApplyDocumentFormat(CharItems(doc.Id, Item("target_paragraph_codes", pa, "char_format", Under())));
                Ok(run, "wb-p-one-item",
                    one.Success && RunsAre(one, pa) && UnderOf(s, doc, doc.Sentence(Jia)) && !UnderOf(s, doc, doc.Sentence(Bing)),
                    one.Success ? "单项段下划线不对" : (one.Error ?? ""));
            }

            using (Planted doc = Plant(s, Body))
            {
                if (doc == null || !doc.HasBody(run, "wb-p-multi"))
                {
                    FailRest(run, "铺稿甲失败", "wb-p-multi");
                    return;
                }

                string pa = doc.Paragraph(Jia);
                string pc = doc.Paragraph(Ding);
                ToolResult multi = s.ApplyDocumentFormat(CharItems(
                    doc.Id,
                    Item("target_paragraph_codes", pa, "char_format", Bold()),
                    Item("target_paragraph_codes", pc, "char_format", Strike())));
                Ok(run, "wb-p-multi",
                    multi.Success
                        && BoldOf(s, doc, doc.Sentence(Jia)) && BoldOf(s, doc, doc.Sentence(Yi)) && !StrikeOf(s, doc, doc.Sentence(Jia))
                        && StrikeOf(s, doc, doc.Sentence(Ding)) && !BoldOf(s, doc, doc.Sentence(Ding))
                        && !BoldOf(s, doc, doc.Sentence(Bing)) && !StrikeOf(s, doc, doc.Sentence(Bing)),
                    multi.Success ? "两段字符格式串了" : (multi.Error ?? ""));
            }

            using (Planted doc = Plant(s, Body))
            {
                if (doc == null || !doc.HasBody(run, "wb-p-gap"))
                {
                    FailRest(run, "铺稿甲失败", "wb-p-gap");
                    return;
                }

                string pa = doc.Paragraph(Jia);
                string pc = doc.Paragraph(Ding);
                ToolResult gap = s.ApplyDocumentFormat(CharItems(doc.Id, Item("target_paragraph_codes", pa + "," + pc, "char_format", Under())));
                Ok(run, "wb-p-gap",
                    gap.Success && RunsAre(gap, pa, pc) && UnderOf(s, doc, doc.Sentence(Jia)) && UnderOf(s, doc, doc.Sentence(Ding)) && !UnderOf(s, doc, doc.Sentence(Bing)),
                    gap.Success ? "跳段未拆开或段乙有下划线" : (gap.Error ?? ""));
            }

            using (Planted doc = Plant(s, Body))
            {
                if (doc == null || !doc.HasBody(run, "wb-p-reverse"))
                {
                    FailRest(run, "铺稿甲失败", "wb-p-reverse");
                    return;
                }

                string pa = doc.Paragraph(Jia);
                string pc = doc.Paragraph(Ding);
                ToolResult rev = s.ApplyDocumentFormat(CharItems(doc.Id, Item("target_paragraph_codes", pc + "," + pa, "char_format", Under())));
                Ok(run, "wb-p-reverse",
                    rev.Success && RunsAre(rev, pc, pa) && UnderOf(s, doc, doc.Sentence(Jia)) && UnderOf(s, doc, doc.Sentence(Ding)) && !UnderOf(s, doc, doc.Sentence(Bing)),
                    rev.Success ? "段码写反没有拆开" : (rev.Error ?? ""));
            }

            using (Planted doc = Plant(s, Dup + "\r" + Mid + "\r" + Dup))
            {
                if (doc == null)
                {
                    Ok(run, "wb-p-collision", false, "稿丙铺失败");
                    return;
                }

                string first = doc.NthParagraphContaining(Dup, 0);
                string second = doc.NthParagraphContaining(Dup, 1);
                if (string.IsNullOrEmpty(first) || string.IsNullOrEmpty(second) || first.IndexOf("_", 2, StringComparison.Ordinal) < 0)
                {
                    Ok(run, "wb-p-collision", false, "没有撞码公开码 " + first + " / " + second);
                    return;
                }

                ToolResult apply = s.ApplyDocumentFormat(CharItems(doc.Id, Item("target_paragraph_codes", first, "char_format", Under())));
                bool firstUl = false;
                bool midUl = false;
                bool secondUl = false;
                bool read = s.TryNthParagraphUnderline(doc.Path, Dup, 0, out firstUl)
                    && s.TryNthParagraphUnderline(doc.Path, Mid, 0, out midUl)
                    && s.TryNthParagraphUnderline(doc.Path, Dup, 1, out secondUl);
                Ok(run, "wb-p-collision",
                    apply.Success && read && firstUl && !midUl && !secondUl && doc.TextsRemain(s, Dup, Mid),
                    apply.Success
                        ? "下划线 first=" + firstUl + " mid=" + midUl + " second=" + secondUl + " codes=" + first + "," + second
                        : (apply.Error ?? ""));
            }
        }

        private static void RunParagraphLayout(TestRun run, WdSession s)
        {
            using (Planted doc = Plant(s, Body))
            {
                if (doc == null || !doc.HasBody(run, "wb-para-one-flat"))
                {
                    FailRest(run, "铺稿甲失败", "wb-para-one-flat");
                    return;
                }

                string pb = doc.Paragraph(Bing);
                ToolResult flat = s.ApplyParagraphFormat(ParaFlat(doc.Id, pb, Align("center")));
                Ok(run, "wb-para-one-flat",
                    flat.Success && LegacyParaHas(flat, pb) && IsAlign(s, doc, pb, "center") && IsAlign(s, doc, doc.Paragraph(Jia), "left") && IsAlign(s, doc, doc.Paragraph(Ding), "left"),
                    flat.Success ? "段乙未居中或回包没有 affected_paragraph_codes" : (flat.Error ?? ""));
            }

            using (Planted doc = Plant(s, Body))
            {
                if (doc == null || !doc.HasBody(run, "wb-para-one-item"))
                {
                    FailRest(run, "铺稿甲失败", "wb-para-one-item");
                    return;
                }

                string pb = doc.Paragraph(Bing);
                ToolResult one = s.ApplyParagraphFormat(ParaItems(doc.Id, ParaItem(pb, Align("center"))));
                Ok(run, "wb-para-one-item",
                    one.Success && RunsAre(one, pb) && IsAlign(s, doc, pb, "center") && IsAlign(s, doc, doc.Paragraph(Jia), "left"),
                    one.Success ? "单项居中不对" : (one.Error ?? ""));
            }

            using (Planted doc = Plant(s, Body))
            {
                if (doc == null || !doc.HasBody(run, "wb-para-multi"))
                {
                    FailRest(run, "铺稿甲失败", "wb-para-multi");
                    return;
                }

                string pa = doc.Paragraph(Jia);
                string pb = doc.Paragraph(Bing);
                string pc = doc.Paragraph(Ding);
                ToolResult multi = s.ApplyParagraphFormat(ParaItems(doc.Id, ParaItem(pa, Align("center")), ParaItem(pc, Align("right"))));
                Ok(run, "wb-para-multi",
                    multi.Success && IsAlign(s, doc, pa, "center") && IsAlign(s, doc, pc, "right") && IsAlign(s, doc, pb, "left"),
                    multi.Success ? "多段对齐串了" : (multi.Error ?? ""));
            }

            using (Planted doc = Plant(s, Body))
            {
                if (doc == null || !doc.HasBody(run, "wb-para-gap"))
                {
                    FailRest(run, "铺稿甲失败", "wb-para-gap");
                    return;
                }

                string pa = doc.Paragraph(Jia);
                string pb = doc.Paragraph(Bing);
                string pc = doc.Paragraph(Ding);
                ToolResult gap = s.ApplyParagraphFormat(ParaItems(doc.Id, ParaItem(pa + "," + pc, Align("center"))));
                Ok(run, "wb-para-gap",
                    gap.Success && RunsAre(gap, pa, pc) && IsAlign(s, doc, pa, "center") && IsAlign(s, doc, pc, "center") && IsAlign(s, doc, pb, "left"),
                    gap.Success ? "跳段版式未拆开" : (gap.Error ?? ""));
            }

            using (Planted doc = Plant(s, Body))
            {
                if (doc == null || !doc.HasBody(run, "wb-para-reverse"))
                {
                    FailRest(run, "铺稿甲失败", "wb-para-reverse");
                    return;
                }

                string pa = doc.Paragraph(Jia);
                string pc = doc.Paragraph(Ding);
                ToolResult rev = s.ApplyParagraphFormat(ParaItems(doc.Id, ParaItem(pc + "," + pa, Align("right"))));
                Ok(run, "wb-para-reverse",
                    rev.Success && RunsAre(rev, pc, pa) && IsAlign(s, doc, pa, "right") && IsAlign(s, doc, pc, "right") && IsAlign(s, doc, doc.Paragraph(Bing), "left"),
                    rev.Success ? "段版式写反没有拆开" : (rev.Error ?? ""));
            }

            using (Planted doc = Plant(s, Body))
            {
                if (doc == null || !doc.HasBody(run, "wb-para-reject-s"))
                {
                    FailRest(run, "铺稿甲失败", "wb-para-reject-s");
                    return;
                }

                ToolResult bad = s.ApplyParagraphFormat(new Dictionary<string, object>
                {
                    ["channel_id"] = doc.Id,
                    ["items"] = new object[]
                    {
                        new Dictionary<string, object>
                        {
                            ["target_codes"] = doc.Sentence(Jia),
                            ["para_format"] = Align("center")
                        }
                    }
                });
                Ok(run, "wb-para-reject-s",
                    bad != null
                        && !bad.Success
                        && (Contains(bad.Error, "target_paragraph_codes") || Contains(bad.Error, "P_"))
                        && IsAlign(s, doc, doc.Paragraph(Jia), "left"),
                    bad != null && bad.Success ? "应失败" : (bad == null ? "空回包" : (bad.Error ?? "对齐被改了")));
            }
        }

        private static void RunCross(TestRun run, WdSession s)
        {
            using (Planted doc = Plant(s, Body))
            {
                if (doc == null || !doc.HasBody(run, "wb-x-s-and-p"))
                {
                    FailRest(run, "铺稿甲失败", "wb-x-s-and-p");
                    return;
                }

                string sb = doc.Sentence(Yi);
                string pc = doc.Paragraph(Ding);
                ToolResult apply = s.ApplyDocumentFormat(CharItems(
                    doc.Id,
                    Item("target_codes", sb, "char_format", Bold()),
                    Item("target_paragraph_codes", pc, "char_format", Under())));
                Ok(run, "wb-x-s-and-p",
                    apply.Success
                        && BoldOf(s, doc, sb) && !UnderOf(s, doc, sb)
                        && !BoldOf(s, doc, doc.Sentence(Jia)) && !UnderOf(s, doc, doc.Sentence(Jia))
                        && UnderOf(s, doc, doc.Sentence(Ding)) && !BoldOf(s, doc, doc.Sentence(Ding))
                        && !BoldOf(s, doc, doc.Sentence(Bing)) && !UnderOf(s, doc, doc.Sentence(Bing)),
                    apply.Success ? "句加粗和段下划线串了" : (apply.Error ?? ""));
            }

            using (Planted doc = Plant(s, Body))
            {
                if (doc == null || !doc.HasBody(run, "wb-x-gaps"))
                {
                    FailRest(run, "铺稿甲失败", "wb-x-gaps");
                    return;
                }

                string sa = doc.Sentence(Jia);
                string sd = doc.Sentence(Ding);
                ToolResult apply = s.ApplyDocumentFormat(CharItems(
                    doc.Id,
                    Item("target_codes", sa + "," + sd, "char_format", Bold()),
                    Item("target_paragraph_codes", doc.Paragraph(Jia) + "," + doc.Paragraph(Ding), "char_format", Under())));
                Ok(run, "wb-x-gaps",
                    apply.Success
                        && BoldOf(s, doc, sa) && UnderOf(s, doc, sa)
                        && !BoldOf(s, doc, doc.Sentence(Yi)) && UnderOf(s, doc, doc.Sentence(Yi))
                        && !BoldOf(s, doc, doc.Sentence(Bing)) && !UnderOf(s, doc, doc.Sentence(Bing))
                        && BoldOf(s, doc, sd) && UnderOf(s, doc, sd),
                    apply.Success ? "跳句加跳段的结果不对" : (apply.Error ?? ""));
            }

            using (Planted doc = Plant(s, Body))
            {
                if (doc == null || !doc.HasBody(run, "wb-x-char-then-para"))
                {
                    FailRest(run, "铺稿甲失败", "wb-x-char-then-para");
                    return;
                }

                string sa = doc.Sentence(Jia);
                string pc = doc.Paragraph(Ding);
                ToolResult ch = s.ApplyDocumentFormat(CharItems(doc.Id, Item("target_codes", sa, "char_format", Bold())));
                ToolResult pa = s.ApplyParagraphFormat(ParaItems(doc.Id, ParaItem(pc, Align("center"))));
                Ok(run, "wb-x-char-then-para",
                    ch.Success && pa.Success
                        && BoldOf(s, doc, sa) && IsAlign(s, doc, doc.Paragraph(Jia), "left")
                        && IsAlign(s, doc, pc, "center") && !BoldOf(s, doc, doc.Sentence(Ding))
                        && !BoldOf(s, doc, doc.Sentence(Yi)) && !BoldOf(s, doc, doc.Sentence(Bing)),
                    "字符=" + (ch.Error ?? "ok") + " 段=" + (pa.Error ?? "ok"));
            }

            using (Planted doc = Plant(s, Body))
            {
                if (doc == null || !doc.HasBody(run, "wb-x-mixed"))
                {
                    FailRest(run, "铺稿甲失败", "wb-x-mixed");
                    return;
                }

                string sa = doc.Sentence(Jia);
                string sc = doc.Sentence(Bing);
                ToolResult apply = s.ApplyDocumentFormat(CharItems(
                    doc.Id,
                    Item("target_codes", sa + "," + doc.Paragraph(Bing), "char_format", Bold()),
                    Item("target_codes", sc, "char_format", Under())));
                JObject warn = WarningAt(apply, 0);
                Ok(run, "wb-x-mixed",
                    apply.Success
                        && warn != null
                        && Contains(Txt(warn, "message"), "混用")
                        && !BoldOf(s, doc, sa)
                        && UnderOf(s, doc, sc)
                        && !BoldOf(s, doc, doc.Sentence(Yi))
                        && !UnderOf(s, doc, doc.Sentence(Ding)),
                    apply.Success ? "混用项应警告且不套" : (apply.Error ?? ""));
            }

            using (Planted doc = Plant(s, Body))
            {
                if (doc == null || !doc.HasBody(run, "wb-x-partial"))
                {
                    FailRest(run, "铺稿甲失败", "wb-x-partial");
                    return;
                }

                string sa = doc.Sentence(Jia);
                ToolResult apply = s.ApplyDocumentFormat(CharItems(
                    doc.Id,
                    Item("target_codes", sa, "char_format", Bold()),
                    Item("target_codes", "S_zzzzz", "char_format", Under())));
                JObject warn = WarningAt(apply, 1);
                Ok(run, "wb-x-partial",
                    apply.Success
                        && warn != null
                        && Contains(Txt(warn, "codes"), "S_zzzzz")
                        && !string.IsNullOrEmpty(Txt(warn, "message"))
                        && BoldOf(s, doc, sa)
                        && !BoldOf(s, doc, doc.Sentence(Yi))
                        && !UnderOf(s, doc, doc.Sentence(Yi)),
                    apply.Success ? "坏码应进 warnings" : (apply.Error ?? ""));
            }

            using (Planted doc = Plant(s, Body))
            {
                if (doc == null || !doc.HasBody(run, "wb-x-none"))
                {
                    FailRest(run, "铺稿甲失败", "wb-x-none");
                    return;
                }

                ToolResult apply = s.ApplyDocumentFormat(CharItems(doc.Id, Item("target_codes", "S_zzzzz", "char_format", Bold())));
                Ok(run, "wb-x-none",
                    apply != null && !apply.Success && Arr(apply, "warnings") != null && Arr(apply, "warnings").Count > 0 && !BoldOf(s, doc, doc.Sentence(Jia)),
                    apply != null && apply.Success ? "全失败却返回成功" : (apply == null ? "空回包" : (apply.Error ?? "无 warnings")));
            }

            using (Planted doc = Plant(s, Body))
            {
                if (doc == null || !doc.HasBody(run, "wb-x-para-partial"))
                {
                    FailRest(run, "铺稿甲失败", "wb-x-para-partial");
                    return;
                }

                string pa = doc.Paragraph(Jia);
                ToolResult apply = s.ApplyParagraphFormat(ParaItems(
                    doc.Id,
                    ParaItem(pa, Align("center")),
                    ParaItem("P_zzzzz", Align("right"))));
                JObject warn = WarningAt(apply, 1);
                Ok(run, "wb-x-para-partial",
                    apply.Success
                        && warn != null
                        && Contains(Txt(warn, "codes"), "P_zzzzz")
                        && IsAlign(s, doc, pa, "center")
                        && IsAlign(s, doc, doc.Paragraph(Bing), "left")
                        && IsAlign(s, doc, doc.Paragraph(Ding), "left"),
                    apply.Success ? "坏段码应进 warnings" : (apply.Error ?? ""));
            }
        }

        private static void RunContract(TestRun run, WdSession s)
        {
            using (Planted doc = Plant(s, Body))
            {
                if (doc == null || !doc.HasBody(run, "wb-neg-empty"))
                {
                    FailRest(run, "铺稿甲失败",
                        "wb-neg-empty", "wb-neg-missing-fmt", "wb-neg-empty-fmt", "wb-neg-mutex",
                        "wb-neg-para-empty", "wb-neg-para-missing-fmt");
                    return;
                }

                ToolResult empty = s.ApplyDocumentFormat(new Dictionary<string, object>
                {
                    ["channel_id"] = doc.Id,
                    ["items"] = new object[0]
                });
                ExpectFailClean(run, s, doc, "wb-neg-empty", empty, "items");

                ToolResult missing = s.ApplyDocumentFormat(new Dictionary<string, object>
                {
                    ["channel_id"] = doc.Id,
                    ["items"] = new object[]
                    {
                        Item("target_codes", doc.Sentence(Jia), "char_format", Bold()),
                        new Dictionary<string, object> { ["target_codes"] = doc.Sentence(Yi) }
                    }
                });
                ExpectFailClean(run, s, doc, "wb-neg-missing-fmt", missing, "char_format");

                ToolResult emptyFmt = s.ApplyDocumentFormat(CharItems(
                    doc.Id,
                    Item("target_codes", doc.Sentence(Jia), "char_format", new Dictionary<string, object>())));
                ExpectFailClean(run, s, doc, "wb-neg-empty-fmt", emptyFmt, "char_format");

                ToolResult mutex = s.ApplyDocumentFormat(new Dictionary<string, object>
                {
                    ["channel_id"] = doc.Id,
                    ["items"] = new object[]
                    {
                        new Dictionary<string, object>
                        {
                            ["target_codes"] = doc.Sentence(Jia),
                            ["target_paragraph_codes"] = doc.Paragraph(Jia),
                            ["char_format"] = Bold()
                        }
                    }
                });
                ExpectFailClean(run, s, doc, "wb-neg-mutex", mutex, "互斥");

                ToolResult paraEmpty = s.ApplyParagraphFormat(new Dictionary<string, object>
                {
                    ["channel_id"] = doc.Id,
                    ["items"] = new object[0]
                });
                Ok(run, "wb-neg-para-empty",
                    paraEmpty != null
                        && !paraEmpty.Success
                        && Contains(paraEmpty.Error, "items")
                        && IsAlign(s, doc, doc.Paragraph(Jia), "left"),
                    paraEmpty != null && paraEmpty.Success ? "应失败" : (paraEmpty == null ? "空回包" : (paraEmpty.Error ?? "对齐被改了")));

                ToolResult paraMissing = s.ApplyParagraphFormat(new Dictionary<string, object>
                {
                    ["channel_id"] = doc.Id,
                    ["items"] = new object[]
                    {
                        ParaItem(doc.Paragraph(Jia), Align("center")),
                        new Dictionary<string, object> { ["target_paragraph_codes"] = doc.Paragraph(Bing) }
                    }
                });
                Ok(run, "wb-neg-para-missing-fmt",
                    paraMissing != null
                        && !paraMissing.Success
                        && Contains(paraMissing.Error, "para_format")
                        && IsAlign(s, doc, doc.Paragraph(Jia), "left"),
                    paraMissing != null && paraMissing.Success ? "应失败" : (paraMissing == null ? "空回包" : (paraMissing.Error ?? "对齐被改了")));
            }
        }

        private static void RunSnapshot(TestRun run, WdSession s)
        {
            using (Planted doc = Plant(s, Body))
            {
                if (doc == null || !doc.HasBody(run, "wb-snap-sent-then-para"))
                {
                    FailRest(run, "铺稿甲失败", "wb-snap-sent-then-para");
                    return;
                }

                string sa = doc.Sentence(Jia);
                string pc = doc.Paragraph(Ding);
                ToolResult ch = s.ApplyDocumentFormat(CharFlat(doc.Id, sa, Bold()));
                ToolResult pa = s.ApplyParagraphFormat(ParaFlat(doc.Id, pc, Align("center")));
                Ok(run, "wb-snap-sent-then-para",
                    ch.Success && pa.Success && BoldOf(s, doc, sa) && IsAlign(s, doc, doc.Paragraph(Jia), "left") && IsAlign(s, doc, pc, "center") && !BoldOf(s, doc, doc.Sentence(Ding)),
                    "句=" + (ch.Error ?? "ok") + " 段=" + (pa.Error ?? "ok"));
            }

            using (Planted doc = PlantParagraphFirst(s, Body))
            {
                if (doc == null || !doc.HasBody(run, "wb-snap-para-then-sent"))
                {
                    FailRest(run, "铺稿甲失败", "wb-snap-para-then-sent");
                    return;
                }

                string pa = doc.Paragraph(Jia);
                string sd = doc.Sentence(Ding);
                ToolResult layout = s.ApplyParagraphFormat(ParaFlat(doc.Id, pa, Align("center")));
                ToolResult ch = s.ApplyDocumentFormat(CharFlat(doc.Id, sd, Bold()));
                Ok(run, "wb-snap-para-then-sent",
                    layout.Success && ch.Success
                        && IsAlign(s, doc, pa, "center")
                        && !BoldOf(s, doc, doc.Sentence(Jia))
                        && !BoldOf(s, doc, doc.Sentence(Yi))
                        && BoldOf(s, doc, sd)
                        && IsAlign(s, doc, doc.Paragraph(Ding), "left"),
                    "段=" + (layout.Error ?? "ok") + " 句=" + (ch.Error ?? "ok"));
            }

            using (Planted doc = Plant(s, Body))
            {
                if (doc == null || !doc.HasBody(run, "wb-snap-after-insert"))
                {
                    FailRest(run, "铺稿甲失败", "wb-snap-after-insert");
                    return;
                }

                if (!s.TypePlain(doc.Path, Jia + Yi + "\r" + Bing + "\r" + Wu + "\r" + Ding))
                {
                    Ok(run, "wb-snap-after-insert", false, "插入戊段失败");
                    return;
                }

                for (int i = 1; i <= 8; i++)
                {
                    s.SetParagraphAlign(doc.Path, i, 0);
                }

                doc.RereadParagraphs(s);
                string pw = doc.Paragraph(Wu);
                ToolResult apply = s.ApplyParagraphFormat(ParaItems(doc.Id, ParaItem(pw, Align("center"))));
                doc.RereadSentences(s);
                string sw = doc.Sentence(Wu);
                string sa = doc.Sentence(Jia);
                string sb = doc.Sentence(Yi);
                string sc = doc.Sentence(Bing);
                string sd = doc.Sentence(Ding);
                bool distinct = !string.IsNullOrEmpty(sw)
                    && sw != sa && sw != sb && sw != sc && sw != sd;
                Ok(run, "wb-snap-after-insert",
                    apply.Success
                        && distinct
                        && IsAlign(s, doc, pw, "center")
                        && IsAlign(s, doc, doc.Paragraph(Jia), "left")
                        && IsAlign(s, doc, doc.Paragraph(Bing), "left")
                        && IsAlign(s, doc, doc.Paragraph(Ding), "left")
                        && doc.TextsRemain(s, Jia, Yi, Bing, Ding, Wu),
                    apply.Success ? "戊段码=" + pw + " 句码=" + sw + " 甲=" + sa : (apply.Error ?? ""));
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
            Ok(run, "wb-no-quit", wordOk && wpsOk, wordDetail + " " + wpsDetail);
        }

        private static void ExpectFailClean(TestRun run, WdSession s, Planted doc, string id, ToolResult result, string needle)
        {
            bool failed = result != null
                && !result.Success
                && (Contains(result.Error, needle) || Contains(JsonConvert.SerializeObject(result.Data), needle));
            Ok(run, id,
                failed && !BoldOf(s, doc, doc.Sentence(Jia)),
                failed ? "契约失败后甲被加粗了" : (result != null && !result.Success ? (result.Error ?? "") : "应失败"));
        }

        private sealed class Planted : IDisposable
        {
            public string Path;
            public string Id;
            public WdSession Session;
            public List<Marker> Sentences = new List<Marker>();
            public List<Marker> Paragraphs = new List<Marker>();
            public string Error;
            private string[] _needles;

            public IEnumerable<string> SentenceCodes
            {
                get
                {
                    foreach (Marker marker in Sentences)
                    {
                        yield return marker.Code;
                    }
                }
            }

            public static Planted FromReads(WdSession s, string path, string id, params string[] needles)
            {
                var doc = new Planted { Path = path, Id = id, Session = s, _needles = needles };
                doc.RereadSentences(s);
                doc.RereadParagraphs(s);
                return doc;
            }

            public bool HasBody(TestRun run, string id)
            {
                if (!string.IsNullOrEmpty(Error))
                {
                    Ok(run, id, false, Error);
                    return false;
                }

                if (string.IsNullOrEmpty(Sentence(Jia)) || string.IsNullOrEmpty(Paragraph(Jia)))
                {
                    Ok(run, id, false, "抄不到甲的句码或段码");
                    return false;
                }

                return true;
            }

            public string Sentence(string needle)
            {
                return Nth(Sentences, needle, 0);
            }

            public string Paragraph(string needle)
            {
                return Nth(Paragraphs, needle, 0);
            }

            public string NthSentenceContaining(string needle, int nth)
            {
                return Nth(Sentences, needle, nth);
            }

            public string NthParagraphContaining(string needle, int nth)
            {
                return Nth(Paragraphs, needle, nth);
            }

            public void RereadSentences(WdSession s)
            {
                ToolResult got = s.GetContent(Id);
                Sentences = Parse(got.Success ? WdSession.Str(got, "document_content") : "", SentencePattern);
            }

            public void RereadParagraphs(WdSession s)
            {
                ToolResult got = s.GetContent(Id, "paragraph");
                Paragraphs = Parse(got.Success ? WdSession.Str(got, "document_content") : "", ParagraphPattern);
            }

            public bool TextsRemain(WdSession s, params string[] extra)
            {
                ToolResult got = s.GetContent(Id);
                string text = got.Success ? WdSession.Str(got, "document_content") : "";
                foreach (string needle in _needles ?? new string[0])
                {
                    if (text.IndexOf(needle, StringComparison.Ordinal) < 0)
                    {
                        return false;
                    }
                }

                foreach (string needle in extra)
                {
                    if (text.IndexOf(needle, StringComparison.Ordinal) < 0)
                    {
                        return false;
                    }
                }

                return true;
            }

            public void Dispose()
            {
                if (Session != null && !string.IsNullOrEmpty(Path))
                {
                    Session.CloseDocByPath(Path);
                }
            }

            private static string Nth(List<Marker> markers, string needle, int nth)
            {
                int seen = 0;
                foreach (Marker marker in markers)
                {
                    if (marker.Text.IndexOf(needle, StringComparison.Ordinal) < 0)
                    {
                        continue;
                    }

                    if (seen == nth)
                    {
                        return marker.Code;
                    }

                    seen++;
                }

                return "";
            }
        }

        private static readonly Regex SentencePattern = new Regex(
            @"\[(S_[0-9A-Za-z]+)-start\]([\s\S]*?)\[\1-end\]",
            RegexOptions.Compiled);

        private static readonly Regex ParagraphPattern = new Regex(
            @"\[(P_[0-9A-Za-z]+(?:_[1-9][0-9]*)?)(?:-start| start)\]([\s\S]*?)\[\1(?:-end| end)\]",
            RegexOptions.Compiled);

        private static Planted Plant(WdSession s, string plain)
        {
            return PlantCore(s, plain, paragraphFirst: false);
        }

        private static Planted PlantParagraphFirst(WdSession s, string plain)
        {
            return PlantCore(s, plain, paragraphFirst: true);
        }

        private static Planted PlantCore(WdSession s, string plain, bool paragraphFirst)
        {
            string path = s.NewDocxPath();
            ToolResult opened = s.Open(path, true, s.App);
            if (!opened.Success)
            {
                return new Planted { Error = opened.Error ?? "新建失败" };
            }

            string id = WdSession.Str(opened, "channel_id");
            if (!s.TypePlain(path, plain))
            {
                s.CloseDocByPath(path);
                return new Planted { Error = "COM 写句失败" };
            }

            for (int i = 1; i <= 8; i++)
            {
                s.SetParagraphChar(path, i, "宋体", 12, false);
                s.SetParagraphAlign(path, i, 0);
            }

            var doc = new Planted
            {
                Path = path,
                Id = id,
                Session = s
            };
            if (paragraphFirst)
            {
                doc.RereadParagraphs(s);
                doc.RereadSentences(s);
            }
            else
            {
                doc.RereadSentences(s);
                doc.RereadParagraphs(s);
            }

            return doc;
        }

        private static List<Marker> Parse(string text, Regex pattern)
        {
            var list = new List<Marker>();
            if (string.IsNullOrEmpty(text))
            {
                return list;
            }

            foreach (Match match in pattern.Matches(text))
            {
                list.Add(new Marker { Code = match.Groups[1].Value, Text = match.Groups[2].Value ?? "" });
            }

            return list;
        }

        private struct Marker
        {
            public string Code;
            public string Text;
        }

        private static Dictionary<string, object> CharFlat(string id, string codes, Dictionary<string, object> format)
        {
            return new Dictionary<string, object>
            {
                ["channel_id"] = id,
                ["target_codes"] = codes,
                ["char_format"] = format
            };
        }

        private static Dictionary<string, object> ParaCharFlat(string id, string codes, Dictionary<string, object> format)
        {
            return new Dictionary<string, object>
            {
                ["channel_id"] = id,
                ["target_paragraph_codes"] = codes,
                ["char_format"] = format
            };
        }

        private static Dictionary<string, object> ParaFlat(string id, string codes, Dictionary<string, object> format)
        {
            return new Dictionary<string, object>
            {
                ["channel_id"] = id,
                ["target_paragraph_codes"] = codes,
                ["para_format"] = format
            };
        }

        private static Dictionary<string, object> CharItems(string id, params Dictionary<string, object>[] items)
        {
            return new Dictionary<string, object>
            {
                ["channel_id"] = id,
                ["items"] = items
            };
        }

        private static Dictionary<string, object> ParaItems(string id, params Dictionary<string, object>[] items)
        {
            return new Dictionary<string, object>
            {
                ["channel_id"] = id,
                ["items"] = items
            };
        }

        private static Dictionary<string, object> Item(string codeKey, string codes, string formatKey, Dictionary<string, object> format)
        {
            return new Dictionary<string, object>
            {
                [codeKey] = codes,
                [formatKey] = format
            };
        }

        private static Dictionary<string, object> ParaItem(string codes, Dictionary<string, object> format)
        {
            return Item("target_paragraph_codes", codes, "para_format", format);
        }

        private static Dictionary<string, object> Bold()
        {
            return new Dictionary<string, object> { ["is_bold"] = true };
        }

        private static Dictionary<string, object> Under()
        {
            return new Dictionary<string, object> { ["has_underline"] = true };
        }

        private static Dictionary<string, object> Strike()
        {
            return new Dictionary<string, object> { ["has_strikethrough"] = true };
        }

        private static Dictionary<string, object> Align(string value)
        {
            return new Dictionary<string, object> { ["alignment"] = value };
        }

        private static bool BoldOf(WdSession s, Planted doc, string code)
        {
            return Flag(CharOf(s, doc.Id, code), "is_bold");
        }

        private static bool UnderOf(WdSession s, Planted doc, string code)
        {
            return Flag(CharOf(s, doc.Id, code), "has_underline");
        }

        private static bool StrikeOf(WdSession s, Planted doc, string code)
        {
            return Flag(CharOf(s, doc.Id, code), "has_strikethrough");
        }

        private static bool IsAlign(WdSession s, Planted doc, string code, string expect)
        {
            string actual = Txt(ParaOf(s, doc.Id, code), "alignment");
            if (string.Equals(expect, "left", StringComparison.OrdinalIgnoreCase)
                && (string.IsNullOrEmpty(actual) || string.Equals(actual, "left", StringComparison.OrdinalIgnoreCase)))
            {
                return true;
            }

            return string.Equals(actual, expect, StringComparison.OrdinalIgnoreCase);
        }

        private static JObject CharOf(WdSession s, string id, string code)
        {
            if (string.IsNullOrEmpty(code))
            {
                return null;
            }

            ToolResult ctx = s.GetFormatContext(new Dictionary<string, object>
            {
                ["channel_id"] = id,
                ["target_codes"] = code
            });
            return Obj(FirstEntry(ctx), "char_format");
        }

        private static JObject ParaOf(WdSession s, string id, string code)
        {
            if (string.IsNullOrEmpty(code))
            {
                return null;
            }

            ToolResult ctx = s.GetFormatContext(new Dictionary<string, object>
            {
                ["channel_id"] = id,
                ["target_paragraph_codes"] = code
            });
            return Obj(FirstEntry(ctx), "para_format");
        }

        private static bool LegacyParaHas(ToolResult result, string code)
        {
            JObject data = Data(result);
            if (data == null || data["applied"] != null)
            {
                return false;
            }

            JToken codes = data["affected_paragraph_codes"] ?? data["applied_codes"];
            return codes != null && Contains(codes.ToString(), code);
        }

        private static bool LegacyHas(ToolResult result, string code)
        {
            JObject data = Data(result);
            if (data == null || data["applied"] != null)
            {
                return false;
            }

            JToken codes = data["applied_codes"];
            return codes != null && Contains(codes.ToString(), code);
        }

        private static bool RunsAre(ToolResult result, params string[] codes)
        {
            JArray applied = Arr(result, "applied");
            if (applied == null || applied.Count != codes.Length)
            {
                return false;
            }

            for (int i = 0; i < codes.Length; i++)
            {
                JObject row = applied[i] as JObject;
                if (row == null || !string.Equals(Txt(row, "codes"), codes[i], StringComparison.Ordinal))
                {
                    return false;
                }
            }

            return true;
        }

        private static bool IndexesAre(ToolResult result, params int[] indexes)
        {
            JArray applied = Arr(result, "applied");
            if (applied == null || applied.Count != indexes.Length)
            {
                return false;
            }

            for (int i = 0; i < indexes.Length; i++)
            {
                JObject row = applied[i] as JObject;
                if (row == null || Txt(row, "item_index") != indexes[i].ToString())
                {
                    return false;
                }
            }

            return true;
        }

        private static int AppliedCount(ToolResult result)
        {
            JArray applied = Arr(result, "applied");
            return applied == null ? 0 : applied.Count;
        }

        private static JObject WarningAt(ToolResult result, int index)
        {
            JArray warnings = Arr(result, "warnings");
            if (warnings == null)
            {
                return null;
            }

            foreach (JToken token in warnings)
            {
                JObject row = token as JObject;
                if (row != null && Txt(row, "item_index") == index.ToString())
                {
                    return row;
                }
            }

            return null;
        }

        private static JArray Arr(ToolResult result, string key)
        {
            JObject data = Data(result);
            return data == null ? null : data[key] as JArray;
        }

        private static JObject Data(ToolResult result)
        {
            if (result == null || result.Data == null)
            {
                return null;
            }

            try
            {
                return JObject.Parse(JsonConvert.SerializeObject(result.Data));
            }
            catch
            {
                return null;
            }
        }

        private static JObject FirstEntry(ToolResult result)
        {
            JArray entries = Arr(result, "entries");
            if (entries == null || entries.Count == 0)
            {
                return null;
            }

            return entries[0] as JObject;
        }

        private static JObject Obj(JObject parent, string key)
        {
            return parent == null ? null : parent[key] as JObject;
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

            string value = o[key].ToString();
            return string.Equals(value, "true", StringComparison.OrdinalIgnoreCase) || value == "1";
        }

        private static void ExpectFail(TestRun run, string id, ToolResult result, string needle)
        {
            if (result != null && !result.Success && (Contains(result.Error, needle) || Contains(JsonConvert.SerializeObject(result.Data), needle)))
            {
                run.CaseOk(id);
                return;
            }

            if (result != null && !result.Success)
            {
                run.CaseFail(id, "失败文案不含「" + needle + "」: " + (result.Error ?? ""));
                return;
            }

            run.CaseFail(id, "应失败");
        }

        private static bool Contains(string text, string needle)
        {
            return !string.IsNullOrEmpty(text)
                && !string.IsNullOrEmpty(needle)
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

        private static void FailRest(TestRun run, string reason, params string[] ids)
        {
            if (ids.Length <= 1)
            {
                return;
            }

            for (int i = 1; i < ids.Length; i++)
            {
                run.CaseFail(ids[i], reason);
            }
        }
    }
}
