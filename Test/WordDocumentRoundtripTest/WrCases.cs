using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using WordAddIn1;

namespace WordDocumentRoundtripTest
{
    /// <summary>
    /// 一次调用用 items 读多处格式。铺稿用 COM，抄当次 S_ / P_。不套格式。
    /// </summary>
    internal static class WrCases
    {
        public const int Target = 27;

        private const string Jia = "易写读格式甲句。";
        private const string Yi = "易写读格式乙句。";
        private const string Bing = "易写读格式丙句。";
        private const string Ding = "易写读格式丁句。";
        private const string Body = Jia + Yi + "\r" + Bing + "\r" + Ding;
        private const string Before = "易写读格式表前句。";
        private const string Left = "易写读格式左格。";
        private const string Right = "易写读格式右格。";
        private const string Dup = "易写读格式重复段。";
        private const string Mid = "易写读格式夹心句。";

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
            RunParagraph(run, s);
            RunCross(run, s);
            RunContract(run, s);
            RunReadonly(run, s);
            RunNoQuit(run, wordBefore, wpsBefore, wordWinBefore, wpsWinBefore);
        }

        private static void RunSentence(TestRun run, WdSession s)
        {
            using (Planted doc = PlantA(s))
            {
                if (doc == null || !doc.HasBody(run, "wr-s-one-flat"))
                {
                    FailRest(run, "铺稿甲失败", "wr-s-one-flat");
                    return;
                }

                string sa = doc.Sentence(Jia);
                ToolResult flat = s.GetFormatContext(FlatSent(doc.Id, sa));
                JObject entry = TargetSent(flat, sa);
                Ok(run, "wr-s-one-flat",
                    flat.Success
                        && LegacyHas(flat, "target_codes", sa)
                        && TargetSentCount(flat) == 1
                        && !HasItemIndex(entry)
                        && IsHeiBold(entry)
                        && ComYiKai(s, doc),
                    flat.Success ? Dump(flat) : (flat.Error ?? ""));
            }

            using (Planted doc = PlantA(s))
            {
                if (doc == null || !doc.HasBody(run, "wr-s-one-item"))
                {
                    FailRest(run, "铺稿甲失败", "wr-s-one-item");
                    return;
                }

                string sa = doc.Sentence(Jia);
                ToolResult one = s.GetFormatContext(Items(doc.Id, Sent(sa)));
                JObject entry = TargetSent(one, sa);
                Ok(run, "wr-s-one-item",
                    one.Success
                        && TargetSentCount(one) == 1
                        && ItemIndex(entry) == 0
                        && Txt(entry, "code") == sa
                        && IsHeiBold(entry),
                    one.Success ? Dump(one) : (one.Error ?? ""));
            }

            using (Planted doc = PlantA(s))
            {
                if (doc == null || !doc.HasBody(run, "wr-s-multi"))
                {
                    FailRest(run, "铺稿甲失败", "wr-s-multi");
                    return;
                }

                string sa = doc.Sentence(Jia);
                string sb = doc.Sentence(Yi);
                string sc = doc.Sentence(Bing);
                ToolResult multi = s.GetFormatContext(Items(doc.Id, Sent(sa), Sent(sb), Sent(sc)));
                JObject ea = TargetSent(multi, sa);
                JObject eb = TargetSent(multi, sb);
                JObject ec = TargetSent(multi, sc);
                Ok(run, "wr-s-multi",
                    multi.Success
                        && TargetSentCount(multi) == 3
                        && ItemIndex(ea) == 0 && ItemIndex(eb) == 1 && ItemIndex(ec) == 2
                        && IsHeiBold(ea)
                        && IsKaiPlain(eb)
                        && IsSongUnder(ec),
                    multi.Success ? Dump(multi) : (multi.Error ?? ""));
            }

            using (Planted doc = PlantA(s))
            {
                if (doc == null || !doc.HasBody(run, "wr-s-run"))
                {
                    FailRest(run, "铺稿甲失败", "wr-s-run");
                    return;
                }

                string sa = doc.Sentence(Jia);
                string sb = doc.Sentence(Yi);
                string sc = doc.Sentence(Bing);
                string sd = doc.Sentence(Ding);
                ToolResult result = s.GetFormatContext(Items(doc.Id, Sent(sa + "," + sb)));
                JObject ea = TargetSent(result, sa);
                JObject eb = TargetSent(result, sb);
                Ok(run, "wr-s-run",
                    result.Success
                        && TargetSentCount(result) == 2
                        && ItemIndex(ea) == 0 && ItemIndex(eb) == 0
                        && IsHeiBold(ea) && IsKaiPlain(eb)
                        && TargetSent(result, sc) == null
                        && TargetSent(result, sd) == null,
                    result.Success ? Dump(result) : (result.Error ?? ""));
            }

            using (Planted doc = PlantA(s))
            {
                if (doc == null || !doc.HasBody(run, "wr-s-gap"))
                {
                    FailRest(run, "铺稿甲失败", "wr-s-gap", "wr-s-gap-flat", "wr-s-reverse", "wr-s-multi-gap");
                    return;
                }

                string sa = doc.Sentence(Jia);
                string sb = doc.Sentence(Yi);
                string sc = doc.Sentence(Bing);
                ToolResult gap = s.GetFormatContext(Items(doc.Id, Sent(sa + "," + sc)));
                JObject ea = TargetSent(gap, sa);
                JObject ec = TargetSent(gap, sc);
                Ok(run, "wr-s-gap",
                    gap.Success
                        && TargetSentCount(gap) == 2
                        && ItemIndex(ea) == 0 && ItemIndex(ec) == 0
                        && IsHeiBold(ea) && IsSongUnder(ec)
                        && TargetSent(gap, sb) == null
                        && !IsKaiPlain(ea) && !IsKaiPlain(ec),
                    gap.Success ? Dump(gap) : (gap.Error ?? ""));
            }

            using (Planted doc = PlantA(s))
            {
                if (doc == null || !doc.HasBody(run, "wr-s-gap-flat"))
                {
                    FailRest(run, "铺稿甲失败", "wr-s-gap-flat");
                    return;
                }

                string sa = doc.Sentence(Jia);
                string sb = doc.Sentence(Yi);
                string sc = doc.Sentence(Bing);
                ToolResult gap = s.GetFormatContext(FlatSent(doc.Id, sa + "," + sc));
                Ok(run, "wr-s-gap-flat",
                    gap.Success
                        && TargetSent(gap, sa) != null
                        && TargetSent(gap, sc) != null
                        && TargetSent(gap, sb) == null
                        && HasItemIndex(TargetSent(gap, sa))
                        && HasItemIndex(TargetSent(gap, sc)),
                    gap.Success ? Dump(gap) : (gap.Error ?? ""));
            }

            using (Planted doc = PlantA(s))
            {
                if (doc == null || !doc.HasBody(run, "wr-s-reverse"))
                {
                    FailRest(run, "铺稿甲失败", "wr-s-reverse");
                    return;
                }

                string sa = doc.Sentence(Jia);
                string sb = doc.Sentence(Yi);
                string sc = doc.Sentence(Bing);
                string sd = doc.Sentence(Ding);
                ToolResult rev = s.GetFormatContext(Items(doc.Id, Sent(sb + "," + sa)));
                List<JObject> sents = TargetSents(rev);
                Ok(run, "wr-s-reverse",
                    rev.Success
                        && sents.Count == 2
                        && Txt(sents[0], "code") == sb
                        && Txt(sents[1], "code") == sa
                        && IsKaiPlain(sents[0]) && IsHeiBold(sents[1])
                        && TargetSent(rev, sc) == null
                        && TargetSent(rev, sd) == null,
                    rev.Success ? Dump(rev) : (rev.Error ?? ""));
            }

            RunCrossSeg(run, s);

            using (Planted doc = PlantA(s))
            {
                if (doc == null || !doc.HasBody(run, "wr-s-multi-gap"))
                {
                    FailRest(run, "铺稿甲失败", "wr-s-multi-gap");
                    return;
                }

                string sa = doc.Sentence(Jia);
                string sb = doc.Sentence(Yi);
                string sc = doc.Sentence(Bing);
                string sd = doc.Sentence(Ding);
                ToolResult mix = s.GetFormatContext(Items(doc.Id, Sent(sa + "," + sd), Sent(sb)));
                JObject ea = TargetSent(mix, sa);
                JObject eb = TargetSent(mix, sb);
                JObject ed = TargetSent(mix, sd);
                Ok(run, "wr-s-multi-gap",
                    mix.Success
                        && ItemIndex(ea) == 0 && ItemIndex(ed) == 0
                        && ItemIndex(eb) == 1 && IsKaiPlain(eb)
                        && TargetSent(mix, sc) == null,
                    mix.Success ? Dump(mix) : (mix.Error ?? ""));
            }
        }

        private static void RunCrossSeg(TestRun run, WdSession s)
        {
            using (Planted doc = PlantB(s))
            {
                if (doc == null || !string.IsNullOrEmpty(doc.Error))
                {
                    Ok(run, "wr-s-cross-seg", false, doc == null ? "稿乙铺失败" : doc.Error);
                    return;
                }

                string before = doc.Sentence(Before);
                string left = doc.Sentence(Left);
                string right = doc.Sentence(Right);
                if (string.IsNullOrEmpty(before) || string.IsNullOrEmpty(left) || string.IsNullOrEmpty(right)
                    || before == left || left == right || before == right)
                {
                    Ok(run, "wr-s-cross-seg", false, "三句若不在不同 seg，记环境失败 " + before + "/" + left + "/" + right);
                    return;
                }

                ToolResult result = s.GetFormatContext(Items(doc.Id, Sent(before + "," + right)));
                JObject eb = TargetSent(result, before);
                JObject er = TargetSent(result, right);
                Ok(run, "wr-s-cross-seg",
                    result.Success
                        && TargetSentCount(result) >= 2
                        && IsHeiBold(eb)
                        && IsKaiUnder(er)
                        && TargetSent(result, left) == null,
                    result.Success ? Dump(result) : (result.Error ?? ""));
            }
        }

        private static void RunParagraph(TestRun run, WdSession s)
        {
            using (Planted doc = PlantA(s))
            {
                if (doc == null || !doc.HasBody(run, "wr-p-one-flat"))
                {
                    FailRest(run, "铺稿甲失败", "wr-p-one-flat");
                    return;
                }

                string pa = doc.Paragraph(Jia);
                string pc = doc.Paragraph(Ding);
                ToolResult flat = s.GetFormatContext(FlatPara(doc.Id, pa));
                JObject entry = TargetPara(flat, pa);
                Ok(run, "wr-p-one-flat",
                    flat.Success
                        && LegacyHas(flat, "target_paragraph_codes", pa)
                        && TargetParaCount(flat) == 1
                        && !HasItemIndex(entry)
                        && IsAlign(entry, "left")
                        && TargetPara(flat, pc) == null,
                    flat.Success ? Dump(flat) : (flat.Error ?? ""));
            }

            using (Planted doc = PlantA(s))
            {
                if (doc == null || !doc.HasBody(run, "wr-p-one-item"))
                {
                    FailRest(run, "铺稿甲失败", "wr-p-one-item");
                    return;
                }

                string pc = doc.Paragraph(Ding);
                ToolResult one = s.GetFormatContext(Items(doc.Id, Para(pc)));
                JObject entry = TargetPara(one, pc);
                Ok(run, "wr-p-one-item",
                    one.Success
                        && TargetParaCount(one) == 1
                        && ItemIndex(entry) == 0
                        && IsAlign(entry, "center"),
                    one.Success ? Dump(one) : (one.Error ?? ""));
            }

            using (Planted doc = PlantA(s))
            {
                if (doc == null || !doc.HasBody(run, "wr-p-multi"))
                {
                    FailRest(run, "铺稿甲失败", "wr-p-multi");
                    return;
                }

                string pa = doc.Paragraph(Jia);
                string pb = doc.Paragraph(Bing);
                string pc = doc.Paragraph(Ding);
                ToolResult multi = s.GetFormatContext(Items(doc.Id, Para(pa), Para(pc)));
                JObject ea = TargetPara(multi, pa);
                JObject ec = TargetPara(multi, pc);
                Ok(run, "wr-p-multi",
                    multi.Success
                        && ItemIndex(ea) == 0 && ItemIndex(ec) == 1
                        && IsAlign(ea, "left") && IsAlign(ec, "center")
                        && TargetPara(multi, pb) == null,
                    multi.Success ? Dump(multi) : (multi.Error ?? ""));
            }

            using (Planted doc = PlantA(s))
            {
                if (doc == null || !doc.HasBody(run, "wr-p-gap"))
                {
                    FailRest(run, "铺稿甲失败", "wr-p-gap");
                    return;
                }

                string pa = doc.Paragraph(Jia);
                string pb = doc.Paragraph(Bing);
                string pc = doc.Paragraph(Ding);
                ToolResult gap = s.GetFormatContext(Items(doc.Id, Para(pa + "," + pc)));
                JObject ea = TargetPara(gap, pa);
                JObject ec = TargetPara(gap, pc);
                Ok(run, "wr-p-gap",
                    gap.Success
                        && TargetParaCount(gap) == 2
                        && ItemIndex(ea) == 0 && ItemIndex(ec) == 0
                        && IsAlign(ea, "left") && IsAlign(ec, "center")
                        && TargetPara(gap, pb) == null,
                    gap.Success ? Dump(gap) : (gap.Error ?? ""));
            }

            using (Planted doc = PlantA(s))
            {
                if (doc == null || !doc.HasBody(run, "wr-p-reverse"))
                {
                    FailRest(run, "铺稿甲失败", "wr-p-reverse");
                    return;
                }

                string pa = doc.Paragraph(Jia);
                string pb = doc.Paragraph(Bing);
                string pc = doc.Paragraph(Ding);
                ToolResult rev = s.GetFormatContext(Items(doc.Id, Para(pc + "," + pa)));
                List<JObject> paras = TargetParas(rev);
                Ok(run, "wr-p-reverse",
                    rev.Success
                        && paras.Count == 2
                        && Txt(paras[0], "paragraph_code") == pc
                        && Txt(paras[1], "paragraph_code") == pa
                        && IsAlign(paras[0], "center") && IsAlign(paras[1], "left")
                        && TargetPara(rev, pb) == null,
                    rev.Success ? Dump(rev) : (rev.Error ?? ""));
            }

            using (Planted doc = PlantC(s))
            {
                if (doc == null || !string.IsNullOrEmpty(doc.Error))
                {
                    Ok(run, "wr-p-collision", false, doc == null ? "稿丙铺失败" : doc.Error);
                    return;
                }

                string first = doc.NthParagraphContaining(Dup, 0);
                string second = doc.NthParagraphContaining(Dup, 1);
                if (string.IsNullOrEmpty(first) || string.IsNullOrEmpty(second)
                    || first.IndexOf("_", 2, StringComparison.Ordinal) < 0)
                {
                    Ok(run, "wr-p-collision", false, "没有撞码公开码 " + first + " / " + second);
                    return;
                }

                ToolResult result = s.GetFormatContext(Items(doc.Id, Para(first)));
                bool firstUl = false;
                bool midUl = false;
                bool secondUl = false;
                bool read = s.TryNthParagraphUnderline(doc.Path, Dup, 0, out firstUl)
                    && s.TryNthParagraphUnderline(doc.Path, Mid, 0, out midUl)
                    && s.TryNthParagraphUnderline(doc.Path, Dup, 1, out secondUl);
                Ok(run, "wr-p-collision",
                    result.Success && read && firstUl && !midUl && !secondUl && doc.TextsRemain(s),
                    result.Success
                        ? "下划线 first=" + firstUl + " mid=" + midUl + " second=" + secondUl + " codes=" + first + "," + second
                        : (result.Error ?? ""));
            }
        }

        private static void RunCross(TestRun run, WdSession s)
        {
            using (Planted doc = PlantA(s))
            {
                if (doc == null || !doc.HasBody(run, "wr-x-s-and-p"))
                {
                    FailRest(run, "铺稿甲失败", "wr-x-s-and-p");
                    return;
                }

                string sa = doc.Sentence(Jia);
                string sb = doc.Sentence(Yi);
                string sd = doc.Sentence(Ding);
                string pc = doc.Paragraph(Ding);
                ToolResult result = s.GetFormatContext(Items(doc.Id, Sent(sb), Para(pc)));
                JObject eb = TargetSent(result, sb);
                JObject ep = TargetPara(result, pc);
                Ok(run, "wr-x-s-and-p",
                    result.Success
                        && TargetSentCount(result) + TargetParaCount(result) >= 2
                        && IsKaiPlain(eb) && ItemIndex(eb) == 0
                        && IsAlign(ep, "center") && ItemIndex(ep) == 1
                        && TargetSent(result, sa) == null
                        && TargetSent(result, sd) == null
                        && !IsHeiBold(eb),
                    result.Success ? Dump(result) : (result.Error ?? ""));
            }

            using (Planted doc = PlantA(s))
            {
                if (doc == null || !doc.HasBody(run, "wr-x-gaps"))
                {
                    FailRest(run, "铺稿甲失败", "wr-x-gaps");
                    return;
                }

                string sa = doc.Sentence(Jia);
                string sb = doc.Sentence(Yi);
                string sd = doc.Sentence(Ding);
                string pa = doc.Paragraph(Jia);
                string pb = doc.Paragraph(Bing);
                string pc = doc.Paragraph(Ding);
                ToolResult result = s.GetFormatContext(Items(doc.Id, Sent(sa + "," + sd), Para(pa + "," + pc)));
                Ok(run, "wr-x-gaps",
                    result.Success
                        && TargetSent(result, sa) != null
                        && TargetSent(result, sd) != null
                        && TargetSent(result, sb) == null
                        && TargetPara(result, pb) == null
                        && IsAlign(TargetPara(result, pc), "center"),
                    result.Success ? Dump(result) : (result.Error ?? ""));
            }

            using (Planted doc = PlantA(s))
            {
                if (doc == null || !doc.HasBody(run, "wr-x-mixed"))
                {
                    FailRest(run, "铺稿甲失败", "wr-x-mixed");
                    return;
                }

                string sa = doc.Sentence(Jia);
                string pb = doc.Paragraph(Bing);
                string sc = doc.Sentence(Bing);
                ToolResult result = s.GetFormatContext(Items(doc.Id, Sent(sa + "," + pb), Sent(sc)));
                JObject warn = WarningAt(result, 0);
                JObject ec = TargetSent(result, sc);
                Ok(run, "wr-x-mixed",
                    result.Success
                        && warn != null
                        && (Contains(Txt(warn, "message"), "混用") || Contains(Txt(warn, "message"), "未读取"))
                        && TargetSent(result, sa) == null
                        && IsSongUnder(ec) && ItemIndex(ec) == 1,
                    result.Success ? Dump(result) : (result.Error ?? ""));
            }

            using (Planted doc = PlantA(s))
            {
                if (doc == null || !doc.HasBody(run, "wr-x-partial"))
                {
                    FailRest(run, "铺稿甲失败", "wr-x-partial");
                    return;
                }

                string sa = doc.Sentence(Jia);
                string sb = doc.Sentence(Yi);
                string sc = doc.Sentence(Bing);
                string sd = doc.Sentence(Ding);
                ToolResult result = s.GetFormatContext(Items(doc.Id, Sent(sa), Sent("S_zzzzz")));
                JObject warn = WarningAt(result, 1);
                Ok(run, "wr-x-partial",
                    result.Success
                        && warn != null
                        && Contains(Txt(warn, "codes") + Txt(warn, "message"), "S_zzzzz")
                        && IsHeiBold(TargetSent(result, sa))
                        && TargetSent(result, sb) == null
                        && TargetSent(result, sc) == null
                        && TargetSent(result, sd) == null,
                    result.Success ? Dump(result) : (result.Error ?? ""));
            }

            using (Planted doc = PlantA(s))
            {
                if (doc == null || !doc.HasBody(run, "wr-x-none"))
                {
                    FailRest(run, "铺稿甲失败", "wr-x-none");
                    return;
                }

                ToolResult result = s.GetFormatContext(Items(doc.Id, Sent("S_zzzzz")));
                Ok(run, "wr-x-none",
                    result != null && !result.Success && Arr(result, "warnings") != null && Arr(result, "warnings").Count > 0
                        && TargetSentCount(result) == 0,
                    result != null && !result.Success ? Dump(result) : "应失败");
            }

            using (Planted doc = PlantA(s))
            {
                if (doc == null || !doc.HasBody(run, "wr-x-para-partial"))
                {
                    FailRest(run, "铺稿甲失败", "wr-x-para-partial");
                    return;
                }

                string pa = doc.Paragraph(Jia);
                string pb = doc.Paragraph(Bing);
                string pc = doc.Paragraph(Ding);
                ToolResult result = s.GetFormatContext(Items(doc.Id, Para(pc), Para("P_zzzzz")));
                JObject warn = WarningAt(result, 1);
                Ok(run, "wr-x-para-partial",
                    result.Success
                        && warn != null
                        && Contains(Txt(warn, "codes") + Txt(warn, "message"), "P_zzzzz")
                        && IsAlign(TargetPara(result, pc), "center")
                        && TargetPara(result, pa) == null
                        && TargetPara(result, pb) == null,
                    result.Success ? Dump(result) : (result.Error ?? ""));
            }
        }

        private static void RunContract(TestRun run, WdSession s)
        {
            using (Planted doc = PlantA(s))
            {
                if (doc == null || !doc.HasBody(run, "wr-neg-empty"))
                {
                    FailRest(run, "铺稿甲失败", "wr-neg-empty", "wr-neg-missing", "wr-neg-mutex", "wr-neg-t");
                    return;
                }

                ToolResult empty = s.GetFormatContext(new Dictionary<string, object>
                {
                    ["channel_id"] = doc.Id,
                    ["context_window"] = 0,
                    ["items"] = new object[0]
                });
                ExpectFailClean(run, s, doc, "wr-neg-empty", empty, "items");

                ToolResult missing = s.GetFormatContext(Items(doc.Id, new Dictionary<string, object>()));
                string missingText = (missing == null ? "" : (missing.Error ?? "")) + Json(missing);
                bool missingNeedle = Contains(missingText, "编码") || Contains(missingText, "target");
                Ok(run, "wr-neg-missing",
                    missing != null && !missing.Success && missingNeedle && NoEntries(missing)
                        && ComJiaHei(s, doc) && ComYiKai(s, doc),
                    missing != null && !missing.Success ? (missing.Error ?? "") : "应失败");

                string sa = doc.Sentence(Jia);
                string pa = doc.Paragraph(Jia);
                ToolResult mutex = s.GetFormatContext(Items(doc.Id, new Dictionary<string, object>
                {
                    ["target_codes"] = sa,
                    ["target_paragraph_codes"] = pa
                }));
                ExpectFailClean(run, s, doc, "wr-neg-mutex", mutex, "互斥");

                ToolResult table = s.GetFormatContext(Items(doc.Id, Sent("T_abcde")));
                bool tableFail = table != null
                    && (!table.Success || TargetSentCount(table) == 0)
                    && Contains((table.Error ?? "") + Json(table), "T_");
                Ok(run, "wr-neg-t",
                    tableFail && ComJiaHei(s, doc) && ComYiKai(s, doc),
                    table != null && table.Success ? Dump(table) : (table == null ? "无回包" : (table.Error ?? "")));
            }
        }

        private static void RunReadonly(TestRun run, WdSession s)
        {
            using (Planted doc = PlantA(s))
            {
                if (doc == null || !doc.HasBody(run, "wr-ro-keep"))
                {
                    FailRest(run, "铺稿甲失败", "wr-ro-keep");
                    return;
                }

                string sa = doc.Sentence(Jia);
                string sc = doc.Sentence(Bing);
                string pc = doc.Paragraph(Ding);
                ToolResult result = s.GetFormatContext(Items(doc.Id, Sent(sa), Sent(sc), Para(pc)));
                bool firstUl = false;
                Ok(run, "wr-ro-keep",
                    result.Success
                        && doc.TextsRemain(s)
                        && ComJiaHei(s, doc)
                        && ComYiKai(s, doc)
                        && s.TryNthParagraphUnderline(doc.Path, Bing, 0, out firstUl) && firstUl
                        && ComDingCenter(s, doc),
                    result.Success ? Dump(result) : (result.Error ?? ""));
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
            Ok(run, "wr-no-quit", wordOk && wpsOk, wordDetail + " " + wpsDetail);
        }

        private static void ExpectFailClean(TestRun run, WdSession s, Planted doc, string id, ToolResult result, string needle)
        {
            bool failed = result != null
                && !result.Success
                && (Contains(result.Error, needle) || Contains(Json(result), needle))
                && NoEntries(result);
            Ok(run, id,
                failed && ComJiaHei(s, doc) && ComYiKai(s, doc),
                failed ? "契约失败后格式被改" : (result != null && !result.Success ? (result.Error ?? "") : "应失败"));
        }

        private sealed class Planted : IDisposable
        {
            public string Path;
            public string Id;
            public WdSession Session;
            public List<Marker> Sentences = new List<Marker>();
            public List<Marker> Paragraphs = new List<Marker>();
            public string Error;
            public string[] Needles;

            public static Planted FromReads(WdSession s, string path, string id, params string[] needles)
            {
                var doc = new Planted { Path = path, Id = id, Session = s, Needles = needles };
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

                if (string.IsNullOrEmpty(Sentence(Jia)) || string.IsNullOrEmpty(Paragraph(Jia))
                    || string.IsNullOrEmpty(Sentence(Yi)) || string.IsNullOrEmpty(Sentence(Bing))
                    || string.IsNullOrEmpty(Sentence(Ding)) || string.IsNullOrEmpty(Paragraph(Ding)))
                {
                    Ok(run, id, false, "抄不到甲乙丙丁的句码或段码");
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
                foreach (string needle in Needles ?? new string[0])
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

        private static Planted PlantA(WdSession s)
        {
            Planted doc = PlantCore(s, Body, Jia, Yi, Bing, Ding);
            if (doc == null || !string.IsNullOrEmpty(doc.Error))
            {
                return doc;
            }

            if (!StyleA(s, doc.Path))
            {
                s.CloseDocByPath(doc.Path);
                return new Planted { Error = "稿甲套 COM 格式失败" };
            }

            doc.RereadSentences(s);
            doc.RereadParagraphs(s);
            return doc;
        }

        private static Planted PlantB(WdSession s)
        {
            string path = s.NewDocxPath();
            ToolResult opened = s.Open(path, true, s.App);
            if (!opened.Success || !s.TypePlain(path, Before) || !s.InsertTable(path, 1, 2, new[] { Left, Right }))
            {
                if (opened.Success)
                {
                    s.CloseDocByPath(path);
                }

                return new Planted { Error = opened.Error ?? "稿乙铺失败" };
            }

            s.SetParagraphChar(path, 1, "黑体", 16, true);
            if (!s.SetTextChar(path, Before, "黑体", 16, true, false)
                || !s.SetTextChar(path, Left, "宋体", 12, false, false)
                || !s.SetTextChar(path, Right, "楷体", 12, false, true))
            {
                s.CloseDocByPath(path);
                return new Planted { Error = "稿乙套 COM 格式失败" };
            }

            string id = WdSession.Str(opened, "channel_id");
            return Planted.FromReads(s, path, id, Before, Left, Right);
        }

        private static Planted PlantC(WdSession s)
        {
            Planted doc = PlantCore(s, Dup + "\r" + Mid + "\r" + Dup, Dup, Mid);
            if (doc == null || !string.IsNullOrEmpty(doc.Error))
            {
                return doc;
            }

            if (!s.SetNthParagraphUnderline(doc.Path, Dup, 0, true))
            {
                s.CloseDocByPath(doc.Path);
                return new Planted { Error = "稿丙下划线失败" };
            }

            doc.RereadSentences(s);
            doc.RereadParagraphs(s);
            return doc;
        }

        private static Planted PlantCore(WdSession s, string plain, params string[] needles)
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
                Session = s,
                Needles = needles
            };
            doc.RereadSentences(s);
            doc.RereadParagraphs(s);
            return doc;
        }

        private static bool StyleA(WdSession s, string path)
        {
            return s.SetTextChar(path, Jia, "黑体", 16, true, false)
                && s.SetTextChar(path, Yi, "楷体", 12, false, false)
                && s.SetTextChar(path, Bing, "宋体", 12, false, true)
                && s.SetParagraphAlign(path, 3, 1);
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

        private static Dictionary<string, object> FlatSent(string id, string codes)
        {
            return new Dictionary<string, object>
            {
                ["channel_id"] = id,
                ["context_window"] = 0,
                ["target_codes"] = codes
            };
        }

        private static Dictionary<string, object> FlatPara(string id, string codes)
        {
            return new Dictionary<string, object>
            {
                ["channel_id"] = id,
                ["context_window"] = 0,
                ["target_paragraph_codes"] = codes
            };
        }

        private static Dictionary<string, object> Items(string id, params Dictionary<string, object>[] items)
        {
            return new Dictionary<string, object>
            {
                ["channel_id"] = id,
                ["context_window"] = 0,
                ["items"] = items
            };
        }

        private static Dictionary<string, object> Sent(string codes)
        {
            return new Dictionary<string, object> { ["target_codes"] = codes };
        }

        private static Dictionary<string, object> Para(string codes)
        {
            return new Dictionary<string, object> { ["target_paragraph_codes"] = codes };
        }

        private static bool ComJiaHei(WdSession s, Planted doc)
        {
            string font;
            bool bold;
            bool under;
            return s.TryTextChar(doc.Path, Jia, out font, out bold, out under)
                && FontLooksLike(font, "黑体")
                && bold;
        }

        private static bool ComYiKai(WdSession s, Planted doc)
        {
            string font;
            bool bold;
            bool under;
            return s.TryTextChar(doc.Path, Yi, out font, out bold, out under)
                && FontLooksLike(font, "楷体")
                && !bold;
        }

        private static bool ComDingCenter(WdSession s, Planted doc)
        {
            int align;
            return s.TryNthParagraphAlign(doc.Path, Ding, 0, out align) && align == 1;
        }

        private static bool IsHeiBold(JObject entry)
        {
            JObject cf = Obj(entry, "char_format");
            return FontLooksLike(Txt(cf, "font_name"), "黑体") && Flag(cf, "is_bold");
        }

        private static bool IsKaiPlain(JObject entry)
        {
            JObject cf = Obj(entry, "char_format");
            return FontLooksLike(Txt(cf, "font_name"), "楷体")
                && !Flag(cf, "is_bold")
                && !Flag(cf, "has_underline");
        }

        private static bool IsKaiUnder(JObject entry)
        {
            JObject cf = Obj(entry, "char_format");
            return FontLooksLike(Txt(cf, "font_name"), "楷体") && Flag(cf, "has_underline");
        }

        private static bool IsSongUnder(JObject entry)
        {
            JObject cf = Obj(entry, "char_format");
            return Flag(cf, "has_underline") && !Flag(cf, "is_bold");
        }

        private static bool IsAlign(JObject entry, string expect)
        {
            string actual = Txt(Obj(entry, "para_format"), "alignment");
            if (string.Equals(expect, "left", StringComparison.OrdinalIgnoreCase)
                && (string.IsNullOrEmpty(actual) || string.Equals(actual, "left", StringComparison.OrdinalIgnoreCase)))
            {
                return true;
            }

            return string.Equals(actual, expect, StringComparison.OrdinalIgnoreCase);
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

        private static bool IsSentEntry(JObject row)
        {
            string code = Txt(row, "code");
            return !string.IsNullOrEmpty(code) && code.StartsWith("S_", StringComparison.Ordinal);
        }

        private static bool IsParaEntry(JObject row)
        {
            return !IsSentEntry(row) && !string.IsNullOrEmpty(Txt(row, "paragraph_code"));
        }

        private static List<JObject> TargetSents(ToolResult result)
        {
            var list = new List<JObject>();
            JArray entries = Arr(result, "entries");
            if (entries == null)
            {
                return list;
            }

            foreach (JToken token in entries)
            {
                JObject row = token as JObject;
                if (row != null && IsSentEntry(row) && Flag(row, "is_target"))
                {
                    list.Add(row);
                }
                else if (row != null && IsSentEntry(row) && row["is_target"] == null)
                {
                    list.Add(row);
                }
            }

            return list;
        }

        private static List<JObject> TargetParas(ToolResult result)
        {
            var list = new List<JObject>();
            JArray entries = Arr(result, "entries");
            if (entries == null)
            {
                return list;
            }

            foreach (JToken token in entries)
            {
                JObject row = token as JObject;
                if (row == null || !IsParaEntry(row))
                {
                    continue;
                }

                if (Flag(row, "is_target") || row["is_target"] == null)
                {
                    list.Add(row);
                }
            }

            return list;
        }

        private static int TargetSentCount(ToolResult result)
        {
            return TargetSents(result).Count;
        }

        private static int TargetParaCount(ToolResult result)
        {
            return TargetParas(result).Count;
        }

        private static JObject TargetSent(ToolResult result, string code)
        {
            if (string.IsNullOrEmpty(code))
            {
                return null;
            }

            foreach (JObject row in TargetSents(result))
            {
                if (Txt(row, "code") == code)
                {
                    return row;
                }
            }

            return null;
        }

        private static JObject TargetPara(ToolResult result, string code)
        {
            if (string.IsNullOrEmpty(code))
            {
                return null;
            }

            foreach (JObject row in TargetParas(result))
            {
                if (Txt(row, "paragraph_code") == code)
                {
                    return row;
                }
            }

            return null;
        }

        private static bool HasItemIndex(JObject row)
        {
            return row != null && row["item_index"] != null && row["item_index"].Type != JTokenType.Null;
        }

        private static int ItemIndex(JObject row)
        {
            if (!HasItemIndex(row))
            {
                return int.MinValue;
            }

            return Convert.ToInt32(row["item_index"].ToString());
        }

        private static bool LegacyHas(ToolResult result, string key, string code)
        {
            JObject data = Data(result);
            if (data == null)
            {
                return false;
            }

            JToken codes = data[key];
            return codes != null && Contains(codes.ToString(), code);
        }

        private static bool NoEntries(ToolResult result)
        {
            JArray entries = Arr(result, "entries");
            return entries == null || entries.Count == 0;
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

        private static bool Contains(string text, string needle)
        {
            return !string.IsNullOrEmpty(text)
                && !string.IsNullOrEmpty(needle)
                && text.IndexOf(needle, StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private static string Json(ToolResult result)
        {
            try
            {
                return JsonConvert.SerializeObject(result == null ? null : result.Data);
            }
            catch
            {
                return "";
            }
        }

        private static string Dump(ToolResult result)
        {
            string json = Json(result);
            if (json.Length > 400)
            {
                json = json.Substring(0, 400);
            }

            return json;
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
            if (ids == null || ids.Length == 0)
            {
                return;
            }

            Ok(run, ids[0], false, reason);
            for (int i = 1; i < ids.Length; i++)
            {
                run.CaseFail(ids[i], reason);
            }
        }
    }
}
