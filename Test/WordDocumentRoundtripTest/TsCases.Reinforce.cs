using System;
using System.Collections.Generic;
using Newtonsoft.Json.Linq;
using WordAddIn1;

namespace WordDocumentRoundtripTest
{
    internal static partial class TsCases
    {
        private const string Three = "甲[1]乙[1]丙[1]。";
        private const string Pair = "甲甲甲甲完";
        private const string ParaTwoMarks = "段前甲[1]段中[1]段后。";
        private const string ParaOther = "另一段甲[1]完。";

        private static void RunReinforce(TestRun run, WdSession s)
        {
            RunObjectForm(run, s);
            RunEmptyText(run, s);
            RunPartial(run, s);
            RunParagraphExtra(run, s);
            RunItemsExtra(run, s);
            RunWriteExtra(run, s);
        }

        private static void RunObjectForm(TestRun run, WdSession s)
        {
            using (Planted doc = Plant(s, One))
            {
                if (!Ready(run, doc, "ts-x-obj-once"))
                {
                    return;
                }

                if (Want("ts-x-obj-once"))
                {
                    string code = doc.Sentence(One);
                    ToolResult applied = s.ApplyDocumentFormat(Flat(doc.Id, "target_codes", code, TextValue("[1]"), Super()));
                    Ok(run, "ts-x-obj-once",
                        applied.Success && OnlySuper(s, doc, code, false, "[1]"),
                        applied.Success ? "不带 index 的对象没有只套 [1]" : (applied.Error ?? ""));
                }
            }

            using (Planted doc = Plant(s, Two))
            {
                if (!Ready(run, doc, "ts-x-obj-amb", "ts-x-index0", "ts-x-index-neg", "ts-x-index-bad"))
                {
                    return;
                }

                string code = doc.Sentence(Two);
                if (Want("ts-x-obj-amb"))
                {
                    ToolResult applied = s.ApplyDocumentFormat(Flat(doc.Id, "target_codes", code, TextValue("[1]"), Super()));
                    Ok(run, "ts-x-obj-amb",
                        Failed(applied, "出现 2 次，有歧义") && !AnySuper(s, doc, code, false),
                        applied != null && !applied.Success ? (applied.Error ?? "") : "对象不带 index 仍套上了");
                }

                if (Want("ts-x-index0"))
                {
                    ToolResult applied = s.ApplyDocumentFormat(Flat(doc.Id, "target_codes", code, TextObj("[1]", 0), Super()));
                    Ok(run, "ts-x-index0",
                        Failed(applied, "index 超出范围") && !AnySuper(s, doc, code, false),
                        applied != null && !applied.Success ? (applied.Error ?? "") : "index 0 仍套上了");
                }

                if (Want("ts-x-index-neg"))
                {
                    ToolResult applied = s.ApplyDocumentFormat(Flat(doc.Id, "target_codes", code, TextObj("[1]", -1), Super()));
                    Ok(run, "ts-x-index-neg",
                        Failed(applied, "index 超出范围") && !AnySuper(s, doc, code, false),
                        applied != null && !applied.Success ? (applied.Error ?? "") : "负数 index 仍套上了");
                }

                if (Want("ts-x-index-bad"))
                {
                    var text = new Dictionary<string, object> { ["value"] = "[1]", ["index"] = "二" };
                    ToolResult applied = s.ApplyDocumentFormat(Flat(doc.Id, "target_codes", code, text, Super()));
                    Ok(run, "ts-x-index-bad",
                        Failed(applied, "index 超出范围") && !AnySuper(s, doc, code, false),
                        applied != null && !applied.Success ? (applied.Error ?? "") : "非整数 index 仍套上了");
                }
            }
        }

        private static void RunEmptyText(TestRun run, WdSession s)
        {
            using (Planted doc = Plant(s, One))
            {
                if (!Ready(run, doc, "ts-x-empty"))
                {
                    return;
                }

                if (Want("ts-x-empty"))
                {
                    string code = doc.Sentence(One);
                    ToolResult applied = s.ApplyDocumentFormat(Flat(doc.Id, "target_codes", code, "", Super()));
                    Ok(run, "ts-x-empty",
                        applied.Success && AllSuper(s, doc, code, false) && doc.Has(One),
                        applied.Success ? "空 text 没有按整句套" : (applied.Error ?? ""));
                }
            }

            using (Planted doc = Plant(s, One))
            {
                if (!Ready(run, doc, "ts-x-blank"))
                {
                    return;
                }

                if (Want("ts-x-blank"))
                {
                    string code = doc.Sentence(One);
                    ToolResult applied = s.ApplyDocumentFormat(Flat(doc.Id, "target_codes", code, "   ", Super()));
                    Ok(run, "ts-x-blank",
                        applied.Success && AllSuper(s, doc, code, false),
                        applied.Success ? "空白 text 没有按整句套" : (applied.Error ?? ""));
                }
            }

            using (Planted doc = Plant(s, One))
            {
                if (!Ready(run, doc, "ts-x-value-empty"))
                {
                    return;
                }

                if (Want("ts-x-value-empty"))
                {
                    string code = doc.Sentence(One);
                    var empty = new Dictionary<string, object> { ["value"] = "", ["index"] = 1 };
                    ToolResult applied = s.ApplyDocumentFormat(Flat(doc.Id, "target_codes", code, empty, Super()));
                    Ok(run, "ts-x-value-empty",
                        Failed(applied, "text.value 不能为空") && !AnySuper(s, doc, code, false),
                        applied != null && !applied.Success ? (applied.Error ?? "") : "空 value 仍套上了");
                }
            }
        }

        private static void RunPartial(TestRun run, WdSession s)
        {
            Slice(run, s, "ts-x-mid", One, "方法");
            Slice(run, s, "ts-x-head", One, "见");
            Slice(run, s, "ts-x-tail", One, "说明");
            Slice(run, s, "ts-x-literal", One, "方法[");

            using (Planted doc = Plant(s, Three))
            {
                if (!Ready(run, doc, "ts-x-three"))
                {
                    return;
                }

                if (Want("ts-x-three"))
                {
                    string code = doc.Sentence(Three);
                    ToolResult applied = s.ApplyDocumentFormat(Flat(doc.Id, "target_codes", code, TextObj("[1]", 2), Super()));
                    JObject entry = Entry(s, doc, code, false);
                    Ok(run, "ts-x-three",
                        applied.Success && NthMarkOnly(entry, "[1]", 2),
                        applied.Success ? "三处里的第二处没有单独上标: " + Dump(entry) : (applied.Error ?? ""));
                }
            }

            using (Planted doc = Plant(s, Pair))
            {
                if (!Ready(run, doc, "ts-x-pair2"))
                {
                    return;
                }

                if (Want("ts-x-pair2"))
                {
                    string code = doc.Sentence(Pair);
                    ToolResult applied = s.ApplyDocumentFormat(Flat(doc.Id, "target_codes", code, TextObj("甲甲", 2), Super()));
                    JObject entry = Entry(s, doc, code, false);
                    Ok(run, "ts-x-pair2",
                        applied.Success && NthMarkOnly(entry, "甲甲", 2),
                        applied.Success ? "第二对甲甲没有单独上标: " + Dump(entry) : (applied.Error ?? ""));
                }
            }
        }

        private static void RunParagraphExtra(TestRun run, WdSession s)
        {
            using (Planted doc = Plant(s, Para + "\r" + ParaOther))
            {
                if (!Ready(run, doc, "ts-x-p-two"))
                {
                    return;
                }

                if (Want("ts-x-p-two"))
                {
                    string a = doc.Paragraph(Para);
                    string b = doc.Paragraph(ParaOther);
                    ToolResult applied = s.ApplyDocumentFormat(Flat(doc.Id, "target_codes", a + "," + b, "[1]", Super()));
                    Ok(run, "ts-x-p-two",
                        Failed(applied, "恰好一个")
                            && !AnySuper(s, doc, doc.Sentence(Para), false)
                            && !AnySuper(s, doc, doc.Sentence(ParaOther), false),
                        applied != null && !applied.Success ? (applied.Error ?? "") : "两个段码带 text 仍套上了");
                }
            }

            using (Planted doc = Plant(s, Para))
            {
                if (!Ready(run, doc, "ts-x-p-via-codes"))
                {
                    return;
                }

                if (Want("ts-x-p-via-codes"))
                {
                    string code = doc.Paragraph(Para);
                    string sentence = doc.Sentence(Para);
                    ToolResult applied = s.ApplyDocumentFormat(Flat(doc.Id, "target_codes", code, "[1]", Super()));
                    Ok(run, "ts-x-p-via-codes",
                        applied.Success && OnlySuper(s, doc, sentence, false, "[1]"),
                        applied.Success ? "target_codes 里的段码没有只套 [1]" : (applied.Error ?? ""));
                }
            }

            using (Planted doc = Plant(s, ParaTwoMarks))
            {
                if (!Ready(run, doc, "ts-x-p-index2"))
                {
                    return;
                }

                if (Want("ts-x-p-index2"))
                {
                    string code = doc.Paragraph(ParaTwoMarks);
                    string sentence = doc.Sentence(ParaTwoMarks);
                    ToolResult applied = s.ApplyDocumentFormat(
                        Flat(doc.Id, "target_paragraph_codes", code, TextObj("[1]", 2), Super()));
                    JObject entry = Entry(s, doc, sentence, false);
                    Ok(run, "ts-x-p-index2",
                        applied.Success && NthMarkOnly(entry, "[1]", 2),
                        applied.Success ? "段内第二处 [1] 没有单独上标: " + Dump(entry) : (applied.Error ?? ""));
                }
            }

            using (Planted doc = Plant(s, ParaTwoMarks))
            {
                if (!Ready(run, doc, "ts-x-p-amb"))
                {
                    return;
                }

                if (Want("ts-x-p-amb"))
                {
                    string code = doc.Paragraph(ParaTwoMarks);
                    string sentence = doc.Sentence(ParaTwoMarks);
                    ToolResult applied = s.ApplyDocumentFormat(Flat(doc.Id, "target_paragraph_codes", code, "[1]", Super()));
                    Ok(run, "ts-x-p-amb",
                        Failed(applied, "出现 2 次，有歧义") && !AnySuper(s, doc, sentence, false),
                        applied != null && !applied.Success ? (applied.Error ?? "") : "段内两处 [1] 仍套上了");
                }
            }

            using (Planted doc = Plant(s, Para))
            {
                if (!Ready(run, doc, "ts-x-p-miss"))
                {
                    return;
                }

                if (Want("ts-x-p-miss"))
                {
                    string code = doc.Paragraph(Para);
                    ToolResult applied = s.ApplyDocumentFormat(Flat(doc.Id, "target_paragraph_codes", code, "[9]", Super()));
                    Ok(run, "ts-x-p-miss",
                        Failed(applied, "没有") && !AnySuper(s, doc, doc.Sentence(Para), false),
                        applied != null && !applied.Success ? (applied.Error ?? "") : "段内找不到仍套上了");
                }
            }

            using (Planted doc = Plant(s, Para))
            {
                if (!Ready(run, doc, "ts-x-p-obj"))
                {
                    return;
                }

                if (Want("ts-x-p-obj"))
                {
                    string code = doc.Paragraph(Para);
                    ToolResult applied = s.ApplyDocumentFormat(
                        Flat(doc.Id, "target_paragraph_codes", code, TextObj("[1]", 1), Super()));
                    Ok(run, "ts-x-p-obj",
                        applied.Success && OnlySuper(s, doc, doc.Sentence(Para), false, "[1]"),
                        applied.Success ? "段内对象 index 1 没有只套 [1]" : (applied.Error ?? ""));
                }
            }
        }

        private static void RunItemsExtra(TestRun run, WdSession s)
        {
            using (Planted doc = Plant(s, One))
            {
                if (!Ready(run, doc, "ts-x-items-both"))
                {
                    return;
                }

                if (Want("ts-x-items-both"))
                {
                    string code = doc.Sentence(One);
                    ToolResult applied = s.ApplyDocumentFormat(Items(
                        doc.Id,
                        Item("target_codes", code, "[1]", Super()),
                        Item("target_codes", code, "说明", Super())));
                    JObject entry = Entry(s, doc, code, false);
                    Ok(run, "ts-x-items-both",
                        applied.Success && OnlyTheseSuper(entry, "[1]", "说明"),
                        applied.Success ? "两项没有分别套上: " + Dump(entry) : (applied.Error ?? ""));
                }
            }

            using (Planted doc = Plant(s, One))
            {
                if (!Ready(run, doc, "ts-x-items-none"))
                {
                    return;
                }

                if (Want("ts-x-items-none"))
                {
                    string code = doc.Sentence(One);
                    ToolResult applied = s.ApplyDocumentFormat(Items(
                        doc.Id,
                        Item("target_codes", code, "[8]", Super()),
                        Item("target_codes", code, "[9]", Super())));
                    Ok(run, "ts-x-items-none",
                        Failed(applied, "没有") && !AnySuper(s, doc, code, false),
                        applied != null && !applied.Success ? (applied.Error ?? "") : "两项都失败仍套上了");
                }
            }

            using (Planted doc = Plant(s, One))
            {
                if (!Ready(run, doc, "ts-x-items-over"))
                {
                    return;
                }

                if (Want("ts-x-items-over"))
                {
                    string code = doc.Sentence(One);
                    ToolResult applied = s.ApplyDocumentFormat(Items(
                        doc.Id,
                        Item("target_codes", code, TextObj("[1]", 9), Super()),
                        Item("target_codes", code, "[1]", Super())));
                    Ok(run, "ts-x-items-over",
                        applied.Success && WarnHas(applied, "超出范围") && OnlySuper(s, doc, code, false, "[1]"),
                        applied.Success ? "越界项没有留下警告或另一项没套上" : (applied.Error ?? ""));
                }
            }

            using (Planted doc = Plant(s, Jia + "\r" + Yi))
            {
                if (!Ready(run, doc, "ts-x-other-sent"))
                {
                    return;
                }

                if (Want("ts-x-other-sent"))
                {
                    string later = doc.Sentence(Yi);
                    string earlier = doc.Sentence(Jia);
                    ToolResult applied = s.ApplyDocumentFormat(Flat(doc.Id, "target_codes", later, "[1]", Super()));
                    Ok(run, "ts-x-other-sent",
                        applied.Success
                            && OnlySuper(s, doc, later, false, "[1]")
                            && !AnySuper(s, doc, earlier, false),
                        applied.Success ? "前一句的 [1] 也被套上了" : (applied.Error ?? ""));
                }
            }

            using (Planted doc = Plant(s, One))
            {
                if (!Ready(run, doc, "ts-x-bold"))
                {
                    return;
                }

                if (Want("ts-x-bold"))
                {
                    string code = doc.Sentence(One);
                    var bold = new Dictionary<string, object> { ["is_bold"] = true };
                    ToolResult applied = s.ApplyDocumentFormat(Flat(doc.Id, "target_codes", code, "[1]", bold));
                    Ok(run, "ts-x-bold",
                        applied.Success && OnlyBold(s, doc, code, "[1]"),
                        applied.Success ? "加粗没有只落在 [1]" : (applied.Error ?? ""));
                }
            }
        }

        private static void RunWriteExtra(TestRun run, WdSession s)
        {
            using (Planted doc = Plant(s, Two))
            {
                if (!Ready(run, doc, "ts-x-rep-index"))
                {
                    return;
                }

                if (Want("ts-x-rep-index"))
                {
                    string code = doc.Sentence(Two);
                    ToolResult replaced = s.ProcessActions(doc.Id, ReplaceSlice(code, TextObj("[1]", 2), "[2]"));
                    Ok(run, "ts-x-rep-index",
                        replaced.Success && doc.Has("见方法[1]与结果[2]") && !doc.Has("结果[1]"),
                        replaced.Success ? "第二处没有换成 [2]" : (replaced.Error ?? ""));
                }
            }

            using (Planted doc = Plant(s, One))
            {
                if (!Ready(run, doc, "ts-x-rep-miss"))
                {
                    return;
                }

                if (Want("ts-x-rep-miss"))
                {
                    string code = doc.Sentence(One);
                    ToolResult replaced = s.ProcessActions(doc.Id, ReplaceSlice(code, "[9]", "[2]"));
                    Ok(run, "ts-x-rep-miss",
                        Failed(replaced, "没有") && doc.Has(One) && !doc.Has("[2]"),
                        replaced != null && !replaced.Success ? (replaced.Error ?? "") : "找不到仍改了正文");
                }
            }

            using (Planted doc = Plant(s, Two))
            {
                if (!Ready(run, doc, "ts-x-ins-index"))
                {
                    return;
                }

                if (Want("ts-x-ins-index"))
                {
                    string code = doc.Sentence(Two);
                    ToolResult inserted = s.ProcessActions(doc.Id, InsertAfterSlice(code, TextObj("[1]", 1), "注"));
                    Ok(run, "ts-x-ins-index",
                        inserted.Success && doc.Has("见方法[1]注与结果[1]"),
                        inserted.Success ? "没有插在第一处 [1] 后面" : (inserted.Error ?? ""));
                }
            }

            using (Planted doc = Plant(s, One))
            {
                if (!Ready(run, doc, "ts-x-ins-miss"))
                {
                    return;
                }

                if (Want("ts-x-ins-miss"))
                {
                    string code = doc.Sentence(One);
                    ToolResult inserted = s.ProcessActions(doc.Id, InsertAfterSlice(code, "[9]", "不应出现"));
                    Ok(run, "ts-x-ins-miss",
                        Failed(inserted, "没有") && doc.Has(One) && !doc.Has("不应出现"),
                        inserted != null && !inserted.Success ? (inserted.Error ?? "") : "找不到仍插入了");
                }
            }

            using (Planted doc = Plant(s, Para))
            {
                if (!Ready(run, doc, "ts-x-p-rep"))
                {
                    return;
                }

                if (Want("ts-x-p-rep"))
                {
                    string code = doc.Paragraph(Para);
                    ToolResult replaced = s.ProcessActions(doc.Id, ReplaceParagraphSlice(code, "[1]", "[2]"));
                    Ok(run, "ts-x-p-rep",
                        replaced.Success && doc.Has("段前甲[2]段后") && !doc.Has("[1]"),
                        replaced.Success ? "段内 [1] 没有换成 [2]" : (replaced.Error ?? ""));
                }
            }
        }

        private static void Slice(TestRun run, WdSession s, string id, string plain, string text)
        {
            using (Planted doc = Plant(s, plain))
            {
                if (!Ready(run, doc, id))
                {
                    return;
                }

                if (!Want(id))
                {
                    return;
                }

                string code = doc.Sentence(plain);
                ToolResult applied = s.ApplyDocumentFormat(Flat(doc.Id, "target_codes", code, text, Super()));
                Ok(run, id,
                    applied.Success && OnlySuper(s, doc, code, false, text),
                    applied.Success ? "没有只套「" + text + "」" : (applied.Error ?? ""));
            }
        }

        private static bool Want(string id)
        {
            if (string.IsNullOrEmpty(IdFilter))
            {
                return true;
            }

            return id.StartsWith(IdFilter, StringComparison.OrdinalIgnoreCase);
        }

        private static Dictionary<string, object> TextValue(string value)
        {
            return new Dictionary<string, object> { ["value"] = value };
        }

        private static bool OnlyTheseSuper(JObject entry, params string[] exacts)
        {
            var need = new HashSet<string>(exacts);
            var seen = new HashSet<string>();
            foreach (JToken token in Runs(entry))
            {
                JObject run = token as JObject;
                if (run == null)
                {
                    continue;
                }

                string text = Txt(run, "text");
                bool super = Flag(Obj(run, "char_format"), "is_superscript");
                if (need.Contains(text))
                {
                    if (!super)
                    {
                        return false;
                    }

                    seen.Add(text);
                }
                else if (super)
                {
                    return false;
                }
            }

            return seen.Count == need.Count;
        }

        private static bool OnlyBold(WdSession s, Planted doc, string code, string exact)
        {
            JObject entry = Entry(s, doc, code, false);
            bool hit = false;
            foreach (JToken token in Runs(entry))
            {
                JObject run = token as JObject;
                if (run == null)
                {
                    continue;
                }

                bool bold = Flag(Obj(run, "char_format"), "is_bold");
                if (Txt(run, "text") == exact)
                {
                    if (!bold)
                    {
                        return false;
                    }

                    hit = true;
                }
                else if (bold)
                {
                    return false;
                }
            }

            return hit;
        }
    }
}
