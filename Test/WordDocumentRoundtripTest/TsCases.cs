using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using WordAddIn1;

namespace WordDocumentRoundtripTest
{
    /// <summary>
    /// 句子内按字符串定位。铺稿用 COM，抄当次 S_ / P_。
    /// </summary>
    internal static class TsCases
    {
        public const int Target = 18;

        private const string One = "见方法[1]与说明。";
        private const string Two = "见方法[1]与结果[1]。";
        private const string Jia = "甲句见[1]。";
        private const string Yi = "乙句见[1]。";
        private const string Para = "段前甲[1]段后。";
        private const string CellLong = "格内前乙。\r格内后乙完。";
        private const string CellOne = "前文乙句。\r乙";

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
            RunOtherTool(run, s);
            RunNoQuit(run, wordBefore, wpsBefore, wordWinBefore, wpsWinBefore);
        }

        private static void RunSentence(TestRun run, WdSession s)
        {
            using (Planted doc = Plant(s, One))
            {
                if (!Ready(run, doc, "ts-s-slice"))
                {
                    return;
                }

                string code = doc.Sentence(One);
                ToolResult applied = s.ApplyDocumentFormat(Flat(doc.Id, "target_codes", code, "[1]", Super()));
                Ok(run, "ts-s-slice",
                    applied.Success && OnlySuper(s, doc, code, false, "[1]") && doc.Has(One),
                    applied.Success ? "上标没有只落在 [1]" : (applied.Error ?? ""));
            }

            using (Planted doc = Plant(s, One))
            {
                if (!Ready(run, doc, "ts-s-whole"))
                {
                    return;
                }

                string code = doc.Sentence(One);
                ToolResult applied = s.ApplyDocumentFormat(Flat(doc.Id, "target_codes", code, null, Super()));
                Ok(run, "ts-s-whole",
                    applied.Success && AllSuper(s, doc, code, false) && doc.Has(One),
                    applied.Success ? "不传 text 时整句没有一起上标" : (applied.Error ?? ""));
            }

            using (Planted doc = Plant(s, One))
            {
                if (!Ready(run, doc, "ts-s-miss"))
                {
                    return;
                }

                string code = doc.Sentence(One);
                ToolResult applied = s.ApplyDocumentFormat(Flat(doc.Id, "target_codes", code, "[9]", Super()));
                Ok(run, "ts-s-miss",
                    Failed(applied, "没有") && !AnySuper(s, doc, code, false) && doc.Has(One),
                    applied != null && !applied.Success ? (applied.Error ?? "") : "找不到仍套上了");
            }

            using (Planted doc = Plant(s, Jia + "\r" + Yi))
            {
                if (!Ready(run, doc, "ts-s-two"))
                {
                    return;
                }

                string a = doc.Sentence(Jia);
                string b = doc.Sentence(Yi);
                ToolResult applied = s.ApplyDocumentFormat(Flat(doc.Id, "target_codes", a + "," + b, "[1]", Super()));
                Ok(run, "ts-s-two",
                    Failed(applied, "恰好一个") && !AnySuper(s, doc, a, false) && !AnySuper(s, doc, b, false),
                    applied != null && !applied.Success ? (applied.Error ?? "") : "两个码带 text 仍套上了");
            }

            using (Planted doc = Plant(s, Two))
            {
                if (!Ready(run, doc, "ts-s-ambiguous", "ts-s-index2", "ts-s-index1", "ts-s-index-over"))
                {
                    return;
                }

                string code = doc.Sentence(Two);
                ToolResult ambiguous = s.ApplyDocumentFormat(Flat(doc.Id, "target_codes", code, "[1]", Super()));
                Ok(run, "ts-s-ambiguous",
                    Failed(ambiguous, "出现 2 次，有歧义")
                        && Failed(ambiguous, "请用 text 里的 index 指定目标")
                        && !AnySuper(s, doc, code, false),
                    ambiguous != null && !ambiguous.Success ? (ambiguous.Error ?? "") : "两处 [1] 仍套上了");
            }

            using (Planted doc = Plant(s, Two))
            {
                if (!Ready(run, doc, "ts-s-index2"))
                {
                    return;
                }

                string code = doc.Sentence(Two);
                ToolResult applied = s.ApplyDocumentFormat(Flat(doc.Id, "target_codes", code, TextObj("[1]", 2), Super()));
                JObject entry = Entry(s, doc, code, false);
                Ok(run, "ts-s-index2",
                    applied.Success && NthMarkOnly(entry, "[1]", 2),
                    applied.Success ? "第二处 [1] 没有单独上标: " + Dump(entry) : (applied.Error ?? ""));
            }

            using (Planted doc = Plant(s, Two))
            {
                if (!Ready(run, doc, "ts-s-index1"))
                {
                    return;
                }

                string code = doc.Sentence(Two);
                ToolResult applied = s.ApplyDocumentFormat(Flat(doc.Id, "target_codes", code, TextObj("[1]", 1), Super()));
                JObject entry = Entry(s, doc, code, false);
                Ok(run, "ts-s-index1",
                    applied.Success && NthMarkOnly(entry, "[1]", 1),
                    applied.Success ? "第一处 [1] 没有单独上标: " + Dump(entry) : (applied.Error ?? ""));
            }

            using (Planted doc = Plant(s, Two))
            {
                if (!Ready(run, doc, "ts-s-index-over"))
                {
                    return;
                }

                string code = doc.Sentence(Two);
                ToolResult applied = s.ApplyDocumentFormat(Flat(doc.Id, "target_codes", code, TextObj("[1]", 3), Super()));
                Ok(run, "ts-s-index-over",
                    Failed(applied, "index 超出范围") && !AnySuper(s, doc, code, false),
                    applied != null && !applied.Success ? (applied.Error ?? "") : "越界仍套上了");
            }

            using (Planted doc = Plant(s, One))
            {
                if (!Ready(run, doc, "ts-s-sibling"))
                {
                    return;
                }

                string code = doc.Sentence(One);
                var args = Flat(doc.Id, "target_codes", code, "[1]", Super());
                args["index"] = 1;
                ToolResult applied = s.ApplyDocumentFormat(args);
                Ok(run, "ts-s-sibling",
                    Failed(applied, "text 对象里面") && !AnySuper(s, doc, code, false),
                    applied != null && !applied.Success ? (applied.Error ?? "") : "并列 index 仍套上了");
            }

            using (Planted doc = Plant(s, One))
            {
                if (!Ready(run, doc, "ts-items-index"))
                {
                    return;
                }

                string code = doc.Sentence(One);
                var bad = new Dictionary<string, object>
                {
                    ["target_codes"] = code,
                    ["text"] = "[1]",
                    ["char_format"] = Super(),
                    ["index"] = 1
                };
                ToolResult applied = s.ApplyDocumentFormat(Items(doc.Id, bad));
                Ok(run, "ts-items-index",
                    Failed(applied, "text 对象里面") && !AnySuper(s, doc, code, false),
                    applied != null && !applied.Success ? (applied.Error ?? "") : "项上的 index 仍套上了");
            }

            using (Planted doc = Plant(s, One))
            {
                if (!Ready(run, doc, "ts-items-warn"))
                {
                    return;
                }

                string code = doc.Sentence(One);
                ToolResult applied = s.ApplyDocumentFormat(Items(
                    doc.Id,
                    Item("target_codes", code, "[9]", Super()),
                    Item("target_codes", code, "[1]", Super())));
                Ok(run, "ts-items-warn",
                    applied.Success
                        && WarnHas(applied, "没有")
                        && OnlySuper(s, doc, code, false, "[1]"),
                    applied.Success ? "一项失败时另一项没有套上 [1]" : (applied.Error ?? ""));
            }
        }

        private static void RunParagraph(TestRun run, WdSession s)
        {
            using (Planted doc = Plant(s, Para))
            {
                if (!Ready(run, doc, "ts-p-slice"))
                {
                    return;
                }

                string code = doc.Paragraph(Para);
                string sentence = doc.Sentence(Para);
                ToolResult applied = s.ApplyDocumentFormat(Flat(doc.Id, "target_paragraph_codes", code, "[1]", Super()));
                Ok(run, "ts-p-slice",
                    applied.Success && OnlySuper(s, doc, sentence, false, "[1]") && doc.Has(Para),
                    applied.Success ? "段落里的 [1] 没有单独上标: " + Dump(Entry(s, doc, sentence, false)) : (applied.Error ?? ""));
            }

            using (Planted doc = PlantTable(s, CellLong))
            {
                if (!Ready(run, doc, "ts-p-cell"))
                {
                    return;
                }

                string early = doc.Paragraph("格内前乙");
                string late = doc.Paragraph("格内后乙完");
                string earlySentence = doc.Sentence("格内前乙");
                string lateSentence = doc.Sentence("格内后乙完");
                if (string.IsNullOrEmpty(early) || string.IsNullOrEmpty(late)
                    || string.IsNullOrEmpty(earlySentence) || string.IsNullOrEmpty(lateSentence))
                {
                    run.CaseFail("ts-p-cell", "抄不到格内两段");
                    return;
                }

                ToolResult applied = s.ApplyDocumentFormat(Flat(doc.Id, "target_paragraph_codes", late, "乙", Super()));
                Ok(run, "ts-p-cell",
                    applied.Success
                        && OnlySuper(s, doc, lateSentence, false, "乙")
                        && !AnySuper(s, doc, earlySentence, false),
                    applied.Success ? "格末段落的「乙」没有单独上标: " + Dump(Entry(s, doc, lateSentence, false)) : (applied.Error ?? ""));
            }

            using (Planted doc = PlantTable(s, CellOne))
            {
                if (!Ready(run, doc, "ts-p-one"))
                {
                    return;
                }

                string early = doc.Paragraph("前文乙句");
                string late = doc.ParagraphExact("乙");
                if (string.IsNullOrEmpty(early) || string.IsNullOrEmpty(late))
                {
                    run.CaseFail("ts-p-one", "抄不到单字段");
                    return;
                }

                ToolResult applied = s.ApplyDocumentFormat(Flat(doc.Id, "target_paragraph_codes", late, "乙", Super()));
                int lateState;
                int earlyState;
                bool lateRead = s.TryParagraphSuper(doc.Path, "乙", true, out lateState);
                bool earlyRead = s.TryParagraphSuper(doc.Path, "前文乙句", false, out earlyState);
                Ok(run, "ts-p-one",
                    applied.Success && lateRead && earlyRead && lateState == 1 && earlyState == 0,
                    applied.Success
                        ? "只有一个「乙」的格末段没有上标: late=" + lateState + " early=" + earlyState
                        : (applied.Error ?? ""));
            }
        }

        private static void RunOtherTool(TestRun run, WdSession s)
        {
            using (Planted doc = Plant(s, One))
            {
                if (!Ready(run, doc, "ts-act-replace"))
                {
                    return;
                }

                string code = doc.Sentence(One);
                ToolResult replaced = s.ProcessActions(doc.Id, ReplaceSlice(code, "[1]", "[2]"));
                Ok(run, "ts-act-replace",
                    replaced.Success && doc.Has("见方法[2]与说明。") && !doc.Has("[1]"),
                    replaced.Success ? "改字没有只换掉 [1]" : (replaced.Error ?? ""));
            }

            using (Planted doc = Plant(s, One))
            {
                if (!Ready(run, doc, "ts-act-insert"))
                {
                    return;
                }

                string code = doc.Sentence(One);
                ToolResult inserted = s.ProcessActions(doc.Id, InsertAfterSlice(code, "[1]", "注"));
                Ok(run, "ts-act-insert",
                    inserted.Success && doc.Has("见方法[1]注与说明。"),
                    inserted.Success ? "插入没有落在 [1] 后面" : (inserted.Error ?? ""));
            }

            using (Planted doc = PlantTable(s, CellLong))
            {
                if (!Ready(run, doc, "ts-act-table"))
                {
                    return;
                }

                if (string.IsNullOrEmpty(doc.TableId))
                {
                    run.CaseFail("ts-act-table", "抄不到表号");
                    return;
                }

                ToolResult inserted = s.ProcessActions(doc.Id, InsertAt(new Dictionary<string, object>
                {
                    ["table"] = doc.TableId,
                    ["text"] = "乙"
                }, "不应出现"));
                Ok(run, "ts-act-table",
                    Failed(inserted, "恰好一个") && !doc.Has("不应出现"),
                    inserted != null && !inserted.Success ? (inserted.Error ?? "") : "表带 text 仍插入了");
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
            Ok(run, "ts-no-quit", wordOk && wpsOk, wordDetail + " " + wpsDetail);
        }

        private static Dictionary<string, object> Flat(
            string id,
            string codeKey,
            string codes,
            object text,
            Dictionary<string, object> format)
        {
            var args = new Dictionary<string, object>
            {
                ["channel_id"] = id,
                [codeKey] = codes,
                ["char_format"] = format
            };
            if (text != null)
            {
                args["text"] = text;
            }

            return args;
        }

        private static Dictionary<string, object> Items(string id, params Dictionary<string, object>[] items)
        {
            return new Dictionary<string, object>
            {
                ["channel_id"] = id,
                ["items"] = items
            };
        }

        private static Dictionary<string, object> Item(string codeKey, string codes, object text, Dictionary<string, object> format)
        {
            return new Dictionary<string, object>
            {
                [codeKey] = codes,
                ["text"] = text,
                ["char_format"] = format
            };
        }

        private static Dictionary<string, object> TextObj(string value, int index)
        {
            return new Dictionary<string, object>
            {
                ["value"] = value,
                ["index"] = index
            };
        }

        private static Dictionary<string, object> Super()
        {
            return new Dictionary<string, object> { ["is_superscript"] = true };
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

        private static List<Dictionary<string, object>> ReplaceSlice(string code, string text, string neu)
        {
            return new List<Dictionary<string, object>>
            {
                new Dictionary<string, object>
                {
                    ["type"] = "replace",
                    ["detail"] = new Dictionary<string, object>
                    {
                        ["original"] = new Dictionary<string, object>
                        {
                            ["codes"] = code,
                            ["text"] = text
                        },
                        ["new"] = neu,
                        ["format"] = DefaultFormat()
                    }
                }
            };
        }

        private static List<Dictionary<string, object>> InsertAfterSlice(string code, string text, string insert)
        {
            return InsertAt(new Dictionary<string, object>
            {
                ["codes"] = code,
                ["text"] = text
            }, insert);
        }

        private static List<Dictionary<string, object>> InsertAt(Dictionary<string, object> target, string insert)
        {
            return new List<Dictionary<string, object>>
            {
                new Dictionary<string, object>
                {
                    ["type"] = "insert",
                    ["detail"] = new Dictionary<string, object>
                    {
                        ["target"] = target,
                        ["insert"] = insert,
                        ["format"] = DefaultFormat()
                    }
                }
            };
        }

        private static bool NthMarkOnly(JObject entry, string exact, int index)
        {
            int seen = 0;
            bool hit = false;
            foreach (JToken token in Runs(entry))
            {
                JObject run = token as JObject;
                if (run == null)
                {
                    continue;
                }

                string text = Txt(run, "text");
                bool super = Flag(Obj(run, "char_format"), "is_superscript");
                int count = CountOf(text, exact);
                if (super)
                {
                    if (text != exact || seen != index - 1)
                    {
                        return false;
                    }

                    hit = true;
                }

                seen += count;
            }

            return hit && seen >= index;
        }

        private static int CountOf(string text, string exact)
        {
            if (string.IsNullOrEmpty(text) || string.IsNullOrEmpty(exact))
            {
                return 0;
            }

            int count = 0;
            int at = 0;
            while (at <= text.Length - exact.Length)
            {
                int found = text.IndexOf(exact, at, StringComparison.Ordinal);
                if (found < 0)
                {
                    break;
                }

                count++;
                at = found + exact.Length;
            }

            return count;
        }

        private static string StripMarks(string text)
        {
            if (string.IsNullOrEmpty(text))
            {
                return "";
            }

            return text.Trim().TrimEnd('\r', '\a', '\n', '\v', '\f');
        }

        private static bool OnlySuper(WdSession s, Planted doc, string code, bool paragraph, string exact)
        {
            JObject entry = Entry(s, doc, code, paragraph);
            List<bool> marks = SuperFlags(entry, exact);
            return marks.Count >= 1 && !marks.Contains(false) && OthersPlain(entry, exact);
        }

        private static bool AllSuper(WdSession s, Planted doc, string code, bool paragraph)
        {
            JObject entry = Entry(s, doc, code, paragraph);
            JArray runs = Runs(entry);
            if (runs.Count == 0)
            {
                return Flag(Obj(entry, "char_format"), "is_superscript");
            }

            foreach (JToken token in runs)
            {
                if (!Flag(Obj(token as JObject, "char_format"), "is_superscript"))
                {
                    return false;
                }
            }

            return true;
        }

        private static bool AnySuper(WdSession s, Planted doc, string code, bool paragraph)
        {
            JObject entry = Entry(s, doc, code, paragraph);
            if (Flag(Obj(entry, "char_format"), "is_superscript"))
            {
                return true;
            }

            foreach (JToken token in Runs(entry))
            {
                if (Flag(Obj(token as JObject, "char_format"), "is_superscript"))
                {
                    return true;
                }
            }

            return false;
        }

        private static List<bool> SuperFlags(JObject entry, string exact)
        {
            var marks = new List<bool>();
            foreach (JToken token in Runs(entry))
            {
                JObject run = token as JObject;
                if (run == null || Txt(run, "text") != exact)
                {
                    continue;
                }

                marks.Add(Flag(Obj(run, "char_format"), "is_superscript"));
            }

            return marks;
        }

        private static bool OthersPlain(JObject entry, string exact)
        {
            foreach (JToken token in Runs(entry))
            {
                JObject run = token as JObject;
                if (run == null || Txt(run, "text") == exact)
                {
                    continue;
                }

                if (Flag(Obj(run, "char_format"), "is_superscript"))
                {
                    return false;
                }
            }

            return true;
        }

        private static JObject Entry(WdSession s, Planted doc, string code, bool paragraph)
        {
            if (doc == null || string.IsNullOrEmpty(code))
            {
                return null;
            }

            var args = new Dictionary<string, object> { ["channel_id"] = doc.Id };
            args[paragraph ? "target_paragraph_codes" : "target_codes"] = code;
            return FirstEntry(s.GetFormatContext(args));
        }

        private static bool Failed(ToolResult result, string needle)
        {
            if (result == null || result.Success)
            {
                return false;
            }

            string blob = (result.Error ?? "") + JsonConvert.SerializeObject(result.Data);
            return blob.IndexOf(needle, StringComparison.Ordinal) >= 0;
        }

        private static bool WarnHas(ToolResult result, string needle)
        {
            JArray warnings = Arr(result, "warnings");
            if (warnings == null)
            {
                return false;
            }

            foreach (JToken token in warnings)
            {
                if (token.ToString().IndexOf(needle, StringComparison.Ordinal) >= 0)
                {
                    return true;
                }
            }

            return false;
        }

        private static Planted Plant(WdSession s, string plain)
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
            }

            var doc = new Planted { Path = path, Id = id, Session = s };
            doc.Reread(s);
            return doc;
        }

        private static Planted PlantTable(WdSession s, string cell)
        {
            string path = s.NewDocxPath();
            ToolResult opened = s.Open(path, true, s.App);
            if (!opened.Success)
            {
                return new Planted { Error = opened.Error ?? "新建失败" };
            }

            string id = WdSession.Str(opened, "channel_id");
            if (!s.InsertTable(path, 1, 1, new[] { cell }))
            {
                s.CloseDocByPath(path);
                return new Planted { Error = "COM 建表失败" };
            }

            var doc = new Planted { Path = path, Id = id, Session = s };
            doc.Reread(s);
            return doc;
        }

        private static bool Ready(TestRun run, Planted doc, params string[] ids)
        {
            if (doc != null && string.IsNullOrEmpty(doc.Error) && !string.IsNullOrEmpty(ids[0]))
            {
                return true;
            }

            string reason = doc == null ? "铺稿失败" : (doc.Error ?? "铺稿失败");
            run.CaseFail(ids[0], reason);
            for (int i = 1; i < ids.Length; i++)
            {
                run.CaseFail(ids[i], reason);
            }

            return false;
        }

        private sealed class Planted : IDisposable
        {
            public string Path;
            public string Id;
            public WdSession Session;
            public string Error;
            public string TableId;
            public List<Marker> Sentences = new List<Marker>();
            public List<Marker> Paragraphs = new List<Marker>();

            public string Sentence(string needle)
            {
                return Nth(Sentences, needle, 0, false);
            }

            public string Paragraph(string needle)
            {
                return Nth(Paragraphs, needle, 0, false);
            }

            public string ParagraphExact(string text)
            {
                return Nth(Paragraphs, text, 0, true);
            }

            public bool Has(string needle)
            {
                ToolResult got = Session.GetContent(Id);
                string text = got.Success ? WdSession.Str(got, "document_content") : "";
                return text.IndexOf(needle, StringComparison.Ordinal) >= 0;
            }

            public void Reread(WdSession s)
            {
                ToolResult sentences = s.GetContent(Id);
                string sentenceText = sentences.Success ? WdSession.Str(sentences, "document_content") : "";
                Sentences = Parse(sentenceText, SentencePattern);
                Match table = TablePattern.Match(sentenceText ?? "");
                TableId = table.Success ? table.Groups[1].Value : "";
                ToolResult paragraphs = s.GetContent(Id, "paragraph");
                Paragraphs = Parse(
                    paragraphs.Success ? WdSession.Str(paragraphs, "document_content") : "",
                    ParagraphPattern);
            }

            public void Dispose()
            {
                if (Session != null && !string.IsNullOrEmpty(Path))
                {
                    Session.CloseDocByPath(Path);
                }
            }

            private static string Nth(List<Marker> markers, string needle, int nth, bool exact)
            {
                int seen = 0;
                foreach (Marker marker in markers)
                {
                    string body = StripMarks(Regex.Replace(
                        marker.Text ?? "",
                        @"\[in_table=[^\]]+\]",
                        ""));
                    bool hit = exact
                        ? string.Equals(body, needle, StringComparison.Ordinal)
                        : body.IndexOf(needle, StringComparison.Ordinal) >= 0;
                    if (!hit)
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

        private sealed class Marker
        {
            public string Code;
            public string Text;
        }

        private static readonly Regex SentencePattern = new Regex(
            @"\[(S_[0-9A-Za-z]+)-start\]([\s\S]*?)\[\1-end\]",
            RegexOptions.Compiled);

        private static readonly Regex ParagraphPattern = new Regex(
            @"\[(P_[0-9A-Za-z]+(?:_[1-9][0-9]*)?)(?:-start| start)\]([\s\S]*?)\[\1(?:-end| end)\]",
            RegexOptions.Compiled);

        private static readonly Regex TablePattern = new Regex(
            @"\[in_table=(T_[0-9A-Za-z]+)\]",
            RegexOptions.Compiled);

        private static List<Marker> Parse(string text, Regex pattern)
        {
            var list = new List<Marker>();
            if (string.IsNullOrEmpty(text))
            {
                return list;
            }

            foreach (Match match in pattern.Matches(text))
            {
                list.Add(new Marker
                {
                    Code = match.Groups[1].Value,
                    Text = match.Groups[2].Value ?? ""
                });
            }

            return list;
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

        private static JArray Arr(ToolResult result, string key)
        {
            JObject data = Data(result);
            return data == null ? null : data[key] as JArray;
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

        private static JArray Runs(JObject entry)
        {
            return entry?["runs"] as JArray ?? new JArray();
        }

        private static JObject Obj(JObject parent, string key)
        {
            return parent == null ? null : parent[key] as JObject;
        }

        private static string Txt(JObject obj, string key)
        {
            if (obj == null || obj[key] == null || obj[key].Type == JTokenType.Null)
            {
                return "";
            }

            return obj[key].ToString();
        }

        private static bool Flag(JObject obj, string key)
        {
            if (obj == null || obj[key] == null || obj[key].Type == JTokenType.Null)
            {
                return false;
            }

            if (obj[key].Type == JTokenType.Boolean)
            {
                return obj[key].Value<bool>();
            }

            string value = obj[key].ToString();
            return string.Equals(value, "true", StringComparison.OrdinalIgnoreCase) || value == "1";
        }

        private static string Dump(JObject obj)
        {
            return obj == null ? "" : obj.ToString(Formatting.None);
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
    }
}
