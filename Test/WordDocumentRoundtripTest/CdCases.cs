using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using WordAddIn1;

namespace WordDocumentRoundtripTest
{
    /// <summary>
    /// 写字时没传的开关写成关掉。对照 docs/测试/word测试/写入未指定字符格式按默认。
    /// </summary>
    internal static class CdCases
    {
        public const int Target = 21;

        private static readonly string[] Ids =
        {
            "cd-super-head", "cd-super-cite", "cd-super-tail",
            "cd-insert-tail", "cd-replace-clear",
            "cd-sub-tail", "cd-italic-tail", "cd-bold-tail", "cd-underline-tail",
            "cd-strike-tail", "cd-shadow-tail", "cd-caps-tail", "cd-rest-tail",
            "cd-color-auto", "cd-color-keep",
            "cd-font-keep", "cd-bg-yellow", "cd-uline-style", "cd-noproof-ok", "cd-apply-keep",
            "cd-no-quit"
        };

        private static readonly string[] AlwaysOff =
        {
            "is_bold", "is_italic", "has_underline", "has_strikethrough",
            "is_superscript", "is_subscript"
        };

        private static readonly string[] RestKeys =
        {
            "outline", "emboss", "engrave", "small_caps", "double_strikethrough",
            "bold_bi", "italic_bi", "contextual_alternates", "disable_character_space_grid"
        };

        private const string Head = "误报较多";
        private const string Cite = "[1]";
        private const string Tail = "张敏和陈涛采用模型";

        public static string IdFilter = "cd-";

        public static int Count(string filter)
        {
            int n = 0;
            foreach (string id in Ids)
            {
                if (Want(id, filter))
                {
                    n++;
                }
            }

            return n;
        }

        public static void Run(
            TestRun run,
            string host,
            string filter,
            int wordBefore,
            int wpsBefore,
            bool wordWinBefore,
            bool wpsWinBefore)
        {
            IdFilter = string.IsNullOrEmpty(filter) ? "cd-" : filter;
            var s = new WdSession(host);
            string warm = Warm(s);
            string path = null;
            try
            {
                if (!AnyWrite())
                {
                    RunNoQuit(run, wordBefore, wpsBefore, wordWinBefore, wpsWinBefore);
                    return;
                }

                if (!Open(s, PlantText(), out path, out string id))
                {
                    FailWanted(run, "铺稿失败");
                    return;
                }

                if (!s.PaintNeedle(path, "已是上标句。", 0, Font("Superscript", -1), null)
                    || !s.PaintNeedle(path, "加粗套用句。", 0, Font("Bold", -1), null))
                {
                    FailWanted(run, "COM 铺上标或加粗失败");
                    return;
                }

                RunSuper(run, s, id);
                RunInsert(run, s, id);
                RunClear(run, s, id);
                RunSwitch(run, s, id, "cd-sub-tail", "下标句。", "is_subscript");
                RunSwitch(run, s, id, "cd-italic-tail", "斜体句。", "is_italic");
                RunSwitch(run, s, id, "cd-bold-tail", "加粗句。", "is_bold");
                RunSwitch(run, s, id, "cd-underline-tail", "下划句。", "has_underline");
                RunSwitch(run, s, id, "cd-strike-tail", "删除句。", "has_strikethrough");
                RunSwitch(run, s, id, "cd-shadow-tail", "阴影句。", "shadow");
                RunSwitch(run, s, id, "cd-caps-tail", "大写句。", "all_caps");
                RunRest(run, s, id);
                RunColor(run, s, id, "cd-color-auto", "红字句。", "色后自动", null);
                RunColor(run, s, id, "cd-color-keep", "蓝字句。", "色后蓝色", "#0000FF");
                RunFont(run, s, id);
                RunUline(run, s, id);
                RunNoProof(run, s, id);
                RunApply(run, s, id);
                RunNoQuit(run, wordBefore, wpsBefore, wordWinBefore, wpsWinBefore);
            }
            finally
            {
                if (!string.IsNullOrEmpty(path))
                {
                    s.CloseDocByPath(path);
                }

                if (!string.IsNullOrEmpty(warm))
                {
                    s.CloseDocByPath(warm);
                }
            }
        }

        private static string PlantText()
        {
            return string.Join("\r", new[]
            {
                "原句甲。", "锚句。", "已是上标句。",
                "下标句。", "斜体句。", "加粗句。", "下划句。", "删除句。", "阴影句。", "大写句。", "其余句。",
                "红字句。", "蓝字句。", "楷体句。", "线型句。", "成功句。", "加粗套用句。"
            });
        }

        private static void RunSuper(TestRun run, WdSession s, string id)
        {
            if (!Want("cd-super-head") && !Want("cd-super-cite") && !Want("cd-super-tail") && !Want("cd-bg-yellow"))
            {
                return;
            }

            string code = Code(s, id, "原句甲。");
            ToolResult wr = Replace(s, id, code, new[]
            {
                Piece(Head, Body()),
                Piece(Cite, Mark(Body(), "is_superscript")),
                Piece("。" + Tail + "。", Body())
            });
            if (!OkWrite(run, wr, "cd-super-head", "cd-super-cite", "cd-super-tail", "cd-bg-yellow"))
            {
                return;
            }

            JArray headRuns = Runs(s, id, Head);
            JArray citeRuns = Runs(s, id, Cite);
            JArray tailRuns = Runs(s, id, Tail);
            JObject head = Find(headRuns, Head);
            JObject cite = Find(citeRuns, Cite);
            JObject tail = Find(tailRuns, Tail);
            Check(run, "cd-super-head", Off(head, "is_superscript"), Dump(head));
            Check(run, "cd-super-cite", On(cite, "is_superscript"), Dump(cite));
            Check(run, "cd-super-tail", Off(tail, "is_superscript"), Dump(tail));
            Check(run, "cd-bg-yellow",
                Contains(Str(head, "highlight_color_index"), "Yellow"),
                Dump(head));
        }

        private static void RunInsert(TestRun run, WdSession s, string id)
        {
            if (!Want("cd-insert-tail"))
            {
                return;
            }

            string code = Code(s, id, "锚句。");
            ToolResult wr = Insert(s, id, code, new[]
            {
                Piece(Head, Body()),
                Piece(Cite, Mark(Body(), "is_superscript")),
                Piece("。" + Tail + "。", Body())
            });
            if (!OkWrite(run, wr, "cd-insert-tail"))
            {
                return;
            }

            JObject cite = Find(Runs(s, id, Cite), Cite);
            JObject tail = Find(Runs(s, id, Tail), Tail);
            Check(run, "cd-insert-tail",
                On(cite, "is_superscript") && Off(tail, "is_superscript"),
                "cite=" + Dump(cite) + " tail=" + Dump(tail));
        }

        private static void RunClear(TestRun run, WdSession s, string id)
        {
            if (!Want("cd-replace-clear"))
            {
                return;
            }

            string code = Code(s, id, "已是上标句。");
            ToolResult wr = Replace(s, id, code, new[] { Piece("新的正文。", Body()) });
            if (!OkWrite(run, wr, "cd-replace-clear"))
            {
                return;
            }

            JObject runCf = Find(Runs(s, id, "新的正文"), "新的正文");
            Check(run, "cd-replace-clear", Off(runCf, "is_superscript"), Dump(runCf));
        }

        private static void RunSwitch(TestRun run, WdSession s, string id, string caseId, string sentence, string key)
        {
            if (!Want(caseId))
            {
                return;
            }

            string code = Code(s, id, sentence);
            string stem = sentence.TrimEnd('。');
            string head = stem + "前";
            string tail = stem + "后";
            ToolResult wr = Replace(s, id, code, new[]
            {
                Piece(head, Mark(Body(), key)),
                Piece(tail + "。", Body())
            });
            if (!OkWrite(run, wr, caseId))
            {
                return;
            }

            JObject headCf = Find(Runs(s, id, head), head);
            JObject tailCf = Find(Runs(s, id, tail), tail);
            bool headOn = key == "shadow" ? On(headCf, key) || TextShadowOn(headCf) : On(headCf, key);
            bool tailOff = key == "shadow"
                ? !On(tailCf, key) && !TextShadowOn(tailCf)
                : (IsAlways(key) ? Off(tailCf, key) : !On(tailCf, key));
            Check(run, caseId, headOn && tailOff,
                "head=" + Dump(headCf) + " tail=" + Dump(tailCf));
        }

        private static void RunRest(TestRun run, WdSession s, string id)
        {
            if (!Want("cd-rest-tail"))
            {
                return;
            }

            string code = Code(s, id, "其余句。");
            var on = Body();
            foreach (string key in RestKeys)
            {
                on[key] = true;
            }

            ToolResult wr = Replace(s, id, code, new[]
            {
                Piece("其余前", on),
                Piece("其余后。", Body())
            });
            if (!OkWrite(run, wr, "cd-rest-tail"))
            {
                return;
            }

            JObject tail = Find(Runs(s, id, "其余后"), "其余后");
            var stuck = new List<string>();
            foreach (string key in RestKeys)
            {
                if (On(tail, key))
                {
                    stuck.Add(key);
                }
            }

            Check(run, "cd-rest-tail", stuck.Count == 0,
                stuck.Count == 0 ? Dump(tail) : string.Join(",", stuck) + " " + Dump(tail));
        }

        private static void RunColor(TestRun run, WdSession s, string id, string caseId, string sentence, string tailText, string keep)
        {
            if (!Want(caseId))
            {
                return;
            }

            string code = Code(s, id, sentence);
            Dictionary<string, object> tailFmt = Body();
            if (keep != null)
            {
                tailFmt["font_color"] = keep;
            }

            ToolResult wr = Replace(s, id, code, new[]
            {
                Piece(caseId + "前", MarkColor(Body(), "#FF0000")),
                Piece(tailText + "。", tailFmt)
            });
            if (!OkWrite(run, wr, caseId))
            {
                return;
            }

            JObject tail = Find(Runs(s, id, tailText), tailText);
            string color = Str(tail, "font_color");
            bool ok = keep == null
                ? string.Equals(color, "automatic", StringComparison.OrdinalIgnoreCase)
                : string.Equals(color, keep, StringComparison.OrdinalIgnoreCase);
            Check(run, caseId, ok, Dump(tail));
        }

        private static void RunFont(TestRun run, WdSession s, string id)
        {
            if (!Want("cd-font-keep"))
            {
                return;
            }

            string code = Code(s, id, "楷体句。");
            var kai = new Dictionary<string, object>
            {
                ["font_name"] = "楷体",
                ["font_size"] = 16,
                ["is_bold"] = false
            };
            ToolResult wr = Replace(s, id, code, new[]
            {
                Piece("楷体前", kai),
                Piece("楷体后。", new Dictionary<string, object> { ["is_bold"] = false })
            });
            if (!OkWrite(run, wr, "cd-font-keep"))
            {
                return;
            }

            JObject tail = Find(Runs(s, id, "楷体后"), "楷体后");
            string name = Str(tail, "font_name");
            int size = Num(tail, "font_size");
            bool nameOk = Contains(name, "楷") || Contains(name, "Kai");
            Check(run, "cd-font-keep", nameOk && size >= 15 && size <= 17, Dump(tail));
        }

        private static void RunUline(TestRun run, WdSession s, string id)
        {
            if (!Want("cd-uline-style"))
            {
                return;
            }

            string code = Code(s, id, "线型句。");
            ToolResult wr = Replace(s, id, code, new[]
            {
                Piece("线型有了。", new Dictionary<string, object> { ["has_underline"] = true })
            });
            if (!OkWrite(run, wr, "cd-uline-style"))
            {
                return;
            }

            JObject cf = Find(Runs(s, id, "线型有了"), "线型有了");
            string line = Str(cf, "underline");
            bool wiped = Contains(line, "None");
            Check(run, "cd-uline-style", On(cf, "has_underline") && !wiped, Dump(cf));
        }

        private static void RunNoProof(TestRun run, WdSession s, string id)
        {
            if (!Want("cd-noproof-ok"))
            {
                return;
            }

            string code = Code(s, id, "成功句。");
            ToolResult wr = Replace(s, id, code, new[] { Piece("写入成功句。", Body()) });
            Check(run, "cd-noproof-ok", wr != null && wr.Success, wr == null ? "null" : (wr.Error ?? "成功"));
        }

        private static void RunApply(TestRun run, WdSession s, string id)
        {
            if (!Want("cd-apply-keep"))
            {
                return;
            }

            string code = Code(s, id, "加粗套用句。");
            ToolResult wr = s.ApplyDocumentFormat(new Dictionary<string, object>
            {
                ["channel_id"] = id,
                ["target_codes"] = code,
                ["char_format"] = new Dictionary<string, object> { ["is_superscript"] = true }
            });
            if (!OkWrite(run, wr, "cd-apply-keep"))
            {
                return;
            }

            JObject cf = Find(Runs(s, id, "加粗套用句"), "加粗套用句");
            Check(run, "cd-apply-keep", On(cf, "is_superscript") && On(cf, "is_bold"), Dump(cf));
        }

        private static void RunNoQuit(
            TestRun run,
            int wordBefore,
            int wpsBefore,
            bool wordWinBefore,
            bool wpsWinBefore)
        {
            if (!Want("cd-no-quit"))
            {
                return;
            }

            bool wordOk = OfficeAlive.NotQuit("WINWORD", wordBefore, wordWinBefore, out string wordDetail);
            bool wpsOk = OfficeAlive.NotQuit("wps", wpsBefore, wpsWinBefore, out string wpsDetail);
            Check(run, "cd-no-quit", wordOk && wpsOk, wordDetail + " " + wpsDetail);
        }

        private static bool Open(WdSession s, string plain, out string path, out string id)
        {
            path = s.NewDocxPath();
            id = null;
            ToolResult opened = s.Open(path, true, s.App);
            if (!opened.Success)
            {
                return false;
            }

            id = WdSession.Str(opened, "channel_id");
            return s.TypePlain(path, plain);
        }

        private static string Warm(WdSession s)
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

            s.ActivateDoc(path);
            s.GetContent(WdSession.Str(opened, "channel_id"));
            return path;
        }

        private static ToolResult Replace(WdSession s, string id, string code, Dictionary<string, object>[] pieces)
        {
            if (string.IsNullOrEmpty(code))
            {
                return new ToolResult { Success = false, Error = "抄不到 S_" };
            }

            return s.ProcessActions(id, new List<Dictionary<string, object>>
            {
                new Dictionary<string, object>
                {
                    ["type"] = "replace",
                    ["detail"] = new Dictionary<string, object>
                    {
                        ["original"] = new Dictionary<string, object> { ["codes"] = code },
                        ["new"] = new List<Dictionary<string, object>>(pieces)
                    }
                }
            });
        }

        private static ToolResult Insert(WdSession s, string id, string code, Dictionary<string, object>[] pieces)
        {
            if (string.IsNullOrEmpty(code))
            {
                return new ToolResult { Success = false, Error = "抄不到 S_" };
            }

            return s.ProcessActions(id, new List<Dictionary<string, object>>
            {
                new Dictionary<string, object>
                {
                    ["type"] = "insert",
                    ["detail"] = new Dictionary<string, object>
                    {
                        ["target"] = new Dictionary<string, object> { ["codes"] = code },
                        ["insert"] = new List<Dictionary<string, object>>(pieces)
                    }
                }
            });
        }

        private static Dictionary<string, object> Piece(string text, Dictionary<string, object> format)
        {
            return new Dictionary<string, object>
            {
                ["text"] = text,
                ["format"] = new Dictionary<string, object> { ["char_format"] = format }
            };
        }

        private static Dictionary<string, object> Body()
        {
            return new Dictionary<string, object>
            {
                ["font_name"] = "宋体",
                ["font_size"] = 12,
                ["is_bold"] = false,
                ["name_bi"] = "Times New Roman"
            };
        }

        private static Dictionary<string, object> Mark(Dictionary<string, object> format, string key)
        {
            format[key] = true;
            return format;
        }

        private static Dictionary<string, object> MarkColor(Dictionary<string, object> format, string color)
        {
            format["font_color"] = color;
            return format;
        }

        private static Dictionary<string, object> Font(string key, object value)
        {
            return new Dictionary<string, object> { [key] = value };
        }

        private static string Code(WdSession s, string id, string sentence)
        {
            ToolResult got = s.GetContent(id);
            string body = got.Success ? WdSession.Str(got, "document_content") : "";
            int at = body.IndexOf(sentence, StringComparison.Ordinal);
            if (at < 0)
            {
                return "";
            }

            int mark = body.LastIndexOf("[S_", at, StringComparison.Ordinal);
            if (mark < 0)
            {
                return "";
            }

            Match m = Regex.Match(body.Substring(mark), @"\[S_([0-9a-zA-Z]{5})-start\]");
            return m.Success ? "S_" + m.Groups[1].Value : "";
        }

        private static JArray Runs(WdSession s, string id, string needle)
        {
            string code = Code(s, id, needle);
            JObject entry = null;
            if (!string.IsNullOrEmpty(code))
            {
                ToolResult got = s.GetFormatContext(new Dictionary<string, object>
                {
                    ["channel_id"] = id,
                    ["target_codes"] = code
                });
                JObject data = Data(got);
                JArray entries = data == null ? null : data["entries"] as JArray;
                if (entries != null && entries.Count > 0)
                {
                    entry = entries[0] as JObject;
                }
            }

            JArray runs = entry == null ? null : entry["runs"] as JArray;
            if (runs != null && runs.Count > 0)
            {
                return runs;
            }

            var one = new JObject();
            one["text"] = entry == null ? "" : (entry["text"] ?? entry["text_preview"] ?? "").ToString();
            one["char_format"] = entry == null ? new JObject() : (entry["char_format"] ?? new JObject());
            return new JArray(one);
        }

        private static JObject Find(JArray runs, string needle)
        {
            if (runs == null)
            {
                return null;
            }

            foreach (JToken token in runs)
            {
                JObject run = token as JObject;
                string text = run == null ? "" : (run["text"] ?? "").ToString();
                if (text.IndexOf(needle, StringComparison.Ordinal) >= 0)
                {
                    return run["char_format"] as JObject;
                }
            }

            return null;
        }

        private static bool TextShadowOn(JObject cf)
        {
            return cf != null && On(cf["text_shadow"] as JObject, "visible");
        }

        private static bool On(JObject cf, string key)
        {
            if (cf == null || cf[key] == null || cf[key].Type == JTokenType.Null)
            {
                return false;
            }

            JToken token = cf[key];
            if (token.Type == JTokenType.Boolean)
            {
                return token.Value<bool>();
            }

            return bool.TryParse(token.ToString(), out bool parsed) && parsed;
        }

        private static bool Off(JObject cf, string key)
        {
            return cf != null && cf[key] != null && cf[key].Type != JTokenType.Null && !On(cf, key);
        }

        private static bool IsAlways(string key)
        {
            foreach (string item in AlwaysOff)
            {
                if (item == key)
                {
                    return true;
                }
            }

            return false;
        }

        private static string Str(JObject cf, string key)
        {
            if (cf == null || cf[key] == null || cf[key].Type == JTokenType.Null)
            {
                return "";
            }

            return cf[key].ToString();
        }

        private static int Num(JObject cf, string key)
        {
            if (cf == null || cf[key] == null)
            {
                return int.MinValue;
            }

            try
            {
                return (int)Math.Round(cf[key].Value<double>());
            }
            catch
            {
                return int.TryParse(cf[key].ToString(), out int n) ? n : int.MinValue;
            }
        }

        private static bool Contains(string text, string needle)
        {
            return !string.IsNullOrEmpty(text)
                && text.IndexOf(needle, StringComparison.OrdinalIgnoreCase) >= 0;
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

        private static string Dump(JToken token)
        {
            return token == null ? "空" : token.ToString(Formatting.None);
        }

        private static bool OkWrite(TestRun run, ToolResult wr, params string[] ids)
        {
            if (wr != null && wr.Success)
            {
                return true;
            }

            string reason = wr == null ? "null" : (wr.Error ?? "失败");
            foreach (string id in ids)
            {
                if (Want(id))
                {
                    run.CaseFail(id, reason);
                }
            }

            return false;
        }

        private static void Check(TestRun run, string id, bool ok, string reason)
        {
            if (!Want(id))
            {
                return;
            }

            if (ok)
            {
                run.CaseOk(id);
            }
            else
            {
                run.CaseFail(id, reason ?? "");
            }
        }

        private static bool AnyWrite()
        {
            foreach (string id in Ids)
            {
                if (id != "cd-no-quit" && Want(id))
                {
                    return true;
                }
            }

            return false;
        }

        private static void FailWanted(TestRun run, string reason)
        {
            foreach (string id in Ids)
            {
                if (id != "cd-no-quit" && Want(id))
                {
                    run.CaseFail(id, reason);
                }
            }
        }

        private static bool Want(string id)
        {
            return Want(id, IdFilter);
        }

        private static bool Want(string id, string filter)
        {
            if (string.IsNullOrEmpty(filter) || filter == "cd" || filter == "cd-")
            {
                return true;
            }

            return id.StartsWith(filter, StringComparison.OrdinalIgnoreCase);
        }
    }
}
