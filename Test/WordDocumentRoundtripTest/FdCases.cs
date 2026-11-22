using System;
using System.Collections.Generic;
using Newtonsoft.Json.Linq;
using WordAddIn1;

namespace WordDocumentRoundtripTest
{
    /// <summary>格式属性恢复默认。对照 docs/测试/word测试/格式属性恢复默认。前缀 fd-。</summary>
    internal static class FdCases
    {
        public const int Target = 12;

        private static HashSet<string> Only;

        public static void Run(TestRun run, string host, string filter)
        {
            Only = Pick(filter, "fd");
            if (Only != null)
            {
                run.TotalTarget = Only.Count;
            }

            var session = new WdSession(host);
            RunChar(run, session);
            RunPara(run, session);
        }

        private static void RunChar(TestRun run, WdSession session)
        {
            Case(run, "fd-char-bold", () =>
            {
                using (FormatProbe doc = Open(session, "直接加粗。"))
                {
                    string err = ReadSent(doc);
                    if (err != null)
                    {
                        return err;
                    }

                    if (!session.PaintNeedle(doc.Path, "加粗", 0, Font("Bold", 1), null))
                    {
                        return "铺加粗失败";
                    }

                    ToolResult applied = ApplyChar(session, doc, doc.Sentence("加粗"), Map("is_bold", "default"));
                    JObject direct = CharDirect(session, doc, doc.Sentence("加粗"), "加粗");
                    return Ok(applied) && Off(direct, "is_bold") ? null : Fail(applied, direct);
                }
            });

            Case(run, "fd-char-keep-size", () =>
            {
                using (FormatProbe doc = Open(session, "加粗十六。"))
                {
                    string err = ReadSent(doc);
                    if (err != null)
                    {
                        return err;
                    }

                    var font = Font("Bold", 1);
                    font["Size"] = 16f;
                    if (!session.PaintNeedle(doc.Path, "十六", 0, font, null))
                    {
                        return "铺字号失败";
                    }

                    ToolResult applied = ApplyChar(session, doc, doc.Sentence("十六"), Map("is_bold", "default"));
                    JObject direct = CharDirect(session, doc, doc.Sentence("十六"), "十六");
                    return Ok(applied) && Off(direct, "is_bold") && Is16(Txt(direct, "font_size"))
                        ? null
                        : Fail(applied, direct);
                }
            });

            Case(run, "fd-char-mix", () =>
            {
                using (FormatProbe doc = Open(session, "混着写。"))
                {
                    string err = ReadSent(doc);
                    if (err != null)
                    {
                        return err;
                    }

                    if (!session.PaintNeedle(doc.Path, "混着", 0, Font("Bold", 1), null))
                    {
                        return "铺加粗失败";
                    }

                    var format = Map("is_bold", "default");
                    format["font_size"] = 18;
                    ToolResult applied = ApplyChar(session, doc, doc.Sentence("混着"), format);
                    JObject direct = CharDirect(session, doc, doc.Sentence("混着"), "混着");
                    return Ok(applied) && Off(direct, "is_bold") && Starts(Txt(direct, "font_size"), "18")
                        ? null
                        : Fail(applied, direct);
                }
            });

            Case(run, "fd-char-skip", () =>
            {
                using (FormatProbe doc = Open(session, "缺键不动。"))
                {
                    string err = ReadSent(doc);
                    if (err != null)
                    {
                        return err;
                    }

                    if (!session.PaintNeedle(doc.Path, "缺键", 0, Font("Bold", 1), null))
                    {
                        return "铺加粗失败";
                    }

                    ToolResult applied = ApplyChar(session, doc, doc.Sentence("缺键"), Map("font_size", 18));
                    JObject direct = CharDirect(session, doc, doc.Sentence("缺键"), "缺键");
                    return Ok(applied) && Flag(direct, "is_bold") && Starts(Txt(direct, "font_size"), "18")
                        ? null
                        : Fail(applied, direct);
                }
            });

            Case(run, "fd-char-token", () =>
            {
                using (FormatProbe doc = Open(session, "大小写。"))
                {
                    string err = ReadSent(doc);
                    if (err != null)
                    {
                        return err;
                    }

                    if (!session.PaintNeedle(doc.Path, "大小写", 0, Font("Bold", 1), null))
                    {
                        return "铺加粗失败";
                    }

                    ToolResult applied = ApplyChar(session, doc, doc.Sentence("大小写"), Map("is_bold", "Default"));
                    JObject direct = CharDirect(session, doc, doc.Sentence("大小写"), "大小写");
                    return Ok(applied) && Flag(direct, "is_bold") ? null : Fail(applied, direct);
                }
            });

            Case(run, "fd-char-font", () =>
            {
                using (FormatProbe doc = Open(session, "换字体。"))
                {
                    string err = ReadSent(doc);
                    if (err != null)
                    {
                        return err;
                    }

                    var font = Font("Name", "黑体");
                    font["NameFarEast"] = "黑体";
                    if (!session.PaintNeedle(doc.Path, "换字体", 0, font, null))
                    {
                        return "铺字体失败";
                    }

                    ToolResult applied = ApplyChar(session, doc, doc.Sentence("换字体"), Map("font_name", "default"));
                    JObject direct = CharDirect(session, doc, doc.Sentence("换字体"), "换字体");
                    string name = Txt(direct, "font_name");
                    bool warned = WarnKey(applied, "font_name");
                    bool kept = name.IndexOf("黑体", StringComparison.Ordinal) >= 0;
                    bool replaced = name.Length > 0 && !kept;
                    return Ok(applied) && ((warned && kept) || (!warned && replaced))
                        ? null
                        : Fail(applied, direct);
                }
            });
        }

        private static void RunPara(TestRun run, WdSession session)
        {
            Case(run, "fd-para-align", () =>
            {
                using (FormatProbe doc = Open(session, "原先居中。"))
                {
                    string err = ReadPara(doc);
                    if (err != null)
                    {
                        return err;
                    }

                    if (!session.SetParagraphAlign(doc.Path, 1, 1))
                    {
                        return "铺居中失败";
                    }

                    string code = doc.Paragraph("原先居中");
                    ToolResult applied = ApplyPara(session, doc, "target_paragraph_codes", code, Map("alignment", "default"));
                    JObject direct = Obj(ParaLayers(session, doc, code), "direct");
                    return Ok(applied) && Txt(direct, "alignment") == "left" ? null : Fail(applied, direct);
                }
            });

            Case(run, "fd-para-mix", () =>
            {
                using (FormatProbe doc = Open(session, "居中改距。"))
                {
                    string err = ReadPara(doc);
                    if (err != null)
                    {
                        return err;
                    }

                    if (!session.SetParagraphAlign(doc.Path, 1, 1))
                    {
                        return "铺居中失败";
                    }

                    string code = doc.Paragraph("居中改距");
                    var format = Map("alignment", "default");
                    format["space_before"] = 18;
                    ToolResult applied = ApplyPara(session, doc, "target_paragraph_codes", code, format);
                    JObject direct = Obj(ParaLayers(session, doc, code), "direct");
                    return Ok(applied) && Txt(direct, "alignment") == "left" && Near(Txt(direct, "space_before"), 18)
                        ? null
                        : Fail(applied, direct);
                }
            });

            Case(run, "fd-para-skip", () =>
            {
                using (FormatProbe doc = Open(session, "居中留着。"))
                {
                    string err = ReadPara(doc);
                    if (err != null)
                    {
                        return err;
                    }

                    if (!session.SetParagraphAlign(doc.Path, 1, 1))
                    {
                        return "铺居中失败";
                    }

                    string code = doc.Paragraph("居中留着");
                    ToolResult applied = ApplyPara(session, doc, "target_paragraph_codes", code, Map("space_before", 6));
                    JObject direct = Obj(ParaLayers(session, doc, code), "direct");
                    return Ok(applied) && Txt(direct, "alignment") == "center" && Near(Txt(direct, "space_before"), 6)
                        ? null
                        : Fail(applied, direct);
                }
            });

            Case(run, "fd-para-inherit", () =>
            {
                using (FormatProbe doc = Open(session, "继承不改。"))
                {
                    string err = ReadPara(doc);
                    if (err != null)
                    {
                        return err;
                    }

                    if (!session.SetParagraphAlign(doc.Path, 1, 1))
                    {
                        return "铺居中失败";
                    }

                    string code = doc.Paragraph("继承不改");
                    ToolResult applied = ApplyPara(session, doc, "target_paragraph_codes", code, Map("alignment", "inherit"));
                    JObject direct = Obj(ParaLayers(session, doc, code), "direct");
                    return Ok(applied) && Txt(direct, "alignment") == "center" ? null : Fail(applied, direct);
                }
            });

            Case(run, "fd-para-style", () =>
            {
                using (FormatProbe doc = Open(session, "标题留着。"))
                {
                    string err = ReadPara(doc);
                    if (err != null)
                    {
                        return err;
                    }

                    if (!session.ApplyParagraphStyle(doc.Path, "标题留着", new[] { "标题 1", "Heading 1", "heading 1" }))
                    {
                        return "套标题失败";
                    }

                    string code = doc.Paragraph("标题留着");
                    var format = Map("paragraph_style", "default");
                    format["alignment"] = "center";
                    ToolResult applied = ApplyPara(session, doc, "target_paragraph_codes", code, format);
                    JObject direct = Obj(ParaLayers(session, doc, code), "direct");
                    string style = ParaStyle(doc, "标题留着");
                    bool heading = style.IndexOf("标题", StringComparison.Ordinal) >= 0
                        || style.IndexOf("Heading", StringComparison.OrdinalIgnoreCase) >= 0;
                    return Ok(applied) && WarnKey(applied, "paragraph_style") && heading && Txt(direct, "alignment") == "center"
                        ? null
                        : Fail(applied, direct) + " 样式=" + style;
                }
            });

            Case(run, "fd-para-tabs", () =>
            {
                using (FormatProbe doc = Open(session, "制表位。"))
                {
                    string err = ReadPara(doc);
                    if (err != null)
                    {
                        return err;
                    }

                    string code = doc.Paragraph("制表位");
                    var tab = new Dictionary<string, object>
                    {
                        ["position"] = 72,
                        ["alignment"] = "left",
                    };
                    ToolResult planted = ApplyPara(
                        session,
                        doc,
                        "target_paragraph_codes",
                        code,
                        Map("tab_stops", new List<object> { tab }));
                    JObject plantedDirect = Obj(ParaLayers(session, doc, code), "direct");
                    if (!Ok(planted) || Arr(plantedDirect, "tab_stops").Count < 1)
                    {
                        return "铺制表位失败 " + Fail(planted, plantedDirect);
                    }

                    ToolResult applied = ApplyPara(session, doc, "target_paragraph_codes", code, Map("tab_stops", "default"));
                    JObject direct = Obj(ParaLayers(session, doc, code), "direct");
                    return Ok(applied) && Arr(direct, "tab_stops").Count == 0 ? null : Fail(applied, direct);
                }
            });
        }

        private static ToolResult ApplyChar(WdSession session, FormatProbe doc, string code, Dictionary<string, object> format)
        {
            Dictionary<string, object> args = Args(doc.Id, "target_codes", code);
            args["char_format"] = format;
            return session.ApplyDocumentFormat(args);
        }

        private static ToolResult ApplyPara(
            WdSession session,
            FormatProbe doc,
            string key,
            string code,
            Dictionary<string, object> format)
        {
            Dictionary<string, object> args = Args(doc.Id, key, code);
            args["para_format"] = format;
            return session.ApplyParagraphFormat(args);
        }

        private static JObject CharDirect(WdSession session, FormatProbe doc, string code, string text)
        {
            JObject data = FormatProbe.Data(session.GetFormatContext(Args(doc.Id, "target_codes", code)));
            JArray entries = data["entries"] as JArray;
            if (entries == null)
            {
                return null;
            }

            foreach (JObject entry in entries)
            {
                JArray runs = entry["runs"] as JArray;
                if (runs == null)
                {
                    continue;
                }

                foreach (JObject run in runs)
                {
                    if (Txt(run, "text").IndexOf(text, StringComparison.Ordinal) >= 0)
                    {
                        return Obj(Obj(run, "format_layers"), "direct");
                    }
                }
            }

            return null;
        }

        private static JObject ParaLayers(WdSession session, FormatProbe doc, string code)
        {
            JObject data = FormatProbe.Data(session.GetFormatContext(Args(doc.Id, "target_paragraph_codes", code)));
            JArray entries = data["entries"] as JArray;
            if (entries == null || entries.Count == 0)
            {
                return null;
            }

            foreach (JObject entry in entries)
            {
                if (Txt(entry, "paragraph_code") == code)
                {
                    return Obj(entry, "format_layers");
                }
            }

            return Obj(entries[0] as JObject, "format_layers");
        }

        private static string ParaStyle(FormatProbe doc, string needle)
        {
            string name = "";
            doc.EditParagraph(needle, para =>
            {
                try
                {
                    object style = WdSession.ComGet(para, "Style");
                    name = Convert.ToString(WdSession.ComGet(style, "NameLocal")) ?? "";
                }
                catch (Exception)
                {
                    name = "";
                }
            });
            return name;
        }

        private static string ReadSent(FormatProbe doc)
        {
            if (!doc.Ready)
            {
                return doc.Id;
            }

            return doc.ReadSentences() ? null : "没有句子码";
        }

        private static string ReadPara(FormatProbe doc)
        {
            if (!doc.Ready)
            {
                return doc.Id;
            }

            return doc.ReadParagraphs() ? null : "没有段落码";
        }

        private static Dictionary<string, object> Args(string id, string key, string value)
        {
            return new Dictionary<string, object> { ["channel_id"] = id, [key] = value };
        }

        private static Dictionary<string, object> Map(string key, object value)
        {
            return new Dictionary<string, object> { [key] = value };
        }

        private static Dictionary<string, object> Font(string key, object value)
        {
            return new Dictionary<string, object> { [key] = value };
        }

        private static bool Ok(ToolResult result)
        {
            return result != null && result.Success;
        }

        private static bool WarnKey(ToolResult result, string key)
        {
            JArray warnings = FormatProbe.Data(result)["format_warnings"] as JArray;
            if (warnings == null)
            {
                return false;
            }

            foreach (JObject warning in warnings)
            {
                if (Txt(warning, "key") == key)
                {
                    return true;
                }
            }

            return false;
        }

        private static bool Flag(JObject data, string key)
        {
            bool bit;
            return data != null && data[key] != null && bool.TryParse(data[key].ToString(), out bit) && bit;
        }

        private static bool Off(JObject data, string key)
        {
            bool bit;
            return data != null && data[key] != null && bool.TryParse(data[key].ToString(), out bit) && !bit;
        }

        private static bool Starts(string text, string prefix)
        {
            return text != null && text.StartsWith(prefix, StringComparison.Ordinal);
        }

        private static bool Is16(string size)
        {
            return size == "16" || size == "16.0";
        }

        private static bool Near(string text, double expected)
        {
            double value;
            return double.TryParse(text, out value) && Math.Abs(value - expected) < 1.5;
        }

        private static JObject Obj(JObject data, string key)
        {
            return data == null ? null : data[key] as JObject;
        }

        private static JArray Arr(JObject data, string key)
        {
            return data == null ? new JArray() : (data[key] as JArray ?? new JArray());
        }

        private static string Txt(JObject data, string key)
        {
            return data == null || data[key] == null ? "" : data[key].ToString();
        }

        private static string Fail(ToolResult result, JToken token)
        {
            string error = result == null ? "无回包" : (result.Success ? "" : result.Error);
            string body = token == null ? "空" : token.ToString();
            return (error + " " + body).Trim();
        }

        private static FormatProbe Open(WdSession session, string text)
        {
            return FormatProbe.Open(session, text, "");
        }

        private static void Case(TestRun run, string id, Func<string> body)
        {
            if (Only != null && !Only.Contains(id))
            {
                return;
            }

            try
            {
                string error = body();
                if (error == null)
                {
                    run.CaseOk(id);
                }
                else
                {
                    run.CaseFail(id, error);
                }
            }
            catch (Exception ex)
            {
                run.CaseFail(id, ex.Message);
            }
        }

        private static HashSet<string> Pick(string filter, string prefix)
        {
            if (string.IsNullOrWhiteSpace(filter)
                || filter.Equals(prefix, StringComparison.OrdinalIgnoreCase)
                || filter.Equals(prefix + "-", StringComparison.OrdinalIgnoreCase))
            {
                return null;
            }

            var picked = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (string part in filter.Split(','))
            {
                string id = part.Trim();
                if (id.Length > 0)
                {
                    picked.Add(id);
                }
            }

            return picked.Count > 0 ? picked : null;
        }
    }
}
