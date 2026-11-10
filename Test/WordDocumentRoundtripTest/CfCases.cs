using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using WordAddIn1;

namespace WordDocumentRoundtripTest
{
    /// <summary>
    /// 字符格式全量：读默认省略、COM 铺后再读、套用、写入。对照 docs/测试/word测试/字符格式读写。
    /// </summary>
    internal static partial class CfCases
    {
        public const int Target = 67;

        private const string Plain = "易写字符甲句。";
        private const string Italic = "易写字符乙斜。";
        private const string Super = "易写字符丙上。";
        private const string Sub = "易写字符丁下。";
        private const string Strike = "易写字符戊删。";
        private const string Red = "易写字符己红。";
        private const string Shadow = "易写字符庚影。";
        private const string Double = "易写字符辛双。";
        private const string Highlight = "易写字符壬亮。";
        private const string Mixed = "易写字符癸混。";

        private const string ApplyA = "易写套用甲句。";
        private const string ApplyB = "易写套用乙句。";
        private const string ApplyC = "易写套用丙句。";
        private const string ApplyD = "易写套用丁句。";
        private const string ApplyE = "易写套用戊句。";
        private const string ApplyF = "易写套用己句。";
        private const string ApplyG = "易写套用庚句。";
        private const string ApplyH = "易写套用辛句。";

        private const string Anchor = "易写写入锚句。";
        private const string WroteItalic = "易写写入斜句。";
        private const string WroteSuper = "易写写入上句。";

        private const int On = -1;
        private const int Off = 0;
        private const int UnderlineDouble = 3;
        private const int ColorRed = 255;
        private const int HighlightYellow = 7;

        private static readonly string[] TenKeys =
        {
            "font_name", "font_size", "font_color", "background_color",
            "is_bold", "is_italic", "has_underline", "has_strikethrough",
            "is_superscript", "is_subscript"
        };

        public static void Run(TestRun run, string host)
        {
            var s = new WdSession(host);
            RunRead(run, s);
            RunApply(run, s);
            RunWrite(run, s);
            RunKeys(run, s);
        }

        private static void RunRead(TestRun run, WdSession s)
        {
            string[] ids =
            {
                "cf-plain-ten", "cf-plain-omit", "cf-plain-flags", "cf-plain-clean",
                "cf-com-italic", "cf-com-super", "cf-com-sub", "cf-com-strike",
                "cf-com-color", "cf-com-shadow", "cf-com-double", "cf-com-highlight",
                "cf-com-mixed"
            };
            string path, id;
            if (!OpenPlain(s, string.Join("\r", new[]
            {
                Plain, Italic, Super, Sub, Strike, Red, Shadow, Double, Highlight, Mixed
            }), out path, out id))
            {
                FailMany(run, "铺读稿失败", ids);
                return;
            }

            try
            {
                bool painted = s.PaintNeedle(path, Italic, 0, Font("Italic", On), null)
                    && s.PaintNeedle(path, Super, 0, Font("Superscript", On), null)
                    && s.PaintNeedle(path, Sub, 0, Font("Subscript", On), null)
                    && s.PaintNeedle(path, Strike, 0, Font("StrikeThrough", On), null)
                    && s.PaintNeedle(path, Red, 0, Font("Color", ColorRed), null)
                    && s.PaintNeedle(path, Shadow, 0, Font("Shadow", On), null)
                    && s.PaintNeedle(path, Double, 0, Font("Underline", UnderlineDouble), null)
                    && s.PaintNeedle(path, Highlight, 0, null, RangeProp("HighlightColorIndex", HighlightYellow))
                    && s.PaintNeedle(path, Mixed, 2, Font("Superscript", On), null);
                if (!painted)
                {
                    FailMany(run, "COM 铺字符格式失败", ids);
                    return;
                }

                s.GetContent(id);
                string body = ReadBody(s, id);
                string plainCode = CodeOf(body, Plain);
                string italicCode = CodeOf(body, Italic);
                string superCode = CodeOf(body, Super);
                string subCode = CodeOf(body, Sub);
                string strikeCode = CodeOf(body, Strike);
                string redCode = CodeOf(body, Red);
                string shadowCode = CodeOf(body, Shadow);
                string doubleCode = CodeOf(body, Double);
                string highlightCode = CodeOf(body, Highlight);
                string mixedCode = CodeOf(body, Mixed);
                if (string.IsNullOrEmpty(plainCode) || string.IsNullOrEmpty(mixedCode))
                {
                    FailMany(run, "抄不到 S_", ids);
                    return;
                }

                JObject plainEntry = ReadEntry(s, id, plainCode);
                JObject plainCf = Obj(plainEntry, "char_format");
                Ok(run, "cf-plain-ten", HasTen(plainCf), Dump(plainCf));
                Ok(run, "cf-plain-omit",
                    !Has(plainCf, "shadow") && !Has(plainCf, "kerning") && !Has(plainCf, "underline")
                    && !Has(plainCf, "highlight_color_index") && !Has(plainCf, "glow")
                    && !Has(plainCf, "color_index") && !Has(plainCf, "ligatures")
                    && !Has(plainCf, "fill"),
                    Dump(plainCf));
                Ok(run, "cf-plain-flags",
                    Has(plainCf, "is_italic") && !Flag(plainCf, "is_italic")
                    && Has(plainCf, "is_superscript") && !Flag(plainCf, "is_superscript")
                    && Has(plainCf, "is_subscript") && !Flag(plainCf, "is_subscript"),
                    Dump(plainCf));
                Ok(run, "cf-plain-clean",
                    plainEntry != null && plainEntry["format_note"] == null
                    && !Blob(plainEntry, "9999999"),
                    Dump(plainEntry));

                JObject italicCf = CharOf(s, id, italicCode);
                Ok(run, "cf-com-italic", Flag(italicCf, "is_italic"), Dump(italicCf));
                JObject superCf = CharOf(s, id, superCode);
                Ok(run, "cf-com-super", Flag(superCf, "is_superscript") && !Flag(superCf, "is_subscript"), Dump(superCf));
                JObject subCf = CharOf(s, id, subCode);
                Ok(run, "cf-com-sub", Flag(subCf, "is_subscript") && !Flag(subCf, "is_superscript"), Dump(subCf));
                Ok(run, "cf-com-strike", Flag(CharOf(s, id, strikeCode), "has_strikethrough"), Dump(CharOf(s, id, strikeCode)));
                string color = Str(CharOf(s, id, redCode), "font_color");
                Ok(run, "cf-com-color", string.Equals(color, "#FF0000", StringComparison.OrdinalIgnoreCase), color ?? "");
                Ok(run, "cf-com-shadow", Flag(CharOf(s, id, shadowCode), "shadow"), Dump(CharOf(s, id, shadowCode)));
                JObject doubleCf = CharOf(s, id, doubleCode);
                Ok(run, "cf-com-double",
                    Flag(doubleCf, "has_underline") && Contains(Str(doubleCf, "underline"), "Double"),
                    Dump(doubleCf));
                JObject hiCf = CharOf(s, id, highlightCode);
                Ok(run, "cf-com-highlight",
                    Contains(Str(hiCf, "highlight_color_index"), "Yellow") && Has(hiCf, "background_color"),
                    Dump(hiCf));

                JObject mixedEntry = ReadEntry(s, id, mixedCode);
                JObject mixedCf = Obj(mixedEntry, "char_format");
                int kinds = Num(mixedEntry, "format_kind_count");
                Ok(run, "cf-com-mixed",
                    kinds == 2
                    && Contains(Str(mixedEntry, "format_note"), "两种")
                    && !Has(mixedCf, "is_superscript")
                    && !Blob(mixedEntry, "9999999"),
                    Dump(mixedEntry));
            }
            finally
            {
                s.CloseDocByPath(path);
            }
        }

        private static void RunApply(TestRun run, WdSession s)
        {
            string[] ids =
            {
                "cf-apply-italic", "cf-apply-super", "cf-apply-sub", "cf-apply-name",
                "cf-apply-partial", "cf-apply-spacing", "cf-apply-caps", "cf-apply-warn",
                "cf-apply-text", "cf-apply-glow", "cf-apply-shading"
            };
            string path, id;
            if (!OpenPlain(s, string.Join("\r", new[]
            {
                ApplyA, ApplyB, ApplyC, ApplyD, ApplyE, ApplyF, ApplyG, ApplyH
            }), out path, out id))
            {
                FailMany(run, "铺套用稿失败", ids);
                return;
            }

            try
            {
                string Fresh(string sentence)
                {
                    s.GetContent(id);
                    return CodeOf(ReadBody(s, id), sentence);
                }

                if (string.IsNullOrEmpty(Fresh(ApplyA)) || string.IsNullOrEmpty(Fresh(ApplyH)))
                {
                    FailMany(run, "抄不到套用 S_", ids);
                    return;
                }

                ToolResult italic = s.ApplyDocumentFormat(Apply(id, Fresh(ApplyA), Char("is_italic", true)));
                JObject italicCf = CharOf(s, id, Fresh(ApplyA));
                Ok(run, "cf-apply-italic", italic.Success && Flag(italicCf, "is_italic"), Why(italic, italicCf));

                ToolResult super = s.ApplyDocumentFormat(Apply(id, Fresh(ApplyB), Char("is_superscript", true)));
                JObject superCf = CharOf(s, id, Fresh(ApplyB));
                Ok(run, "cf-apply-super", super.Success && Flag(superCf, "is_superscript") && !Flag(superCf, "is_subscript"), Why(super, superCf));

                ToolResult sub = s.ApplyDocumentFormat(Apply(id, Fresh(ApplyC), Char("is_subscript", true)));
                JObject subCf = CharOf(s, id, Fresh(ApplyC));
                Ok(run, "cf-apply-sub", sub.Success && Flag(subCf, "is_subscript") && !Flag(subCf, "is_superscript"), Why(sub, subCf));

                ToolResult name = s.ApplyDocumentFormat(Apply(id, Fresh(ApplyD), new Dictionary<string, object>
                {
                    ["font_name"] = "楷体",
                    ["font_size"] = 16
                }));
                JObject nameCf = CharOf(s, id, Fresh(ApplyD));
                Ok(run, "cf-apply-name",
                    name.Success && FontLooksLike(Str(nameCf, "font_name"), "楷") && Near(Num(nameCf, "font_size"), 16, 0.6),
                    Why(name, nameCf));

                ToolResult partial = s.ApplyDocumentFormat(Apply(id, Fresh(ApplyE), Char("is_italic", true)));
                JObject partialCf = CharOf(s, id, Fresh(ApplyE));
                Ok(run, "cf-apply-partial",
                    partial.Success && Flag(partialCf, "is_italic") && Has(partialCf, "is_bold") && !Flag(partialCf, "is_bold"),
                    Why(partial, partialCf));

                ToolResult spacing = s.ApplyDocumentFormat(Apply(id, Fresh(ApplyF), Char("spacing", 2)));
                JObject spacingCf = CharOf(s, id, Fresh(ApplyF));
                Ok(run, "cf-apply-spacing", spacing.Success && Near(Num(spacingCf, "spacing"), 2, 0.3), Why(spacing, spacingCf));

                ToolResult caps = s.ApplyDocumentFormat(Apply(id, Fresh(ApplyG), Char("small_caps", true)));
                JObject capsCf = CharOf(s, id, Fresh(ApplyG));
                Ok(run, "cf-apply-caps", caps.Success && Flag(capsCf, "small_caps"), Why(caps, capsCf));

                var warnFmt = Char("is_italic", true);
                warnFmt["underline"] = "not-an-underline";
                ToolResult warn = s.ApplyDocumentFormat(Apply(id, Fresh(ApplyH), warnFmt));
                JObject warnCf = CharOf(s, id, Fresh(ApplyH));
                Ok(run, "cf-apply-warn",
                    warn.Success && Flag(warnCf, "is_italic") && Blob(DataObj(warn), "format_warnings"),
                    Why(warn, DataObj(warn)));

                Ok(run, "cf-apply-text",
                    StillHas(s, id, ApplyA) && StillHas(s, id, ApplyH),
                    "套用后句子丢了");

                ToolResult glow = s.ApplyDocumentFormat(Apply(id, Fresh(ApplyA), new Dictionary<string, object>
                {
                    ["glow"] = new Dictionary<string, object> { ["radius"] = 8 }
                }));
                JObject glowCf = CharOf(s, id, Fresh(ApplyA));
                JObject glowObj = Obj(glowCf, "glow");
                Ok(run, "cf-apply-glow", glow.Success && Near(Num(glowObj, "radius"), 8, 0.6), Why(glow, glowCf));

                ToolResult shading = s.ApplyDocumentFormat(Apply(id, Fresh(ApplyB), new Dictionary<string, object>
                {
                    ["shading"] = new Dictionary<string, object>
                    {
                        ["background_pattern_color"] = "#00FF00"
                    }
                }));
                JObject shadingObj = Obj(CharOf(s, id, Fresh(ApplyB)), "shading");
                string shade = Str(shadingObj, "background_pattern_color");
                Ok(run, "cf-apply-shading",
                    shading.Success && string.Equals(shade, "#00FF00", StringComparison.OrdinalIgnoreCase),
                    shading.Success ? (shade ?? Dump(shadingObj)) : (shading.Error ?? ""));
            }
            finally
            {
                s.CloseDocByPath(path);
            }
        }

        private static void RunWrite(TestRun run, WdSession s)
        {
            string[] ids = { "cf-write-italic", "cf-write-bg", "cf-write-super", "cf-write-text" };
            string path, id;
            if (!OpenPlain(s, Anchor, out path, out id))
            {
                FailMany(run, "铺写入稿失败", ids);
                return;
            }

            try
            {
                string Fresh(string sentence)
                {
                    s.GetContent(id);
                    return CodeOf(ReadBody(s, id), sentence);
                }

                if (string.IsNullOrEmpty(Fresh(Anchor)))
                {
                    FailMany(run, "抄不到锚句", ids);
                    return;
                }

                ToolResult italic = s.ProcessActions(id, InsertAfter(Fresh(Anchor), WroteItalic, new Dictionary<string, object>
                {
                    ["font_name"] = "宋体",
                    ["font_size"] = 12,
                    ["is_italic"] = true
                }));
                s.GetContent(id);
                string italicCode = CodeOf(ReadBody(s, id), WroteItalic);
                JObject italicCf = CharOf(s, id, italicCode);
                Ok(run, "cf-write-italic", italic.Success && Flag(italicCf, "is_italic"), italic.Error ?? Dump(italicCf));
                Ok(run, "cf-write-bg",
                    Has(italicCf, "background_color") && Contains(Str(italicCf, "highlight_color_index"), "Yellow"),
                    Dump(italicCf));

                ToolResult super = s.ProcessActions(id, InsertAfter(Fresh(Anchor), WroteSuper, new Dictionary<string, object>
                {
                    ["font_name"] = "宋体",
                    ["font_size"] = 12,
                    ["is_superscript"] = true,
                    ["background_color"] = "automatic"
                }));
                s.GetContent(id);
                JObject superCf = CharOf(s, id, CodeOf(ReadBody(s, id), WroteSuper));
                Ok(run, "cf-write-super", super.Success && Flag(superCf, "is_superscript"), super.Error ?? Dump(superCf));
                Ok(run, "cf-write-text", StillHas(s, id, WroteItalic) && StillHas(s, id, WroteSuper) && StillHas(s, id, Anchor), "写入后句子丢了");
            }
            finally
            {
                s.CloseDocByPath(path);
            }
        }

        private static bool OpenPlain(WdSession s, string plain, out string path, out string id)
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

        private static Dictionary<string, object> Font(string key, object value)
        {
            return new Dictionary<string, object> { [key] = value };
        }

        private static Dictionary<string, object> RangeProp(string key, object value)
        {
            return new Dictionary<string, object> { [key] = value };
        }

        private static Dictionary<string, object> Char(string key, object value)
        {
            return new Dictionary<string, object> { [key] = value };
        }

        private static Dictionary<string, object> Apply(string id, string code, Dictionary<string, object> format)
        {
            return new Dictionary<string, object>
            {
                ["channel_id"] = id,
                ["target_codes"] = code,
                ["char_format"] = format
            };
        }

        private static List<Dictionary<string, object>> InsertAfter(string code, string text, Dictionary<string, object> charFormat)
        {
            return new List<Dictionary<string, object>>
            {
                new Dictionary<string, object>
                {
                    ["type"] = "insert",
                    ["detail"] = new Dictionary<string, object>
                    {
                        ["target"] = new Dictionary<string, object> { ["codes"] = code },
                        ["insert"] = text,
                        ["format"] = new Dictionary<string, object> { ["char_format"] = charFormat }
                    }
                }
            };
        }

        private static JObject CharOf(WdSession s, string id, string code)
        {
            return Obj(ReadEntry(s, id, code), "char_format");
        }

        private static JObject ReadEntry(WdSession s, string id, string code)
        {
            if (string.IsNullOrEmpty(code))
            {
                return null;
            }

            return FirstEntry(s.GetFormatContext(new Dictionary<string, object>
            {
                ["channel_id"] = id,
                ["target_codes"] = code
            }));
        }

        private static string ReadBody(WdSession s, string id)
        {
            ToolResult got = s.GetContent(id);
            return got.Success ? WdSession.Str(got, "document_content") : "";
        }

        private static bool StillHas(WdSession s, string id, string needle)
        {
            return ReadBody(s, id).IndexOf(needle, StringComparison.Ordinal) >= 0;
        }

        private static string CodeOf(string body, string sentence)
        {
            if (string.IsNullOrEmpty(body) || string.IsNullOrEmpty(sentence))
            {
                return "";
            }

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

        private static bool HasTen(JObject cf)
        {
            if (cf == null)
            {
                return false;
            }

            foreach (string key in TenKeys)
            {
                if (!Has(cf, key))
                {
                    return false;
                }
            }

            return true;
        }

        private static bool Has(JObject o, string key)
        {
            return o != null && o[key] != null && o[key].Type != JTokenType.Null;
        }

        private static bool Flag(JObject o, string key)
        {
            if (!Has(o, key))
            {
                return false;
            }

            JToken token = o[key];
            if (token.Type == JTokenType.Boolean)
            {
                return token.Value<bool>();
            }

            bool parsed;
            return bool.TryParse(token.ToString(), out parsed) && parsed;
        }

        private static string Str(JObject o, string key)
        {
            return Has(o, key) ? o[key].ToString() : "";
        }

        private static int Num(JObject o, string key)
        {
            if (!Has(o, key))
            {
                return int.MinValue;
            }

            try
            {
                return (int)Math.Round(o[key].Value<double>());
            }
            catch
            {
                int n;
                return int.TryParse(o[key].ToString(), out n) ? n : int.MinValue;
            }
        }

        private static bool Near(int actual, int expect, double tol)
        {
            return actual != int.MinValue && Math.Abs(actual - expect) <= tol;
        }

        private static bool FontLooksLike(string actual, string expect)
        {
            return !string.IsNullOrEmpty(actual) && actual.IndexOf(expect, StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private static bool Contains(string text, string needle)
        {
            return !string.IsNullOrEmpty(text) && text.IndexOf(needle, StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private static bool Blob(JToken token, string needle)
        {
            if (token == null)
            {
                return false;
            }

            return token.ToString(Formatting.None).IndexOf(needle, StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private static JObject DataObj(ToolResult r)
        {
            if (r == null || r.Data == null)
            {
                return null;
            }

            try
            {
                return JObject.Parse(JsonConvert.SerializeObject(r.Data));
            }
            catch
            {
                return null;
            }
        }

        private static JObject FirstEntry(ToolResult r)
        {
            JObject data = DataObj(r);
            JArray entries = data == null ? null : data["entries"] as JArray;
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

        private static string Dump(JToken token)
        {
            return token == null ? "空" : token.ToString(Formatting.None);
        }

        private static string Why(ToolResult result, JToken dump)
        {
            if (result != null && !result.Success)
            {
                return result.Error ?? "";
            }

            return Dump(dump);
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
