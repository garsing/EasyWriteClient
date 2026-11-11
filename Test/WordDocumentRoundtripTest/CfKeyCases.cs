using System;
using System.Collections.Generic;
using Newtonsoft.Json.Linq;
using WordAddIn1;

namespace WordDocumentRoundtripTest
{
    internal static partial class CfCases
    {
        private sealed class KeyProbe
        {
            public string Id;
            public string Sentence;
            public Dictionary<string, object> Format;
            public Func<JObject, bool> Pass;
            public bool WpsRejects;
        }

        private static void RunKeys(TestRun run, WdSession s)
        {
            KeyProbe[] probes = KeyProbes();
            var ids = new string[probes.Length];
            var lines = new string[probes.Length];
            for (int i = 0; i < probes.Length; i++)
            {
                ids[i] = probes[i].Id;
                lines[i] = probes[i].Sentence;
            }

            string path, id;
            if (!OpenPlain(s, string.Join("\r", lines), out path, out id))
            {
                FailMany(run, "铺逐键稿失败", ids);
                return;
            }

            try
            {
                s.GetContent(id);
                if (string.IsNullOrEmpty(CodeOf(ReadBody(s, id), probes[0].Sentence)))
                {
                    FailMany(run, "抄不到逐键 S_", ids);
                    return;
                }

                foreach (KeyProbe probe in probes)
                {
                    s.GetContent(id);
                    string code = CodeOf(ReadBody(s, id), probe.Sentence);
                    ToolResult applied = s.ApplyDocumentFormat(Apply(id, code, probe.Format));
                    if (probe.WpsRejects && s.IsWps)
                    {
                        Ok(run, probe.Id,
                            applied != null && !applied.Success && Contains(applied.Error, "不支持"),
                            applied == null ? "无回包" : (applied.Error ?? Dump(applied.Data == null ? null : JToken.FromObject(applied.Data))));
                        continue;
                    }

                    s.GetContent(id);
                    string after = CodeOf(ReadBody(s, id), probe.Sentence);
                    JObject cf = CharOf(s, id, string.IsNullOrEmpty(after) ? code : after);
                    bool pass = applied.Success && probe.Pass(cf);
                    Ok(run, probe.Id, pass, applied.Success ? Dump(cf) : (applied.Error ?? ""));
                }
            }
            finally
            {
                s.CloseDocByPath(path);
            }
        }

        private static KeyProbe[] KeyProbes()
        {
            return new[]
            {
                Bool("cf-key-bold", "易写键加粗。", "is_bold"),
                Bool("cf-key-bold-bi", "易写键复加粗。", "bold_bi"),
                Bool("cf-key-italic-bi", "易写键复斜体。", "italic_bi"),
                Bool("cf-key-double-strike", "易写键双删。", "double_strikethrough"),
                Bool("cf-key-outline", "易写键空心。", "outline"),
                Bool("cf-key-emboss", "易写键阳文。", "emboss"),
                Bool("cf-key-engrave", "易写键阴文。", "engrave"),
                Bool("cf-key-all-caps", "易写键全大写。", "all_caps"),
                Bool("cf-key-grid", "易写键网格。", "disable_character_space_grid"),
                Bool("cf-key-alternates", "易写键替字。", "contextual_alternates"),
                WpsReject(Bool("cf-key-noproof", "易写键不校对。", "no_proofing")),
                Text("cf-key-name", "易写键名称。", Char("name", "黑体"), "font_name", "黑"),
                Text("cf-key-name-ascii", "易写键西文。", Char("name_ascii", "Calibri"), "name_ascii", "Calibri"),
                Text("cf-key-name-fareast", "易写键远东。", Char("name_far_east", "楷体"), "font_name", "楷"),
                Text("cf-key-name-bi", "易写键复杂。", Char("name_bi", "Arial"), "name_bi", "Arial"),
                Text("cf-key-name-other", "易写键其它。", Char("name_other", "Times New Roman"), "name_other", "Times"),
                NumKey("cf-key-size-bi", "易写键复字号。", "size_bi", 18),
                Absent("cf-key-color-index", "易写键色号。", Char("color_index", "wdRed"), "color_index"),
                Absent("cf-key-color-index-bi", "易写键复色号。", Char("color_index_bi", "wdBlue"), "color_index_bi"),
                Text("cf-key-text-color", "易写键主题色。",
                    new Dictionary<string, object>
                    {
                        ["text_color"] = new Dictionary<string, object> { ["rgb"] = "#112233" }
                    },
                    "font_color", "#112233"),
                Text("cf-key-underline-color", "易写键线色。", Char("underline_color", "#FF00FF"), "underline_color", "#FF00FF"),
                Absent("cf-key-diacritic-color", "易写键音符色。", Char("diacritic_color", "#00FFFF"), "diacritic_color"),
                Text("cf-key-emphasis", "易写键着重。", Char("emphasis_mark", "wdEmphasisMarkOverSolidCircle"), "emphasis_mark", "SolidCircle"),
                NumKey("cf-key-scaling", "易写键缩放。", "scaling", 150),
                NumKey("cf-key-position", "易写键位置。", "position", 3),
                NumKey("cf-key-kerning", "易写键字距。", "kerning", 12),
                Text("cf-key-ligatures", "易写键连字。", Char("ligatures", "wdLigaturesHistorical"), "ligatures", "Historical"),
                Text("cf-key-number-form", "易写键数字形。", Char("number_form", "wdNumberFormLining"), "number_form", "Lining"),
                Text("cf-key-number-spacing", "易写键数字距。", Char("number_spacing", "wdNumberSpacingProportional"), "number_spacing", "Proportional"),
                Text("cf-key-stylistic", "易写键样式集。", Char("stylistic_set", "wdStylisticSet01"), "stylistic_set", "01"),
                Text("cf-key-background", "易写键底色。", Char("background_color", "#00FF00"), "background_color", "#00FF00"),
                Nested("cf-key-borders", "易写键边框。", "borders",
                    new Dictionary<string, object> { ["outside_line_style"] = "wdLineStyleSingle" },
                    "outside_line_style", "Single"),
                Nested("cf-key-fill", "易写键填充。", "fill",
                    new Dictionary<string, object> { ["visible"] = true, ["transparency"] = 0.4 },
                    "transparency", "0.4"),
                Nested("cf-key-line", "易写键线条。", "line",
                    new Dictionary<string, object> { ["visible"] = true, ["weight"] = 2 },
                    "weight", "2"),
                Nested("cf-key-reflection", "易写键倒影。", "reflection",
                    new Dictionary<string, object> { ["type"] = "msoReflectionType1" },
                    "type", "Type1"),
                Nested("cf-key-text-shadow", "易写键文字影。", "text_shadow",
                    new Dictionary<string, object> { ["visible"] = true },
                    "visible", "true"),
                Text("cf-key-lang", "KeyLangEn.", Char("language_id", "wdEnglishUS"), "language_id", "English"),
                Text("cf-key-lang-east", "易写键东亚语。", Char("language_id_far_east", "wdJapanese"), "language_id_far_east", "Japanese"),
                NumKey("cf-key-fit", "易写键撑宽。", "fit_text_width", 72)
            };
        }

        private static KeyProbe WpsReject(KeyProbe probe)
        {
            probe.WpsRejects = true;
            return probe;
        }

        private static KeyProbe Bool(string id, string sentence, string key)
        {
            return new KeyProbe
            {
                Id = id,
                Sentence = sentence,
                Format = Char(key, true),
                Pass = cf => Flag(cf, key)
            };
        }

        private static KeyProbe NumKey(string id, string sentence, string key, int expect)
        {
            return new KeyProbe
            {
                Id = id,
                Sentence = sentence,
                Format = Char(key, expect),
                Pass = cf => Near(Num(cf, key), expect, 0.6)
            };
        }

        private static KeyProbe Absent(string id, string sentence, Dictionary<string, object> format, string key)
        {
            return new KeyProbe
            {
                Id = id,
                Sentence = sentence,
                Format = format,
                Pass = cf => Has(cf, "font_name") && !Has(cf, key)
            };
        }

        private static KeyProbe Text(string id, string sentence, Dictionary<string, object> format, string key, string needle)
        {
            return new KeyProbe
            {
                Id = id,
                Sentence = sentence,
                Format = format,
                Pass = cf => Contains(Str(cf, key), needle)
            };
        }

        private static KeyProbe Nested(
            string id,
            string sentence,
            string key,
            Dictionary<string, object> inner,
            string field,
            string needle)
        {
            return new KeyProbe
            {
                Id = id,
                Sentence = sentence,
                Format = new Dictionary<string, object> { [key] = inner },
                Pass = cf => Contains(Str(Obj(cf, key), field), needle)
            };
        }
    }
}
