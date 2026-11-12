using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text.RegularExpressions;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using WordAddIn1;
using Word = Microsoft.Office.Interop.Word;

namespace WordDocumentRoundtripTest
{
    /// <summary>
    /// 段落格式全量。对照 docs/测试/word测试/段落格式读写。铺稿用 COM，读和套走产品工具。
    /// </summary>
    internal static class PfCases
    {
        public const int Target = 83;

        private const int AlignCenter = 1;
        private const int Heading1 = -2;
        private const int GreenBgr = 65280;

        public static void Run(TestRun run, string host)
        {
            var s = new WdSession(host);
            RunRead(run, s);
            RunApply(run, s);
            RunInherit(run, s);
            RunLists(run, s);
            RunDrop(run, s);
            RunDropTable(run, s);
        }

        private static void RunRead(TestRun run, WdSession s)
        {
            string[] ids =
            {
                "pf-plain-omit", "pf-plain-clean", "pf-plain-text",
                "pf-com-center", "pf-com-indent", "pf-com-space", "pf-com-line",
                "pf-com-widow", "pf-com-keep", "pf-com-style", "pf-com-shade",
                "pf-com-border", "pf-com-tab", "pf-com-outline"
            };
            string[] lines =
            {
                "段落甲。", "段落乙。", "段落丙。", "段落丁。", "段落戊。", "段落己。",
                "段落庚。", "段落辛。", "段落壬。", "段落癸。", "段落子。", "段落丑。"
            };
            string path, id;
            if (!OpenPlain(s, string.Join("\r", lines), out path, out id))
            {
                FailMany(run, "铺读稿失败", ids);
                return;
            }

            try
            {
                bool centerOk = Paint(s, path, "段落乙。",
                    p => p.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter,
                    p => SetPf(p, "Alignment", AlignCenter));
                bool indentOk = Paint(s, path, "段落丙。",
                    p => p.Range.ParagraphFormat.FirstLineIndent = 24f,
                    p => SetPf(p, "FirstLineIndent", 24f));
                bool spaceOk = Paint(s, path, "段落丁。",
                    p => p.Range.ParagraphFormat.SpaceBefore = 12f,
                    p => SetPf(p, "SpaceBefore", 12f));
                bool lineOk = Paint(s, path, "段落戊。",
                    p => p.Range.ParagraphFormat.LineSpacingRule = Word.WdLineSpacing.wdLineSpace1pt5,
                    p => SetPf(p, "LineSpacingRule", 1));
                bool widowOk = Paint(s, path, "段落己。",
                    p => p.Range.ParagraphFormat.WidowControl = 0,
                    p => SetPf(p, "WidowControl", 0));
                bool keepOk = Paint(s, path, "段落庚。",
                    p => p.Range.ParagraphFormat.KeepWithNext = -1,
                    p => SetPf(p, "KeepWithNext", -1));
                bool styleOk = Paint(s, path, "段落辛。", SetHeading, SetHeadingLate);
                bool shadeOk = Paint(s, path, "段落壬。",
                    p => p.Range.ParagraphFormat.Shading.BackgroundPatternColor = (Word.WdColor)GreenBgr,
                    p =>
                    {
                        object shading = WdSession.ComGet(Pf(p), "Shading");
                        WdSession.ComSet(shading, "BackgroundPatternColor", GreenBgr);
                    });
                bool borderOk = Paint(s, path, "段落癸。",
                    p => p.Range.ParagraphFormat.Borders[Word.WdBorderType.wdBorderTop].LineStyle = Word.WdLineStyle.wdLineStyleSingle,
                    p =>
                    {
                        object borders = WdSession.ComGet(Pf(p), "Borders");
                        object top = WdSession.ComCall(borders, "Item", (int)Word.WdBorderType.wdBorderTop);
                        WdSession.ComSet(top, "LineStyle", (int)Word.WdLineStyle.wdLineStyleSingle);
                    });
                bool tabOk = Paint(s, path, "段落子。", AddLeftTab, AddLeftTabLate);
                bool outlineOk = Paint(s, path, "段落丑。",
                    p => p.Range.ParagraphFormat.OutlineLevel = Word.WdOutlineLevel.wdOutlineLevel1,
                    p => SetPf(p, "OutlineLevel", 1));

                string body = ReadBody(s, id);
                string jia = ParaCode(body, "段落甲。");
                if (string.IsNullOrEmpty(jia) || string.IsNullOrEmpty(ParaCode(body, "段落丑。")))
                {
                    FailMany(run, "抄不到 P_", ids);
                    return;
                }

                JObject plain = Para(s, id, jia);
                Ok(run, "pf-plain-omit", MissingDefaults(plain), Dump(plain));
                Ok(run, "pf-plain-clean",
                    !Blob(plain, "9999999")
                    && !Has(plain, "shading") && !Has(plain, "borders")
                    && !Has(plain, "list") && !Has(plain, "drop_cap") && !Has(plain, "tab_stops"),
                    Dump(plain));
                Ok(run, "pf-plain-text", body.IndexOf("段落甲。", StringComparison.Ordinal) >= 0, "正文丢了");

                Ok(run, "pf-com-center", centerOk && Eq(Para(s, id, ParaCode(body, "段落乙。")), "alignment", "center"), centerOk ? Dump(Para(s, id, ParaCode(body, "段落乙。"))) : "COM 居中失败");
                Ok(run, "pf-com-indent", indentOk && Near(Para(s, id, ParaCode(body, "段落丙。")), "first_line_indent", 24, 1.2), indentOk ? Dump(Para(s, id, ParaCode(body, "段落丙。"))) : "COM 首行缩进失败");
                Ok(run, "pf-com-space", spaceOk && Near(Para(s, id, ParaCode(body, "段落丁。")), "space_before", 12, 1.2), spaceOk ? Dump(Para(s, id, ParaCode(body, "段落丁。"))) : "COM 段前失败");
                JObject line = Para(s, id, ParaCode(body, "段落戊。"));
                Ok(run, "pf-com-line", lineOk && Eq(line, "line_spacing_rule", "1.5") && Near(line, "line_spacing", 1.5, 0.15), lineOk ? Dump(line) : "COM 行距失败");
                Ok(run, "pf-com-widow", widowOk && Has(Para(s, id, ParaCode(body, "段落己。")), "widow_control") && !Flag(Para(s, id, ParaCode(body, "段落己。")), "widow_control"), widowOk ? Dump(Para(s, id, ParaCode(body, "段落己。"))) : "COM 孤行控制失败");
                Ok(run, "pf-com-keep", keepOk && Flag(Para(s, id, ParaCode(body, "段落庚。")), "keep_with_next"), keepOk ? Dump(Para(s, id, ParaCode(body, "段落庚。"))) : "COM 与下段同页失败");
                string style = Str(Para(s, id, ParaCode(body, "段落辛。")), "paragraph_style");
                Ok(run, "pf-com-style", styleOk && (Contains(style, "标题") || Contains(style, "Heading")), styleOk ? style : "COM 标题样式失败");
                Ok(run, "pf-com-shade", shadeOk && Contains(Deep(Para(s, id, ParaCode(body, "段落壬。")), "shading", "background_pattern_color"), "00FF00"), shadeOk ? Dump(Para(s, id, ParaCode(body, "段落壬。"))) : "COM 底纹失败");
                Ok(run, "pf-com-border", borderOk && Contains(Deep(Para(s, id, ParaCode(body, "段落癸。")), "borders", "top", "line_style"), "Single"), borderOk ? Dump(Para(s, id, ParaCode(body, "段落癸。"))) : "COM 边框失败");
                Ok(run, "pf-com-tab", tabOk && TabNear(Para(s, id, ParaCode(body, "段落子。")), 72), tabOk ? Dump(Para(s, id, ParaCode(body, "段落子。"))) : "COM 制表位失败");
                string outline = Str(Para(s, id, ParaCode(body, "段落丑。")), "outline_level");
                Ok(run, "pf-com-outline", outlineOk && Contains(outline, "1") && !Contains(outline, "BodyText"), outlineOk ? outline : "COM 大纲级别失败");
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
                "pf-apply-center", "pf-apply-15", "pf-apply-partial", "pf-apply-style", "pf-apply-style-indent",
                "pf-apply-hanging", "pf-apply-warn", "pf-apply-chars", "pf-apply-keep", "pf-apply-page",
                "pf-apply-shade", "pf-apply-border", "pf-apply-tabs", "pf-apply-tabs-clear", "pf-apply-east",
                "pf-apply-med", "pf-items-two", "pf-unit", "pf-auto", "pf-text",
                "pf-apply-justify", "pf-apply-distribute", "pf-apply-hi", "pf-apply-low", "pf-apply-thai",
                "pf-apply-right", "pf-apply-after", "pf-apply-cu-left", "pf-apply-cu-right",
                "pf-apply-lu-before", "pf-apply-lu-after", "pf-auto-after",
                "pf-apply-noline", "pf-apply-mirror", "pf-apply-grid", "pf-apply-half", "pf-apply-wrap", "pf-apply-collapse",
                "pf-apply-hyphen", "pf-apply-adjright", "pf-apply-fe-break", "pf-apply-hangpunct", "pf-apply-digit",
                "pf-apply-baseline", "pf-apply-rtl", "pf-apply-tight",
                "pf-apply-texture", "pf-apply-fore", "pf-apply-border-sides", "pf-apply-border-ink", "pf-apply-tab-look"
            };
            var lines = new List<string>();
            foreach (string caseId in ids)
            {
                if (caseId == "pf-text" || caseId == "pf-items-two")
                {
                    continue;
                }

                lines.Add(caseId + "。");
            }

            lines.Add("pf-items-午。");
            lines.Add("pf-items-未。");
            string path, id;
            if (!OpenPlain(s, string.Join("\r", lines), out path, out id))
            {
                FailMany(run, "铺套用稿失败", ids);
                return;
            }

            try
            {
                string body = ReadBody(s, id);
                var code = new Dictionary<string, string>();
                foreach (string line in lines)
                {
                    code[line] = ParaCode(body, line);
                }

                if (string.IsNullOrEmpty(code["pf-apply-center。"]) || string.IsNullOrEmpty(code["pf-items-未。"]))
                {
                    FailMany(run, "抄不到套用 P_", ids);
                    return;
                }

                string heading = s.BuiltinStyleName(path, Heading1);
                CheckApply(run, s, id, code, "pf-apply-center", Map("alignment", "center"), p => Eq(p, "alignment", "center"));
                CheckApply(run, s, id, code, "pf-apply-15",
                    Map("line_spacing_rule", "1.5", "line_spacing", 1.5),
                    p => Eq(p, "line_spacing_rule", "1.5") && Near(p, "line_spacing", 1.5, 0.15));
                CheckApply(run, s, id, code, "pf-apply-partial", Map("space_before", 6),
                    p => Near(p, "space_before", 6, 1.2) && !Has(p, "alignment"));
                CheckApply(run, s, id, code, "pf-apply-style", Map("paragraph_style", heading),
                    p => Contains(Str(p, "paragraph_style"), "标题") || Contains(Str(p, "paragraph_style"), "Heading"));
                CheckApply(run, s, id, code, "pf-apply-style-indent",
                    Map("paragraph_style", heading, "left_indent", 18),
                    p => (Contains(Str(p, "paragraph_style"), "标题") || Contains(Str(p, "paragraph_style"), "Heading"))
                        && Near(p, "left_indent", 18, 1.2));
                ToolResult hanging = Apply(s, id, code["pf-apply-hanging。"], Map("first_line_indent", 24, "hanging_indent", 12));
                JObject hangingFmt = Para(s, id, code["pf-apply-hanging。"]);
                Ok(run, "pf-apply-hanging",
                    hanging.Success && Near(hangingFmt, "hanging_indent", 12, 1.2) && !Has(hangingFmt, "first_line_indent") && WarnHas(hanging, "first_line_indent"),
                    Why(hanging, hangingFmt));
                ToolResult warn = Apply(s, id, code["pf-apply-warn。"], Map("alignment", "center", "textbox_tight_wrap", "wdNotARealWrap"));
                JObject warnFmt = Para(s, id, code["pf-apply-warn。"]);
                Ok(run, "pf-apply-warn",
                    warn.Success && Eq(warnFmt, "alignment", "center") && WarnHas(warn, "textbox_tight_wrap"),
                    Why(warn, warnFmt));
                CheckApply(run, s, id, code, "pf-apply-chars", Map("character_unit_first_line_indent", 2),
                    p => Near(p, "character_unit_first_line_indent", 2, 0.3));
                CheckApply(run, s, id, code, "pf-apply-keep", Map("keep_together", true), p => Flag(p, "keep_together"));
                CheckApply(run, s, id, code, "pf-apply-page", Map("page_break_before", true), p => Flag(p, "page_break_before"));
                CheckApply(run, s, id, code, "pf-apply-shade",
                    new Dictionary<string, object>
                    {
                        ["shading"] = new Dictionary<string, object> { ["background_pattern_color"] = "#00FF00" }
                    },
                    p => Contains(Deep(p, "shading", "background_pattern_color"), "00FF00"));
                CheckApply(run, s, id, code, "pf-apply-border",
                    Border("top", Map("line_style", "wdLineStyleSingle")),
                    p => Contains(Deep(p, "borders", "top", "line_style"), "Single"));
                ToolResult tabsFirst = Apply(s, id, code["pf-apply-tabs。"], Tabs(36, 72));
                ToolResult tabsSecond = Apply(s, id, code["pf-apply-tabs。"], Tabs(48));
                JObject tabsFmt = Para(s, id, code["pf-apply-tabs。"]);
                Ok(run, "pf-apply-tabs", tabsFirst.Success && tabsSecond.Success && OnlyTabNear(tabsFmt, 48), Why(tabsSecond, tabsFmt));
                ToolResult clearFirst = Apply(s, id, code["pf-apply-tabs-clear。"], Tabs(36));
                ToolResult clearSecond = Apply(s, id, code["pf-apply-tabs-clear。"], new Dictionary<string, object>
                {
                    ["tab_stops"] = new List<Dictionary<string, object>>()
                });
                JObject clearFmt = Para(s, id, code["pf-apply-tabs-clear。"]);
                Ok(run, "pf-apply-tabs-clear", clearFirst.Success && clearSecond.Success && !Has(clearFmt, "tab_stops"), Why(clearSecond, clearFmt));
                CheckApply(run, s, id, code, "pf-apply-east", Map("add_space_between_far_east_and_alpha", false),
                    p => Has(p, "add_space_between_far_east_and_alpha") && !Flag(p, "add_space_between_far_east_and_alpha"));
                CheckApply(run, s, id, code, "pf-apply-med", Map("alignment", "justify_med"), p => Eq(p, "alignment", "justify_med"));
                ToolResult items = s.ApplyParagraphFormat(new Dictionary<string, object>
                {
                    ["channel_id"] = id,
                    ["items"] = new object[]
                    {
                        new Dictionary<string, object>
                        {
                            ["target_paragraph_codes"] = code["pf-items-午。"],
                            ["para_format"] = Map("alignment", "center")
                        },
                        new Dictionary<string, object>
                        {
                            ["target_paragraph_codes"] = code["pf-items-未。"],
                            ["para_format"] = Map("alignment", "right")
                        }
                    }
                });
                JObject wu = Para(s, id, code["pf-items-午。"]);
                JObject wei = Para(s, id, code["pf-items-未。"]);
                string afterItems = ReadBody(s, id);
                Ok(run, "pf-items-two",
                    items.Success && Eq(wu, "alignment", "center") && Eq(wei, "alignment", "right")
                    && afterItems.IndexOf("pf-items-午。", StringComparison.Ordinal) >= 0
                    && afterItems.IndexOf("pf-items-未。", StringComparison.Ordinal) >= 0,
                    Why(items, wu) + " | " + Dump(wei));
                CheckApply(run, s, id, code, "pf-unit", Map("character_unit_first_line_indent", 2),
                    p => Near(p, "character_unit_first_line_indent", 2, 0.3));
                ToolResult autoFirst = Apply(s, id, code["pf-auto。"], Map("space_before", 12));
                ToolResult autoSecond = Apply(s, id, code["pf-auto。"], Map("space_before_auto", true));
                JObject autoFmt = Para(s, id, code["pf-auto。"]);
                Ok(run, "pf-auto", autoFirst.Success && autoSecond.Success && Flag(autoFmt, "space_before_auto") && !Has(autoFmt, "space_before"), Why(autoSecond, autoFmt));
                CheckApply(run, s, id, code, "pf-apply-justify", Map("alignment", "justify"), p => Eq(p, "alignment", "justify"));
                CheckApply(run, s, id, code, "pf-apply-distribute", Map("alignment", "distribute"), p => Eq(p, "alignment", "distribute"));
                CheckApply(run, s, id, code, "pf-apply-hi", Map("alignment", "justify_hi"), p => Eq(p, "alignment", "justify_hi"));
                CheckApply(run, s, id, code, "pf-apply-low", Map("alignment", "justify_low"), p => Eq(p, "alignment", "justify_low"));
                CheckApply(run, s, id, code, "pf-apply-thai", Map("alignment", "thai_justify"), p => Eq(p, "alignment", "thai_justify"));
                CheckApply(run, s, id, code, "pf-apply-right", Map("right_indent", 18), p => Near(p, "right_indent", 18, 1.2));
                CheckApply(run, s, id, code, "pf-apply-after", Map("space_after", 8), p => Near(p, "space_after", 8, 1.2));
                CheckApply(run, s, id, code, "pf-apply-cu-left", Map("character_unit_left_indent", 2), p => Near(p, "character_unit_left_indent", 2, 0.3));
                CheckApply(run, s, id, code, "pf-apply-cu-right", Map("character_unit_right_indent", 1), p => Near(p, "character_unit_right_indent", 1, 0.3));
                CheckApply(run, s, id, code, "pf-apply-lu-before", Map("line_unit_before", 1), p => Near(p, "line_unit_before", 1, 0.3));
                CheckApply(run, s, id, code, "pf-apply-lu-after", Map("line_unit_after", 1), p => Near(p, "line_unit_after", 1, 0.3));
                ToolResult afterFirst = Apply(s, id, code["pf-auto-after。"], Map("space_after", 12));
                ToolResult afterSecond = Apply(s, id, code["pf-auto-after。"], Map("space_after_auto", true));
                JObject afterFmt = Para(s, id, code["pf-auto-after。"]);
                Ok(run, "pf-auto-after", afterFirst.Success && afterSecond.Success && Flag(afterFmt, "space_after_auto") && !Has(afterFmt, "space_after"), Why(afterSecond, afterFmt));
                CheckApply(run, s, id, code, "pf-apply-noline", Map("no_line_number", true), p => Flag(p, "no_line_number"));
                CheckApply(run, s, id, code, "pf-apply-mirror", Map("mirror_indents", true), p => Flag(p, "mirror_indents"));
                CheckApply(run, s, id, code, "pf-apply-grid", Map("disable_line_height_grid", true), p => Flag(p, "disable_line_height_grid"));
                CheckApply(run, s, id, code, "pf-apply-half", Map("half_width_punctuation_on_top_of_line", true), p => Flag(p, "half_width_punctuation_on_top_of_line"));
                CheckApply(run, s, id, code, "pf-apply-wrap", Map("word_wrap", true), p => Flag(p, "word_wrap"));
                CheckApply(run, s, id, code, "pf-apply-collapse", Map("collapsed_by_default", true), p => Flag(p, "collapsed_by_default"));
                CheckApply(run, s, id, code, "pf-apply-hyphen", Map("hyphenation", false), p => Has(p, "hyphenation") && !Flag(p, "hyphenation"));
                CheckApply(run, s, id, code, "pf-apply-adjright", Map("auto_adjust_right_indent", false), p => Has(p, "auto_adjust_right_indent") && !Flag(p, "auto_adjust_right_indent"));
                CheckApply(run, s, id, code, "pf-apply-fe-break", Map("far_east_line_break_control", false), p => Has(p, "far_east_line_break_control") && !Flag(p, "far_east_line_break_control"));
                CheckApply(run, s, id, code, "pf-apply-hangpunct", Map("hanging_punctuation", false), p => Has(p, "hanging_punctuation") && !Flag(p, "hanging_punctuation"));
                CheckApply(run, s, id, code, "pf-apply-digit", Map("add_space_between_far_east_and_digit", false), p => Has(p, "add_space_between_far_east_and_digit") && !Flag(p, "add_space_between_far_east_and_digit"));
                CheckApply(run, s, id, code, "pf-apply-baseline", Map("baseline_alignment", "wdBaselineAlignCenter"), p => Contains(Str(p, "baseline_alignment"), "Center"));
                CheckApply(run, s, id, code, "pf-apply-rtl", Map("reading_order", "wdReadingOrderRtl"), p => Contains(Str(p, "reading_order"), "Rtl"));
                CheckApply(run, s, id, code, "pf-apply-tight", Map("textbox_tight_wrap", "wdTightAllLines"), p => Contains(Str(p, "textbox_tight_wrap"), "TightAllLines"));
                CheckApply(run, s, id, code, "pf-apply-texture",
                    new Dictionary<string, object> { ["shading"] = Map("texture", "wdTexture5Percent") },
                    p => Contains(Deep(p, "shading", "texture"), "5Percent"));
                CheckApply(run, s, id, code, "pf-apply-fore",
                    new Dictionary<string, object> { ["shading"] = Map("foreground_pattern_color", "#0000FF") },
                    p => Contains(Deep(p, "shading", "foreground_pattern_color"), "0000FF"));
                CheckApply(run, s, id, code, "pf-apply-border-sides",
                    new Dictionary<string, object>
                    {
                        ["borders"] = new Dictionary<string, object>
                        {
                            ["bottom"] = Map("line_style", "wdLineStyleSingle"),
                            ["left"] = Map("line_style", "wdLineStyleSingle"),
                            ["right"] = Map("line_style", "wdLineStyleSingle"),
                            ["horizontal"] = Map("line_style", "wdLineStyleSingle")
                        }
                    },
                    p => Contains(Deep(p, "borders", "bottom", "line_style"), "Single")
                        && Contains(Deep(p, "borders", "left", "line_style"), "Single")
                        && Contains(Deep(p, "borders", "right", "line_style"), "Single")
                        && Contains(Deep(p, "borders", "horizontal", "line_style"), "Single"));
                CheckApply(run, s, id, code, "pf-apply-border-ink",
                    Border("bottom", Map("line_style", "wdLineStyleSingle", "line_width", "wdLineWidth150pt", "color", "#FF0000")),
                    p => Contains(Deep(p, "borders", "bottom", "line_width"), "150")
                        && Contains(Deep(p, "borders", "bottom", "color"), "FF0000"));
                CheckApply(run, s, id, code, "pf-apply-tab-look",
                    new Dictionary<string, object>
                    {
                        ["tab_stops"] = new List<Dictionary<string, object>>
                        {
                            Map("position", 72, "alignment", "wdAlignTabCenter", "leader", "wdTabLeaderDots")
                        }
                    },
                    p =>
                    {
                        JArray tabs = p?["tab_stops"] as JArray;
                        if (tabs == null || tabs.Count != 1)
                        {
                            return false;
                        }

                        JObject tab = tabs[0] as JObject;
                        return Near(tab, "position", 72, 1.5)
                            && Contains(Str(tab, "alignment"), "Center")
                            && Contains(Str(tab, "leader"), "Dots");
                    });
                string endBody = ReadBody(s, id);
                Ok(run, "pf-text",
                    endBody.IndexOf("pf-apply-center。", StringComparison.Ordinal) >= 0
                    && endBody.IndexOf("pf-apply-15。", StringComparison.Ordinal) >= 0
                    && endBody.IndexOf("pf-apply-partial。", StringComparison.Ordinal) >= 0,
                    "套用后正文丢了");
            }
            finally
            {
                s.CloseDocByPath(path);
            }
        }

        private static void RunInherit(TestRun run, WdSession s)
        {
            string path, id;
            if (!OpenPlain(s, "pf-apply-inherit。", out path, out id))
            {
                FailMany(run, "铺 inherit 稿失败", "pf-apply-inherit");
                return;
            }

            try
            {
                if (!Paint(s, path, "pf-apply-inherit。",
                    p => p.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter,
                    p => SetPf(p, "Alignment", AlignCenter)))
                {
                    FailMany(run, "COM 居中失败", "pf-apply-inherit");
                    return;
                }

                string code = ParaCode(ReadBody(s, id), "pf-apply-inherit。");
                ToolResult applied = Apply(s, id, code, Map("alignment", "inherit"));
                JObject fmt = Para(s, id, code);
                Ok(run, "pf-apply-inherit",
                    applied.Success && Eq(fmt, "alignment", "center") && !WarnHas(applied, "alignment"),
                    Why(applied, fmt));
            }
            finally
            {
                s.CloseDocByPath(path);
            }
        }

        private static void RunLists(TestRun run, WdSession s)
        {
            OneList(run, s, "pf-list-number", "pf-list-number。", List("kind", "number"),
                (fmt, _) => Eq(Child(fmt, "list"), "kind", "number") && !string.IsNullOrEmpty(Str(Child(fmt, "list"), "list_string")));
            OneList(run, s, "pf-list-level", "pf-list-level。", List("kind", "outline", "level", 2),
                (fmt, _) => Eq(Child(fmt, "list"), "kind", "outline") && Near(Child(fmt, "list"), "level", 2, 0.1));
            TwoList(run, s, "pf-list-restart", "pf-list-restart甲。", "pf-list-restart乙。",
                List("kind", "number"), List("kind", "number", "continue_previous", false),
                (a, b, _) => FirstNumber(Str(Child(b, "list"), "list_string")) == "1");
            TwoList(run, s, "pf-list-shared", "pf-list-shared甲。", "pf-list-shared乙。",
                List("kind", "number"), List("kind", "number", "number_format", "%1)"),
                (a, b, _) => Contains(Str(Child(a, "list"), "number_format"), ")")
                    && Contains(Str(Child(b, "list"), "number_format"), ")"));
            OneList(run, s, "pf-list-bullet", "pf-list-bullet。", List("kind", "bullet"),
                (fmt, _) => Eq(Child(fmt, "list"), "kind", "bullet")
                    && !string.Equals((Str(Child(fmt, "list"), "list_string") ?? "").Trim(), "1.", StringComparison.Ordinal));
            RunListRemove(run, s);
            RunListString(run, s);
            OneList(run, s, "pf-list-style", "pf-list-style。", List("kind", "number", "number_style", "wdListNumberStyleUppercaseLetter"),
                (fmt, _) => Contains(Str(Child(fmt, "list"), "number_style"), "UppercaseLetter"));
            OneList(run, s, "pf-list-start", "pf-list-start。", List("kind", "number", "start_at", 5),
                (fmt, _) => Near(Child(fmt, "list"), "start_at", 5, 0.1)
                    && Contains(Str(Child(fmt, "list"), "list_string"), "5"));
            OneList(run, s, "pf-list-glyph", "pf-list-glyph。", List("kind", "bullet", "bullet_char", "●"),
                (fmt, _) => Eq(Child(fmt, "list"), "kind", "bullet")
                    && string.Equals((Str(Child(fmt, "list"), "bullet_char") ?? "").Trim(), "●", StringComparison.Ordinal));
            RunPicture(run, s);
            TwoList(run, s, "pf-list-continue", "pf-list-continue甲。", "pf-list-continue乙。",
                List("kind", "number"), List("kind", "number", "continue_previous", true),
                (a, b, _) => FirstNumber(Str(Child(b, "list"), "list_string")) == "2");
            RunReadonlyKinds(run, s);
        }

        private static void RunListRemove(TestRun run, WdSession s)
        {
            string path, id;
            if (!OpenPlain(s, "pf-list-remove。", out path, out id))
            {
                FailMany(run, "铺去掉编号稿失败", "pf-list-remove");
                return;
            }

            try
            {
                string code = ParaCode(ReadBody(s, id), "pf-list-remove。");
                ToolResult on = Apply(s, id, code, List("kind", "number"));
                ToolResult off = Apply(s, id, code, List("kind", "none"));
                JObject fmt = Para(s, id, code);
                Ok(run, "pf-list-remove", on.Success && off.Success && !Has(fmt, "list"), Why(off, fmt));
            }
            finally
            {
                s.CloseDocByPath(path);
            }
        }

        private static void RunListString(TestRun run, WdSession s)
        {
            string path, id;
            if (!OpenPlain(s, "pf-list-string。", out path, out id))
            {
                FailMany(run, "铺 list_string 稿失败", "pf-list-string");
                return;
            }

            try
            {
                string code = ParaCode(ReadBody(s, id), "pf-list-string。");
                ToolResult applied = Apply(s, id, code, List("kind", "number", "list_string", "ZZZ"));
                JObject fmt = Para(s, id, code);
                Ok(run, "pf-list-string",
                    applied.Success && WarnHas(applied, "list") && !string.Equals(Str(Child(fmt, "list"), "list_string"), "ZZZ", StringComparison.Ordinal),
                    Why(applied, fmt));
            }
            finally
            {
                s.CloseDocByPath(path);
            }
        }

        private static void RunPicture(TestRun run, WdSession s)
        {
            string png = Path.Combine(Path.GetTempPath(), "ew-pf-" + Guid.NewGuid().ToString("N") + ".png");
            File.WriteAllBytes(png, Convert.FromBase64String(
                "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg=="));
            string path, id;
            if (!OpenPlain(s, "pf-list-picture。", out path, out id))
            {
                FailMany(run, "铺图片项目符号稿失败", "pf-list-picture");
                return;
            }

            try
            {
                string code = ParaCode(ReadBody(s, id), "pf-list-picture。");
                ToolResult applied = Apply(s, id, code, List("kind", "picture_bullet", "picture_path", png));
                JObject fmt = Para(s, id, code);
                Ok(run, "pf-list-picture", applied.Success && Eq(Child(fmt, "list"), "kind", "picture_bullet"), Why(applied, fmt));
            }
            finally
            {
                s.CloseDocByPath(path);
                try
                {
                    File.Delete(png);
                }
                catch
                {
                }
            }
        }

        private static void RunReadonlyKinds(TestRun run, WdSession s)
        {
            string path, id;
            if (!OpenPlain(s, "pf-list-readonly甲。\rpf-list-readonly乙。", out path, out id))
            {
                FailMany(run, "铺只读 kind 稿失败", "pf-list-readonly");
                return;
            }

            try
            {
                string body = ReadBody(s, id);
                string a = ParaCode(body, "pf-list-readonly甲。");
                string b = ParaCode(body, "pf-list-readonly乙。");
                ToolResult mixed = Apply(s, id, a, List("kind", "mixed"));
                ToolResult only = Apply(s, id, b, List("kind", "list_num_only"));
                string kindA = Str(Child(Para(s, id, a), "list"), "kind");
                string kindB = Str(Child(Para(s, id, b), "list"), "kind");
                Ok(run, "pf-list-readonly",
                    mixed.Success && only.Success && WarnHas(mixed, "list") && WarnHas(only, "list")
                    && kindA != "mixed" && kindB != "list_num_only",
                    Why(mixed, Para(s, id, a)) + " | " + Why(only, Para(s, id, b)));
            }
            finally
            {
                s.CloseDocByPath(path);
            }
        }

        private static void RunDrop(TestRun run, WdSession s)
        {
            string[] ids = { "pf-drop", "pf-drop-off", "pf-drop-look" };
            string path, id;
            if (!OpenPlain(s, "pf-drop。\rpf-drop-look。", out path, out id))
            {
                FailMany(run, "铺首字下沉稿失败", ids);
                return;
            }

            try
            {
                string body = ReadBody(s, id);
                string drop = ParaCode(body, "pf-drop。");
                string look = ParaCode(body, "pf-drop-look。");
                ToolResult on = Apply(s, id, drop, Drop(Map("position", "drop", "lines", 3)));
                JObject onFmt = Para(s, id, drop);
                Ok(run, "pf-drop",
                    on.Success && Eq(Child(onFmt, "drop_cap"), "position", "drop") && Near(Child(onFmt, "drop_cap"), "lines", 3, 0.1),
                    Why(on, onFmt));
                ToolResult lookApply = Apply(s, id, look, Drop(Map("position", "drop", "lines", 3, "font_name", "宋体", "distance", 6)));
                JObject lookFmt = Child(Para(s, id, look), "drop_cap");
                Ok(run, "pf-drop-look",
                    lookApply.Success && Eq(lookFmt, "position", "drop") && Near(lookFmt, "lines", 3, 0.1)
                    && Contains(Str(lookFmt, "font_name"), "宋体") && Near(lookFmt, "distance", 6, 1.2),
                    Why(lookApply, lookFmt));
                ToolResult off = Apply(s, id, drop, Drop(Map("position", "none")));
                JObject offFmt = Para(s, id, drop);
                Ok(run, "pf-drop-off", off.Success && !Has(offFmt, "drop_cap"), Why(off, offFmt));
            }
            finally
            {
                s.CloseDocByPath(path);
            }
        }

        private static void RunDropTable(TestRun run, WdSession s)
        {
            string path, id;
            if (!OpenPlain(s, "表外甲。", out path, out id))
            {
                FailMany(run, "铺表内稿失败", "pf-drop-table");
                return;
            }

            try
            {
                if (!s.InsertTable(path, 1, 1, new[] { "表内甲。" }))
                {
                    FailMany(run, "插表失败", "pf-drop-table");
                    return;
                }

                string code = ParaCode(ReadBody(s, id), "表内甲。");
                ToolResult applied = Apply(s, id, code, new Dictionary<string, object>
                {
                    ["alignment"] = "center",
                    ["drop_cap"] = Map("position", "drop", "lines", 3)
                });
                JObject fmt = Para(s, id, code);
                bool dropped = Eq(Child(fmt, "drop_cap"), "position", "drop");
                Ok(run, "pf-drop-table",
                    applied.Success && Eq(fmt, "alignment", "center") && (dropped || WarnHas(applied, "drop_cap")),
                    Why(applied, fmt));
            }
            finally
            {
                s.CloseDocByPath(path);
            }
        }

        private static void OneList(
            TestRun run,
            WdSession s,
            string caseId,
            string sentence,
            Dictionary<string, object> format,
            Func<JObject, ToolResult, bool> check)
        {
            string path, id;
            if (!OpenPlain(s, sentence, out path, out id))
            {
                FailMany(run, "铺列表稿失败", caseId);
                return;
            }

            try
            {
                string code = ParaCode(ReadBody(s, id), sentence);
                ToolResult applied = Apply(s, id, code, format);
                JObject fmt = Para(s, id, code);
                Ok(run, caseId, applied.Success && check(fmt, applied), Why(applied, fmt));
            }
            finally
            {
                s.CloseDocByPath(path);
            }
        }

        private static void TwoList(
            TestRun run,
            WdSession s,
            string caseId,
            string first,
            string second,
            Dictionary<string, object> firstFmt,
            Dictionary<string, object> secondFmt,
            Func<JObject, JObject, ToolResult, bool> check)
        {
            string path, id;
            if (!OpenPlain(s, first + "\r" + second, out path, out id))
            {
                FailMany(run, "铺双段列表稿失败", caseId);
                return;
            }

            try
            {
                string body = ReadBody(s, id);
                string a = ParaCode(body, first);
                string b = ParaCode(body, second);
                ToolResult firstApply = Apply(s, id, a, firstFmt);
                ToolResult secondApply = Apply(s, id, b, secondFmt);
                JObject fa = Para(s, id, a);
                JObject fb = Para(s, id, b);
                Ok(run, caseId, firstApply.Success && secondApply.Success && check(fa, fb, secondApply), Why(secondApply, fb));
            }
            finally
            {
                s.CloseDocByPath(path);
            }
        }

        private static void CheckApply(
            TestRun run,
            WdSession s,
            string id,
            Dictionary<string, string> code,
            string caseId,
            Dictionary<string, object> format,
            Func<JObject, bool> check)
        {
            string para = code[caseId + "。"];
            ToolResult applied = Apply(s, id, para, format);
            JObject fmt = Para(s, id, para);
            Ok(run, caseId, applied.Success && check(fmt), Why(applied, fmt));
        }

        private static void SetHeading(Word.Paragraph para)
        {
            object index = Word.WdBuiltinStyle.wdStyleHeading1;
            Word.Style style = para.Range.Document.Styles.get_Item(ref index);
            object set = style;
            para.Range.set_Style(ref set);
        }

        private static void SetHeadingLate(object para)
        {
            object range = WdSession.ComGet(para, "Range");
            object doc = WdSession.ComGet(range, "Document");
            object styles = WdSession.ComGet(doc, "Styles");
            object style = WdSession.ComCall(styles, "Item", Heading1);
            WdSession.ComSet(range, "Style", style);
        }

        private static void AddLeftTab(Word.Paragraph para)
        {
            object align = Word.WdTabAlignment.wdAlignTabLeft;
            object leader = Word.WdTabLeader.wdTabLeaderSpaces;
            para.Range.ParagraphFormat.TabStops.Add(72f, ref align, ref leader);
        }

        private static void AddLeftTabLate(object para)
        {
            object stops = WdSession.ComGet(Pf(para), "TabStops");
            WdSession.ComCall(stops, "Add", 72f, 0, 0);
        }

        private static object Pf(object para)
        {
            return WdSession.ComGet(WdSession.ComGet(para, "Range"), "ParagraphFormat");
        }

        private static void SetPf(object para, string name, object value)
        {
            WdSession.ComSet(Pf(para), name, value);
        }

        private static bool Paint(WdSession s, string path, string needle, Action<Word.Paragraph> word, Action<object> wps)
        {
            try
            {
                return s.IsWps ? s.EditParagraph(path, needle, wps) : s.EditWordParagraph(path, needle, word);
            }
            catch
            {
                return false;
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

        private static ToolResult Apply(WdSession s, string id, string code, Dictionary<string, object> format)
        {
            return s.ApplyParagraphFormat(new Dictionary<string, object>
            {
                ["channel_id"] = id,
                ["target_paragraph_codes"] = code,
                ["para_format"] = format
            });
        }

        private static Dictionary<string, object> Map(params object[] pairs)
        {
            var map = new Dictionary<string, object>();
            for (int i = 0; i + 1 < pairs.Length; i += 2)
            {
                map[(string)pairs[i]] = pairs[i + 1];
            }

            return map;
        }

        private static Dictionary<string, object> List(params object[] pairs)
        {
            return new Dictionary<string, object> { ["list"] = Map(pairs) };
        }

        private static Dictionary<string, object> Drop(Dictionary<string, object> drop)
        {
            return new Dictionary<string, object> { ["drop_cap"] = drop };
        }

        private static Dictionary<string, object> Border(string side, Dictionary<string, object> edge)
        {
            return new Dictionary<string, object>
            {
                ["borders"] = new Dictionary<string, object> { [side] = edge }
            };
        }

        private static Dictionary<string, object> Tabs(params int[] positions)
        {
            var list = new List<Dictionary<string, object>>();
            foreach (int position in positions)
            {
                list.Add(Map("position", position));
            }

            return new Dictionary<string, object> { ["tab_stops"] = list };
        }

        private static string ReadBody(WdSession s, string id)
        {
            ToolResult got = s.GetContent(id, "paragraph");
            return got.Success ? WdSession.Str(got, "document_content") : "";
        }

        private static string ParaCode(string body, string sentence)
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

            int mark = body.LastIndexOf("[P_", at, StringComparison.Ordinal);
            if (mark < 0)
            {
                return "";
            }

            Match m = Regex.Match(body.Substring(mark), @"\[P_([0-9a-zA-Z]{5}(?:_[1-9][0-9]*)?)(?:-start| start)\]");
            return m.Success ? "P_" + m.Groups[1].Value : "";
        }

        private static JObject Para(WdSession s, string id, string code)
        {
            if (string.IsNullOrEmpty(code))
            {
                return null;
            }

            return Obj(FirstEntry(s.GetFormatContext(new Dictionary<string, object>
            {
                ["channel_id"] = id,
                ["target_paragraph_codes"] = code
            })), "para_format");
        }

        private static readonly string[] DefaultKeys =
        {
            "alignment", "line_spacing_rule", "line_spacing",
            "left_indent", "right_indent", "first_line_indent", "hanging_indent",
            "space_before", "space_after", "paragraph_style",
            "outline_level", "baseline_alignment", "reading_order", "textbox_tight_wrap",
            "character_unit_left_indent", "character_unit_right_indent", "character_unit_first_line_indent",
            "line_unit_before", "line_unit_after",
            "widow_control", "keep_with_next", "keep_together", "page_break_before", "no_line_number",
            "hyphenation", "mirror_indents", "auto_adjust_right_indent",
            "space_before_auto", "space_after_auto", "disable_line_height_grid",
            "far_east_line_break_control", "word_wrap", "hanging_punctuation",
            "half_width_punctuation_on_top_of_line",
            "add_space_between_far_east_and_alpha", "add_space_between_far_east_and_digit",
            "collapsed_by_default"
        };

        private static bool MissingDefaults(JObject para)
        {
            if (para == null)
            {
                return false;
            }

            foreach (string key in DefaultKeys)
            {
                if (Has(para, key))
                {
                    return false;
                }
            }

            return true;
        }

        private static bool TabNear(JObject para, double expect)
        {
            JArray tabs = para?["tab_stops"] as JArray;
            if (tabs == null)
            {
                return false;
            }

            foreach (JToken token in tabs)
            {
                if (Near(token as JObject, "position", expect, 1.5))
                {
                    return true;
                }
            }

            return false;
        }

        private static bool OnlyTabNear(JObject para, double expect)
        {
            JArray tabs = para?["tab_stops"] as JArray;
            return tabs != null && tabs.Count == 1 && Near(tabs[0] as JObject, "position", expect, 1.5);
        }

        private static string FirstNumber(string text)
        {
            Match m = Regex.Match(text ?? "", @"\d+");
            return m.Success ? m.Value : "";
        }

        private static bool WarnHas(ToolResult result, string key)
        {
            JArray warnings = DataObj(result)?["format_warnings"] as JArray;
            if (warnings == null)
            {
                return false;
            }

            foreach (JToken token in warnings)
            {
                if (string.Equals(Str(token as JObject, "key"), key, StringComparison.Ordinal))
                {
                    return true;
                }
            }

            return false;
        }

        private static string Why(ToolResult result, JToken fmt)
        {
            if (result != null && !result.Success)
            {
                return result.Error ?? "套用失败";
            }

            return Dump(fmt);
        }

        private static JObject DataObj(ToolResult result)
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
            JArray entries = DataObj(result)?["entries"] as JArray;
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

        private static JObject Child(JObject parent, string key)
        {
            return Obj(parent, key);
        }

        private static string Deep(JObject root, params string[] path)
        {
            JToken cur = root;
            foreach (string key in path)
            {
                if (cur == null)
                {
                    return "";
                }

                cur = cur[key];
            }

            return cur == null || cur.Type == JTokenType.Null ? "" : cur.ToString();
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

        private static bool Eq(JObject o, string key, string expect)
        {
            return string.Equals(Str(o, key), expect, StringComparison.OrdinalIgnoreCase);
        }

        private static bool Near(JObject o, string key, double expect, double tol)
        {
            if (!Has(o, key))
            {
                return false;
            }

            double value;
            if (!double.TryParse(o[key].ToString(), NumberStyles.Float, CultureInfo.InvariantCulture, out value)
                && !double.TryParse(o[key].ToString(), out value))
            {
                return false;
            }

            return Math.Abs(value - expect) <= tol;
        }

        private static bool Contains(string text, string needle)
        {
            return !string.IsNullOrEmpty(text) && text.IndexOf(needle, StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private static bool Blob(JToken token, string needle)
        {
            return token != null && token.ToString(Formatting.None).IndexOf(needle, StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private static string Dump(JToken token)
        {
            return token == null ? "(空)" : token.ToString(Formatting.None);
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
