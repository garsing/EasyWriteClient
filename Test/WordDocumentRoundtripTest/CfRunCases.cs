using System;
using System.Collections.Generic;
using System.Text;
using Newtonsoft.Json.Linq;
using WordAddIn1;

namespace WordDocumentRoundtripTest
{
    internal static partial class CfCases
    {
        private const string SpanBold = "易写区间整句加粗。";
        private const string SpanHeadBold = "易写区间头两加粗。";
        private const string SpanHeadItalic = "易写区间头两斜体。";
        private const string SpanHeadRed = "易写区间头两红色。";
        private const string SpanHeadSuper = "易写区间头两上标。";
        private const string SpanHeadSub = "易写区间头两下标。";
        private const string SpanHeadStrike = "易写区间头两删除。";
        private const string SpanHeadSize = "易写区间头两字号。";
        private const string SpanHeadName = "易写区间头两黑体。";
        private const string SpanHeadLine = "易写区间头两下划。";
        private const string SpanThree = "易写区间三段混排。";
        private const string SpanMerge = "易写区间相邻合并。";
        private const string SpanGap = "易写区间隔开加粗。";
        private const string SpanMiddle = "易写区间中间红色。";
        private const string SpanHi = "易写区间头两高亮。";
        private const string SpanColors = "易写区间两头异色。";
        private const string SpanBoth = "易写区间头两双饰。";
        private const string SpanKai = "易写区间整句楷体。";
        private const string SpanSpace = "易写区间头两间距。";
        private const string SpanCaps = "易写区间头两全大。";

        private const int ColorBlue = 16711680;

        private static void RunSpans(TestRun run, WdSession s)
        {
            string[] sentences =
            {
                SpanBold, SpanHeadBold, SpanHeadItalic, SpanHeadRed, SpanHeadSuper, SpanHeadSub,
                SpanHeadStrike, SpanHeadSize, SpanHeadName, SpanHeadLine, SpanThree, SpanMerge,
                SpanGap, SpanMiddle, SpanHi, SpanColors, SpanBoth, SpanKai, SpanSpace, SpanCaps
            };
            string[] ids =
            {
                "cf-run-bold-one", "cf-run-head-bold", "cf-run-head-italic", "cf-run-head-red",
                "cf-run-head-super", "cf-run-head-sub", "cf-run-head-strike", "cf-run-head-size",
                "cf-run-head-name", "cf-run-head-line", "cf-run-three", "cf-run-merge",
                "cf-run-gap", "cf-run-join", "cf-run-shape", "cf-run-ten",
                "cf-run-clean", "cf-run-size-split", "cf-run-no-index", "cf-run-middle",
                "cf-run-highlight", "cf-run-two-color", "cf-run-both", "cf-run-super-only",
                "cf-run-para", "cf-run-batch", "cf-run-kai", "cf-run-spacing",
                "cf-run-caps", "cf-run-tail-text"
            };
            string path, id;
            if (!OpenPlain(s, string.Join("\r", sentences), out path, out id))
            {
                FailMany(run, "铺区间稿失败", ids);
                return;
            }

            try
            {
                var nameHei = new Dictionary<string, object> { ["Name"] = "黑体", ["NameFarEast"] = "黑体" };
                var nameKai = new Dictionary<string, object> { ["Name"] = "楷体", ["NameFarEast"] = "楷体" };
                var both = new Dictionary<string, object> { ["Bold"] = On, ["Italic"] = On };
                bool painted = s.PaintSlice(path, SpanBold, 0, 0, Font("Bold", On), null)
                    && s.PaintSlice(path, SpanHeadBold, 0, 2, Font("Bold", On), null)
                    && s.PaintSlice(path, SpanHeadItalic, 0, 2, Font("Italic", On), null)
                    && s.PaintSlice(path, SpanHeadRed, 0, 2, Font("Color", ColorRed), null)
                    && s.PaintSlice(path, SpanHeadSuper, 0, 2, Font("Superscript", On), null)
                    && s.PaintSlice(path, SpanHeadSub, 0, 2, Font("Subscript", On), null)
                    && s.PaintSlice(path, SpanHeadStrike, 0, 2, Font("StrikeThrough", On), null)
                    && s.PaintSlice(path, SpanHeadSize, 0, 2, Font("Size", 18), null)
                    && s.PaintSlice(path, SpanHeadName, 0, 2, nameHei, null)
                    && s.PaintSlice(path, SpanHeadLine, 0, 2, Font("Underline", 1), null)
                    && s.PaintSlice(path, SpanThree, 0, 2, Font("Bold", On), null)
                    && s.PaintSlice(path, SpanThree, 4, 0, Font("Italic", On), null)
                    && s.PaintSlice(path, SpanMerge, 0, 2, Font("Bold", On), null)
                    && s.PaintSlice(path, SpanMerge, 2, 2, Font("Bold", On), null)
                    && s.PaintSlice(path, SpanGap, 0, 2, Font("Bold", On), null)
                    && s.PaintSlice(path, SpanGap, 4, 0, Font("Bold", On), null)
                    && s.PaintSlice(path, SpanMiddle, 3, 2, Font("Color", ColorRed), null)
                    && s.PaintSlice(path, SpanHi, 0, 2, null, RangeProp("HighlightColorIndex", HighlightYellow))
                    && s.PaintSlice(path, SpanColors, 0, 2, Font("Color", ColorRed), null)
                    && s.PaintSlice(path, SpanColors, 6, 0, Font("Color", ColorBlue), null)
                    && s.PaintSlice(path, SpanBoth, 0, 2, both, null)
                    && s.PaintSlice(path, SpanKai, 0, 0, nameKai, null)
                    && s.PaintSlice(path, SpanSpace, 0, 2, Font("Spacing", 3), null)
                    && s.PaintSlice(path, SpanCaps, 0, 2, Font("AllCaps", On), null);
                if (!painted)
                {
                    FailMany(run, "COM 铺区间失败", ids);
                    return;
                }

                s.GetContent(id);
                string body = ReadBody(s, id);
                var code = new Dictionary<string, string>();
                foreach (string sentence in sentences)
                {
                    code[sentence] = CodeOf(body, sentence);
                }

                if (string.IsNullOrEmpty(code[SpanBold]) || string.IsNullOrEmpty(code[SpanCaps]))
                {
                    FailMany(run, "抄不到区间 S_", ids);
                    return;
                }

                JObject boldEntry = ReadEntry(s, id, code[SpanBold]);
                JArray boldRuns = Runs(boldEntry);
                Ok(run, "cf-run-bold-one",
                    boldRuns.Count == 1
                    && RunText(boldRuns[0]) == SpanBold
                    && Flag(RunCf(boldRuns[0]), "is_bold"),
                    Dump(boldEntry));

                JArray headBold = Runs(ReadEntry(s, id, code[SpanHeadBold]));
                Ok(run, "cf-run-head-bold",
                    headBold.Count == 2
                    && RunText(headBold[0]).Length == 2
                    && Flag(RunCf(headBold[0]), "is_bold")
                    && !Flag(RunCf(headBold[1]), "is_bold"),
                    Dump(headBold));

                JArray headItalic = Runs(ReadEntry(s, id, code[SpanHeadItalic]));
                Ok(run, "cf-run-head-italic",
                    headItalic.Count == 2
                    && Flag(RunCf(headItalic[0]), "is_italic")
                    && !Flag(RunCf(headItalic[1]), "is_italic"),
                    Dump(headItalic));

                JArray headRed = Runs(ReadEntry(s, id, code[SpanHeadRed]));
                Ok(run, "cf-run-head-red",
                    headRed.Count == 2
                    && string.Equals(Str(RunCf(headRed[0]), "font_color"), "#FF0000", StringComparison.OrdinalIgnoreCase)
                    && !string.Equals(Str(RunCf(headRed[1]), "font_color"), "#FF0000", StringComparison.OrdinalIgnoreCase),
                    Dump(headRed));

                JArray headSuper = Runs(ReadEntry(s, id, code[SpanHeadSuper]));
                Ok(run, "cf-run-head-super",
                    headSuper.Count == 2
                    && Flag(RunCf(headSuper[0]), "is_superscript")
                    && !Flag(RunCf(headSuper[1]), "is_superscript"),
                    Dump(headSuper));

                JArray headSub = Runs(ReadEntry(s, id, code[SpanHeadSub]));
                Ok(run, "cf-run-head-sub",
                    headSub.Count == 2
                    && Flag(RunCf(headSub[0]), "is_subscript")
                    && !Flag(RunCf(headSub[1]), "is_subscript"),
                    Dump(headSub));

                JArray headStrike = Runs(ReadEntry(s, id, code[SpanHeadStrike]));
                Ok(run, "cf-run-head-strike",
                    headStrike.Count == 2
                    && Flag(RunCf(headStrike[0]), "has_strikethrough")
                    && !Flag(RunCf(headStrike[1]), "has_strikethrough"),
                    Dump(headStrike));

                JArray headSize = Runs(ReadEntry(s, id, code[SpanHeadSize]));
                Ok(run, "cf-run-head-size",
                    headSize.Count == 2 && Near(Num(RunCf(headSize[0]), "font_size"), 18, 0.6),
                    Dump(headSize));

                JArray headName = Runs(ReadEntry(s, id, code[SpanHeadName]));
                Ok(run, "cf-run-head-name",
                    headName.Count == 2
                    && FontLooksLike(Str(RunCf(headName[0]), "font_name"), "黑")
                    && !FontLooksLike(Str(RunCf(headName[1]), "font_name"), "黑"),
                    Dump(headName));

                JArray headLine = Runs(ReadEntry(s, id, code[SpanHeadLine]));
                Ok(run, "cf-run-head-line",
                    headLine.Count == 2
                    && Flag(RunCf(headLine[0]), "has_underline")
                    && !Flag(RunCf(headLine[1]), "has_underline"),
                    Dump(headLine));

                JArray three = Runs(ReadEntry(s, id, code[SpanThree]));
                Ok(run, "cf-run-three",
                    three.Count == 3
                    && RunText(three[0]).Length == 2
                    && RunText(three[1]).Length == 2
                    && Flag(RunCf(three[0]), "is_bold")
                    && !Flag(RunCf(three[1]), "is_bold")
                    && !Flag(RunCf(three[1]), "is_italic")
                    && Flag(RunCf(three[2]), "is_italic")
                    && !Flag(RunCf(three[2]), "is_bold"),
                    Dump(three));

                JArray merged = Runs(ReadEntry(s, id, code[SpanMerge]));
                Ok(run, "cf-run-merge",
                    merged.Count == 2
                    && RunText(merged[0]).Length == 4
                    && Flag(RunCf(merged[0]), "is_bold")
                    && !Flag(RunCf(merged[1]), "is_bold"),
                    Dump(merged));

                JArray gap = Runs(ReadEntry(s, id, code[SpanGap]));
                Ok(run, "cf-run-gap",
                    gap.Count == 3
                    && Flag(RunCf(gap[0]), "is_bold")
                    && !Flag(RunCf(gap[1]), "is_bold")
                    && Flag(RunCf(gap[2]), "is_bold")
                    && RunText(gap[0]).Length == 2
                    && RunText(gap[1]).Length == 2,
                    Dump(gap));

                Ok(run, "cf-run-join", JoinRuns(three) == SpanThree, JoinRuns(three));

                JObject threeEntry = ReadEntry(s, id, code[SpanThree]);
                Ok(run, "cf-run-shape",
                    threeEntry != null
                    && threeEntry["format_note"] == null
                    && threeEntry["format_kind_count"] == null
                    && threeEntry["char_format"] == null
                    && Runs(threeEntry).Count == 3,
                    Dump(threeEntry));

                Ok(run, "cf-run-ten",
                    headBold.Count == 2 && HasTen(RunCf(headBold[0])) && HasTen(RunCf(headBold[1])),
                    Dump(headBold));

                Ok(run, "cf-run-clean", !Blob(threeEntry, "9999999") && !Blob(threeEntry, "9999998"), Dump(threeEntry));

                Ok(run, "cf-run-size-split",
                    headSize.Count == 2 && !Near(Num(RunCf(headSize[1]), "font_size"), 18, 0.6),
                    Dump(headSize));

                Ok(run, "cf-run-no-index",
                    headRed.Count == 2
                    && !Has(RunCf(headRed[0]), "color_index")
                    && !Has(RunCf(headRed[0]), "diacritic_color"),
                    Dump(headRed));

                JArray middle = Runs(ReadEntry(s, id, code[SpanMiddle]));
                Ok(run, "cf-run-middle",
                    middle.Count == 3
                    && RunText(middle[1]).Length == 2
                    && string.Equals(Str(RunCf(middle[1]), "font_color"), "#FF0000", StringComparison.OrdinalIgnoreCase)
                    && !string.Equals(Str(RunCf(middle[0]), "font_color"), "#FF0000", StringComparison.OrdinalIgnoreCase)
                    && !string.Equals(Str(RunCf(middle[2]), "font_color"), "#FF0000", StringComparison.OrdinalIgnoreCase),
                    Dump(middle));

                JArray hi = Runs(ReadEntry(s, id, code[SpanHi]));
                Ok(run, "cf-run-highlight",
                    hi.Count == 2
                    && Contains(Str(RunCf(hi[0]), "highlight_color_index"), "Yellow")
                    && !Has(RunCf(hi[1]), "highlight_color_index"),
                    Dump(hi));

                JArray colors = Runs(ReadEntry(s, id, code[SpanColors]));
                Ok(run, "cf-run-two-color",
                    colors.Count == 3
                    && string.Equals(Str(RunCf(colors[0]), "font_color"), "#FF0000", StringComparison.OrdinalIgnoreCase)
                    && string.Equals(Str(RunCf(colors[2]), "font_color"), "#0000FF", StringComparison.OrdinalIgnoreCase)
                    && !string.Equals(Str(RunCf(colors[1]), "font_color"), "#FF0000", StringComparison.OrdinalIgnoreCase)
                    && !string.Equals(Str(RunCf(colors[1]), "font_color"), "#0000FF", StringComparison.OrdinalIgnoreCase),
                    Dump(colors));

                JArray bothRuns = Runs(ReadEntry(s, id, code[SpanBoth]));
                Ok(run, "cf-run-both",
                    bothRuns.Count == 2
                    && Flag(RunCf(bothRuns[0]), "is_bold")
                    && Flag(RunCf(bothRuns[0]), "is_italic")
                    && !Flag(RunCf(bothRuns[1]), "is_bold")
                    && !Flag(RunCf(bothRuns[1]), "is_italic"),
                    Dump(bothRuns));

                Ok(run, "cf-run-super-only",
                    headSuper.Count == 2
                    && Flag(RunCf(headSuper[0]), "is_superscript")
                    && !Flag(RunCf(headSuper[0]), "is_subscript"),
                    Dump(headSuper));

                Ok(run, "cf-run-para",
                    boldEntry != null
                    && Has(boldEntry, "paragraph_code")
                    && Contains(Str(boldEntry, "text_preview"), "易写区间整句"),
                    Dump(boldEntry));

                ToolResult batch = s.GetFormatContext(new Dictionary<string, object>
                {
                    ["channel_id"] = id,
                    ["target_codes"] = code[SpanBold] + "," + code[SpanHeadBold]
                });
                JArray batchEntries = AllEntries(batch);
                Ok(run, "cf-run-batch",
                    batch.Success
                    && batchEntries.Count == 2
                    && JoinRuns(Runs((JObject)batchEntries[0])) == SpanBold
                    && JoinRuns(Runs((JObject)batchEntries[1])) == SpanHeadBold,
                    batch.Success ? Dump(batchEntries) : (batch.Error ?? ""));

                JArray kai = Runs(ReadEntry(s, id, code[SpanKai]));
                Ok(run, "cf-run-kai",
                    kai.Count == 1 && FontLooksLike(Str(RunCf(kai[0]), "font_name"), "楷"),
                    Dump(kai));

                JArray space = Runs(ReadEntry(s, id, code[SpanSpace]));
                Ok(run, "cf-run-spacing",
                    space.Count == 2
                    && Near(Num(RunCf(space[0]), "spacing"), 3, 0.6)
                    && !Has(RunCf(space[1]), "spacing"),
                    Dump(space));

                JArray caps = Runs(ReadEntry(s, id, code[SpanCaps]));
                Ok(run, "cf-run-caps",
                    caps.Count == 2
                    && Flag(RunCf(caps[0]), "all_caps")
                    && !Flag(RunCf(caps[1]), "all_caps"),
                    Dump(caps));

                Ok(run, "cf-run-tail-text",
                    headBold.Count == 2 && RunText(headBold[1]) == SpanHeadBold.Substring(2),
                    Dump(headBold));
            }
            finally
            {
                s.CloseDocByPath(path);
            }
        }

        private static JArray AllEntries(ToolResult result)
        {
            JObject data = DataObj(result);
            return data == null ? new JArray() : data["entries"] as JArray ?? new JArray();
        }

        private static string RunText(JToken run)
        {
            return run?["text"]?.ToString() ?? "";
        }

        private static JObject RunCf(JToken run)
        {
            return (run as JObject)?["char_format"] as JObject ?? new JObject();
        }

        private static string JoinRuns(JArray runs)
        {
            var text = new StringBuilder();
            if (runs == null)
            {
                return "";
            }

            foreach (JToken run in runs)
            {
                text.Append(RunText(run));
            }

            return text.ToString();
        }
    }
}
