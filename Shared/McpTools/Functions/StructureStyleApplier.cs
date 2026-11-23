using System;
using System.Collections.Generic;
using System.Linq;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    /// <summary>套用。先清这一侧的直接格式，再写样式。</summary>
    public static class StructureStyleApplier
    {
        public static Dictionary<string, object> Apply(
            Word.Document doc,
            Dictionary<string, object> item,
            IList<StructureStyleRecord> styles,
            bool disallowBackendApi)
        {
            ApplyPlan plan = Plan(doc, item, styles, disallowBackendApi);
            if (!plan.WillWrite)
            {
                return plan.Finished;
            }

            StructureStyleRecord style = plan.Style;
            if (style == null)
            {
                return Fail("没有样式 " + Arg(item, "structure") + "。");
            }

            string code = plan.Code;
            List<string> warnings = plan.Warnings;
            var paragraphs = new List<Word.Paragraph>();
            var owners = new List<Word.Range>();
            foreach (Word.Range range in plan.Ranges)
            {
                foreach (Word.Paragraph paragraph in range.Paragraphs)
                {
                    if (Intersects(range, paragraph))
                    {
                        paragraphs.Add(paragraph);
                        owners.Add(range);
                    }
                }
            }

            if (paragraphs.Count == 0)
            {
                return Fail("定位失败。");
            }

            bool paragraphStyle = style.Support == "both" || style.Support == "paragraph";
            bool characterStyle = style.Support == "both" || style.Support == "character";
            var seenParagraphs = new HashSet<int>();
            try
            {
                if (paragraphStyle)
                {
                    foreach (Word.Range range in plan.Ranges)
                    {
                        foreach (Word.Paragraph paragraph in range.Paragraphs)
                        {
                            if (!Intersects(range, paragraph) || paragraph.Range == null)
                            {
                                continue;
                            }

                            if (!seenParagraphs.Add(paragraph.Range.Start))
                            {
                                continue;
                            }

                            ApplyParagraphSide(doc, style, paragraph);
                        }
                    }
                }

                if (characterStyle)
                {
                    foreach (Word.Range range in plan.Ranges)
                    {
                        foreach (Word.Paragraph paragraph in range.Paragraphs)
                        {
                            if (!Intersects(range, paragraph))
                            {
                                continue;
                            }

                            Word.Range slice = Slice(doc, range, paragraph);
                            if (slice != null)
                            {
                                ApplySlice(doc, style, slice);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                return Fail(ex.Message);
            }

            string scope = "text";
            foreach (Word.Range range in plan.Ranges)
            {
                foreach (Word.Paragraph paragraph in range.Paragraphs)
                {
                    if (Intersects(range, paragraph) && CoversWhole(range, paragraph))
                    {
                        scope = "paragraph";
                    }
                }
            }

            return Done(code, style, scope, warnings, null);
        }

        private sealed class ApplyPlan
        {
            public bool WillWrite;
            public Dictionary<string, object> Finished;
            public StructureStyleRecord Style;
            public string Code;
            public List<string> Warnings;
            public List<Word.Range> Ranges;
        }

        private static ApplyPlan Plan(
            Word.Document doc,
            Dictionary<string, object> item,
            IList<StructureStyleRecord> styles,
            bool disallowBackendApi)
        {
            var plan = new ApplyPlan();
            string structure = Arg(item, "structure");
            if (string.IsNullOrWhiteSpace(structure))
            {
                plan.Finished = Fail("缺少 structure。");
                return plan;
            }

            StructureStyleRecord style = StructureStyleCatalog.Find(styles, structure);
            bool onHost = style != null || HostHasStyle(doc, structure);
            if (!onHost)
            {
                plan.Finished = Fail("没有样式 " + structure + "。");
                return plan;
            }

            string codeError = ReadStyleCodes(item, out string code, out List<string> codes, out bool paragraphMode);
            if (codeError != null)
            {
                plan.Finished = Fail(codeError);
                return plan;
            }
            string sibling = TextSpanArg.RejectSiblingIndex(item);
            if (sibling != null)
            {
                plan.Finished = Fail(sibling);
                return plan;
            }

            var warnings = new List<string>();
            if (item != null && (item.ContainsKey("text_format") || item.ContainsKey("paragraph_format")))
            {
                warnings.Add("套用忽略 text_format 和 paragraph_format，按样式定义套。");
            }

            bool narrowing = TextSpanArg.HasNarrowingText(item);
            if (narrowing && codes.Count != 1)
            {
                plan.Finished = Fail(TextSpanArg.SingleCodeError);
                return plan;
            }

            TextSpan span = null;
            if (item != null && item.ContainsKey("text"))
            {
                if (!TextSpanArg.TryRead(item, out span, out string parseError))
                {
                    plan.Finished = Fail(parseError ?? "text 无法解析。");
                    return plan;
                }
            }

            bool anyT = codes.Count > 0 && codes.All(EmptyCellCode.IsEmptyCellCode);
            if (anyT && codes.Count != 1)
            {
                plan.Finished = Fail("几何码一次一条。");
                return plan;
            }

            bool paragraph = paragraphMode;
            List<List<string>> runs;
            if (anyT)
            {
                runs = new List<List<string>> { codes };
            }
            else
            {
                string snapshot = paragraph
                    ? DocumentState.ParagraphSegSnapshot
                    : DocumentState.SegSnapshot;
                runs = FormatApplyBatchCompiler.SplitRuns(codes, snapshot);
                if (runs.Count == 0)
                {
                    runs.Add(codes);
                }
            }

            var locators = new List<Dictionary<string, object>>();
            foreach (List<string> run in runs)
            {
                var one = new Dictionary<string, object>
                {
                    [paragraph ? "paragraph_codes" : "codes"] = string.Join(",", run)
                };
                if (span != null && !string.IsNullOrEmpty(span.Value))
                {
                    one["text"] = span.Value;
                }

                locators.Add(one);
            }

            LocateProbeLog.Line(
                "apply.Plan codes=" + code
                + " paragraph=" + paragraph
                + " locators=" + locators.Count
                + " first=" + (locators.Count == 0 ? "" : string.Join(",", locators[0].Keys)));
            InsertTargetLocatorHelper.LocateSpanBatchResult batch = InsertTargetLocatorHelper.TryLocateSpans(
                doc,
                locators,
                disallowBackendApi);
            LocateProbeLog.Line(
                "apply.locate blocked=" + batch.BatchBlocked
                + " err=" + (batch.Error ?? "")
                + " items=" + (batch.Items == null ? 0 : batch.Items.Count)
                + " firstOk=" + (batch.Items != null && batch.Items.Count > 0 && batch.Items[0] != null && batch.Items[0].Success)
                + " " + TocFieldStamp.DumpInterior());
            if (batch.BatchBlocked)
            {
                plan.Finished = Fail(string.IsNullOrEmpty(batch.Error) ? "定位失败。" : batch.Error);
                return plan;
            }

            var ranges = new List<Word.Range>();
            foreach (InsertTargetLocatorHelper.LocateSpanItem located in batch.Items)
            {
                if (located == null || !located.Success || located.Range == null)
                {
                    plan.Finished = Fail(located == null || string.IsNullOrEmpty(located.Error) ? "定位失败。" : located.Error);
                    return plan;
                }

                ranges.Add(located.Range);
            }

            if (style == null)
            {
                plan.WillWrite = true;
                return plan;
            }

            plan.WillWrite = true;
            plan.Style = style;
            plan.Code = code;
            plan.Warnings = warnings;
            plan.Ranges = ranges;
            return plan;
        }

        private static string ReadStyleCodes(
            Dictionary<string, object> item,
            out string code,
            out List<string> codes,
            out bool paragraph)
        {
            code = "";
            codes = new List<string>();
            paragraph = false;
            string paraCsv = Arg(item, "target_paragraph_codes");
            string geoCsv = Arg(item, "target_codes");
            bool hasPara = !string.IsNullOrWhiteSpace(paraCsv);
            bool hasGeo = !string.IsNullOrWhiteSpace(geoCsv);
            if (hasPara && hasGeo)
            {
                return "target_codes 与 target_paragraph_codes 不能同时传";
            }

            if (hasPara)
            {
                codes = SplitCodes(paraCsv);
                code = paraCsv.Trim();
                paragraph = true;
                if (codes.Count == 0 || codes.Any(c => !FormatContextHelper.IsParagraphCode(c)))
                {
                    return "target_paragraph_codes 仅支持 P_";
                }

                return null;
            }

            if (!hasGeo)
            {
                return "缺少 target_paragraph_codes 或几何码 target_codes";
            }

            codes = SplitCodes(geoCsv);
            code = geoCsv.Trim();
            if (codes.Any(c => FormatContextHelper.IsSentenceCode(c)))
            {
                return InsertTargetLocatorHelper.CodesRejectSentenceError.Replace(
                    "paragraph_codes", "target_paragraph_codes");
            }

            if (codes.Any(c => FormatContextHelper.IsParagraphCode(c)))
            {
                return "正文请用 target_paragraph_codes";
            }

            if (codes.Count == 0 || !codes.All(EmptyCellCode.IsEmptyCellCode))
            {
                return "target_codes 仅支持一条几何码 T_…_r_c";
            }

            return null;
        }

        private static List<string> SplitCodes(string raw)
        {
            var codes = new List<string>();
            if (string.IsNullOrWhiteSpace(raw))
            {
                return codes;
            }

            foreach (string part in raw.Split(','))
            {
                string code = part.Trim();
                if (code.Length > 0)
                {
                    codes.Add(code);
                }
            }

            return codes;
        }

        private static bool ParagraphCodes(List<string> codes)
        {
            if (codes == null || codes.Count == 0)
            {
                return false;
            }

            foreach (string code in codes)
            {
                if (!FormatContextHelper.IsParagraphCode(code))
                {
                    return false;
                }
            }

            return true;
        }

        private static bool Intersects(Word.Range range, Word.Paragraph paragraph)
        {
            if (range == null || paragraph == null || paragraph.Range == null)
            {
                return false;
            }

            return paragraph.Range.Start < range.End && paragraph.Range.End > range.Start;
        }

        private static bool CoversWhole(Word.Range range, Word.Paragraph paragraph)
        {
            int paraStart = paragraph.Range.Start;
            int paraEnd = paragraph.Range.End;
            int contentEnd = paraEnd > paraStart ? paraEnd - 1 : paraStart;
            return range.Start <= paraStart && range.End >= contentEnd;
        }

        private static Word.Range Slice(Word.Document doc, Word.Range range, Word.Paragraph paragraph)
        {
            int paraStart = paragraph.Range.Start;
            int paraEnd = paragraph.Range.End;
            int contentEnd = paraEnd > paraStart ? paraEnd - 1 : paraStart;
            int start = Math.Max(range.Start, paraStart);
            int end = Math.Min(range.End, contentEnd);
            if (end <= start)
            {
                return null;
            }

            return doc.Range(start, end);
        }

        private static void ApplyParagraphSide(
            Word.Document doc,
            StructureStyleRecord style,
            Word.Paragraph paragraph)
        {
            paragraph.Format.Reset();
            SetStyle(doc, paragraph.Range, Names(style, false), false);
        }

        private static void ApplyWhole(Word.Document doc, StructureStyleRecord style, Word.Paragraph paragraph)
        {
            if (style.Support == "both")
            {
                ResetFont(ParagraphText(paragraph));
                paragraph.Format.Reset();
                SetStyle(doc, paragraph.Range, Names(style, false), false);
            }
            else if (style.Support == "paragraph")
            {
                paragraph.Format.Reset();
                SetStyle(doc, paragraph.Range, Names(style, false), false);
            }
            else
            {
                Word.Range target = ParagraphText(paragraph);
                target.Font.Reset();
                SetStyle(doc, target, Names(style, true), true);
            }
        }

        private static void ApplySlice(Word.Document doc, StructureStyleRecord style, Word.Range slice)
        {
            if (style.Support == "paragraph")
            {
                return;
            }

            if (style.Support == "both" && style.Structure == "body")
            {
                slice.Font.Reset();
                return;
            }

            if (style.Support == "both")
            {
                slice.Font.Reset();
                if (string.IsNullOrEmpty(style.CharacterStyleName))
                {
                    WriteRunText(slice, style.TextFormat);
                }
                else
                {
                    SetStyle(doc, slice, Names(style, true), true);
                }

                return;
            }

            slice.Font.Reset();
            SetStyle(doc, slice, Names(style, true), true);
        }

        private static bool HostHasStyle(Word.Document doc, string structure)
        {
            if (doc == null || string.IsNullOrWhiteSpace(structure))
            {
                return false;
            }

            var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (string name in HostNames(structure.Trim()))
            {
                foreach (string equivalent in StructureStyleCatalog.EquivalentNames(name))
                {
                    if (!string.IsNullOrEmpty(equivalent))
                    {
                        names.Add(equivalent);
                    }
                }
            }

            foreach (Word.Style style in doc.Styles)
            {
                try
                {
                    string local = style.NameLocal;
                    if (!string.IsNullOrEmpty(local) && names.Contains(local))
                    {
                        return true;
                    }
                }
                catch (Exception)
                {
                }
            }

            return false;
        }

        private static IEnumerable<string> HostNames(string structure)
        {
            yield return structure;
            if (structure.Length == 9
                && structure.StartsWith("heading_", StringComparison.OrdinalIgnoreCase)
                && structure[8] >= '1'
                && structure[8] <= '9')
            {
                string level = structure.Substring(8);
                yield return "标题 " + level;
                yield return "heading " + level;
                yield break;
            }

            if (string.Equals(structure, "body", StringComparison.OrdinalIgnoreCase))
            {
                yield return "正文";
                yield return "Normal";
            }
            else if (string.Equals(structure, "title", StringComparison.OrdinalIgnoreCase))
            {
                yield return "标题";
                yield return "Title";
            }
            else if (string.Equals(structure, "subtitle", StringComparison.OrdinalIgnoreCase))
            {
                yield return "副标题";
                yield return "Subtitle";
            }
        }

        private const string CharacterApplyNote =
            "段落格式没有发生作用。这次只盖住了段里的一截，没有整段。要想段落格式生效，需要套整段。";

        private const string ParagraphSkippedNote =
            "段落格式没有发生作用。这次只盖住了段里的一截，没有整段。要想段落格式生效，需要套整段。没有字符格式可套，所以这次套用什么也不会发生。";

        private const string MixedParagraphNote =
            "段落格式只在盖住完整的段落上生效。半截不生效。要想半截的段落格式生效，需要套整段。";

        private static Dictionary<string, object> Done(
            string code,
            StructureStyleRecord style,
            string scope,
            List<string> warnings,
            string note)
        {
            var done = new Dictionary<string, object>
            {
                ["ok"] = true,
                ["structure"] = style.Structure,
                ["scope"] = scope
            };
            if (!string.IsNullOrEmpty(code) && code.Trim().StartsWith("P_", StringComparison.Ordinal))
            {
                done["target_paragraph_codes"] = code.Trim();
            }
            else
            {
                done["target_codes"] = code.Trim();
            }
            if (!string.IsNullOrEmpty(note))
            {
                done["note"] = note;
            }

            if (warnings.Count > 0)
            {
                done["format_warnings"] = warnings;
            }

            return done;
        }

        private static Word.Range ParagraphText(Word.Paragraph paragraph)
        {
            Word.Range range = paragraph.Range.Duplicate;
            if (range.End > range.Start)
            {
                range.MoveEnd(Word.WdUnits.wdCharacter, -1);
            }

            return range;
        }

        private static void ResetFont(Word.Range range)
        {
            range.Font.Reset();
        }

        private static void WriteRunText(Word.Range range, Dictionary<string, object> text)
        {
            if (range == null || text == null || text.Count == 0)
            {
                return;
            }

            Word.Font font = range.Font;
            object face;
            if (text.TryGetValue("font_name", out face) && face != null)
            {
                font.Name = face.ToString();
                font.NameFarEast = face.ToString();
            }

            object size;
            if (text.TryGetValue("font_size", out size) && size != null)
            {
                font.Size = Convert.ToSingle(size, System.Globalization.CultureInfo.InvariantCulture);
            }

            object color;
            if (text.TryGetValue("font_color", out color) && color != null)
            {
                font.Color = (Word.WdColor)RgbToBgr(color.ToString());
            }

            object bold;
            if (text.TryGetValue("is_bold", out bold))
            {
                font.Bold = True(bold) ? -1 : 0;
            }

            object italic;
            if (text.TryGetValue("is_italic", out italic))
            {
                font.Italic = True(italic) ? -1 : 0;
            }

            object underline;
            if (text.TryGetValue("has_underline", out underline))
            {
                font.Underline = True(underline)
                    ? Word.WdUnderline.wdUnderlineSingle
                    : Word.WdUnderline.wdUnderlineNone;
            }
        }

        private static bool True(object value)
        {
            if (value is bool bit)
            {
                return bit;
            }

            string text = value == null ? "" : value.ToString();
            return text == "1" || text.Equals("true", StringComparison.OrdinalIgnoreCase);
        }

        private static int RgbToBgr(string raw)
        {
            string hex = raw.Trim().TrimStart('#');
            int rgb = int.Parse(hex, System.Globalization.NumberStyles.HexNumber, System.Globalization.CultureInfo.InvariantCulture);
            return ((rgb & 0xFF) << 16) | (rgb & 0xFF00) | ((rgb >> 16) & 0xFF);
        }

        private static void SetStyle(Word.Document doc, Word.Range range, IList<string> names, bool character)
        {
            Exception last = null;
            foreach (string name in names)
            {
                if (string.IsNullOrWhiteSpace(name))
                {
                    continue;
                }

                Word.Style found = FindStyle(doc, name);
                if (found == null)
                {
                    continue;
                }

                bool isCharacter = found.Type == Word.WdStyleType.wdStyleTypeCharacter;
                if (isCharacter != character)
                {
                    continue;
                }

                try
                {
                    object styleRef = found;
                    range.set_Style(ref styleRef);
                    return;
                }
                catch (Exception ex)
                {
                    last = ex;
                }
            }

            throw last ?? new InvalidOperationException("文档里找不到对应的样式。");
        }

        private static Word.Style FindStyle(Word.Document doc, string name)
        {
            try
            {
                return doc.Styles[name];
            }
            catch (Exception)
            {
                return null;
            }
        }

        private static IList<string> Names(StructureStyleRecord style, bool character)
        {
            var names = new List<string>();
            if (character)
            {
                Add(names, style.CharacterStyleName);
                Add(names, style.StyleName);
            }
            else
            {
                Add(names, style.StyleName);
            }

            foreach (string name in StructureStyleCatalog.LookupNames(style))
            {
                Add(names, name);
            }

            return names;
        }

        private static void Add(List<string> names, string name)
        {
            if (!string.IsNullOrWhiteSpace(name) && !names.Contains(name))
            {
                names.Add(name);
            }
        }

        private static string Arg(Dictionary<string, object> item, string key)
        {
            if (item == null || !item.ContainsKey(key) || item[key] == null)
            {
                return null;
            }

            return item[key].ToString().Trim();
        }

        private static Dictionary<string, object> Fail(string error)
        {
            return new Dictionary<string, object> { ["ok"] = false, ["error"] = error };
        }
    }
}
