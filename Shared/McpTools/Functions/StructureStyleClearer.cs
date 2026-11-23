using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Xml.Linq;
using WordAddIn1.DocumentMapping.CodeResolve;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    /// <summary>断开套用。在没收起的句子 display 上决定断哪一截，再查找 Range。</summary>
    public static class StructureStyleClearer
    {
        public static Dictionary<string, object> ClearItem(
            Word.Document doc,
            Dictionary<string, object> item,
            IList<StructureStyleRecord> styles,
            string display,
            XElement body,
            bool disallowBackendApi)
        {
            // 按定位圈窗口先不能用：太复杂。
            return Fail("按定位清样式先不能用。圈窗口太复杂。请用全文清。");
#if false
            Stopwatch total = Stopwatch.StartNew();
            Stopwatch sw = Stopwatch.StartNew();
            string codeError = ReadCodes(item, out List<string> codes, out bool paragraph, out TextSpan span);
            if (codeError != null)
            {
                return Fail(codeError);
            }

            string prepared = Prepare(doc, disallowBackendApi);
            EasyWriteDiagnostics.LogTiming(
                "style_clear.prepare",
                sw.ElapsedMilliseconds,
                codes == null ? "" : "codes=" + string.Join(",", codes));
            if (prepared != null)
            {
                return Fail(prepared);
            }

            sw.Restart();
            string snapshot = paragraph
                ? DocumentState.ParagraphSegSnapshot
                : DocumentState.SegSnapshot;
            List<List<string>> runs = FormatApplyBatchCompiler.SplitRuns(codes, snapshot);
            if (runs.Count == 0)
            {
                runs.Add(codes);
            }

            string check = Validate(doc, runs, paragraph);
            EasyWriteDiagnostics.LogTiming(
                "style_clear.validate",
                sw.ElapsedMilliseconds,
                "runs=" + runs.Count);
            if (check != null)
            {
                return Fail(check);
            }

            sw.Restart();
            display = F_DocumentStructureTool.ReadLocateDisplay();
            EasyWriteDiagnostics.LogTiming(
                "style_clear.read_display",
                sw.ElapsedMilliseconds,
                "len=" + (display == null ? 0 : display.Length));
            if (string.IsNullOrEmpty(display))
            {
                return Fail("没有定位用的句子 display。请先 F_get_document_content。");
            }

            sw.Restart();
            var windows = new List<int[]>();
            foreach (List<string> run in runs)
            {
                if (!StructureSpanLister.TryClearWindow(
                    display,
                    run,
                    paragraph,
                    span,
                    out int start,
                    out int end,
                    out string windowError))
                {
                    return Fail(windowError);
                }

                windows.Add(new[] { start, end });
            }

            EasyWriteDiagnostics.LogTiming(
                "style_clear.windows",
                sw.ElapsedMilliseconds,
                "n=" + windows.Count);

            sw.Restart();
            List<StructureSpanLister.StyledSpan> chosen = Choose(display, body, styles, windows);
            EasyWriteDiagnostics.LogTiming(
                "style_clear.choose",
                sw.ElapsedMilliseconds,
                "n=" + chosen.Count);
            Dictionary<string, object> result = Unlink(doc, chosen, styles, disallowBackendApi);
            EasyWriteDiagnostics.LogTiming("style_clear.item_total", total.ElapsedMilliseconds);
            return result;
#endif
        }

        public static Dictionary<string, object> ClearAll(
            Word.Document doc,
            IList<StructureStyleRecord> styles,
            string display,
            XElement body,
            bool disallowBackendApi)
        {
            Stopwatch total = Stopwatch.StartNew();
            Stopwatch sw = Stopwatch.StartNew();
            string prepared = Prepare(doc, disallowBackendApi);
            EasyWriteDiagnostics.LogTiming("style_clear.prepare", sw.ElapsedMilliseconds, "all");
            if (prepared != null)
            {
                return Fail(prepared);
            }

            sw.Restart();
            display = F_DocumentStructureTool.ReadLocateDisplay();
            EasyWriteDiagnostics.LogTiming(
                "style_clear.read_display",
                sw.ElapsedMilliseconds,
                "len=" + (display == null ? 0 : display.Length));
            if (string.IsNullOrEmpty(display))
            {
                return Fail("没有定位用的句子 display。请先 F_get_document_content。");
            }

            sw.Restart();
            List<StructureSpanLister.StyledSpan> chosen = Choose(display, body, styles, null);
            EasyWriteDiagnostics.LogTiming(
                "style_clear.choose",
                sw.ElapsedMilliseconds,
                "n=" + chosen.Count);
            Dictionary<string, object> result = Unlink(doc, chosen, styles, disallowBackendApi);
            EasyWriteDiagnostics.LogTiming("style_clear.item_total", total.ElapsedMilliseconds, "all");
            return result;
        }

        private static List<StructureSpanLister.StyledSpan> Choose(
            string display,
            XElement body,
            IList<StructureStyleRecord> styles,
            List<int[]> windows)
        {
            var chosen = new List<StructureSpanLister.StyledSpan>();
            Stopwatch sw = Stopwatch.StartNew();
            StructureSpanLister.LocateStyledResult located = StructureSpanLister.LocateStyled(display, body, styles);
            EasyWriteDiagnostics.LogTiming(
                "style_clear.locate_styled",
                sw.ElapsedMilliseconds,
                "spans=" + located.Spans.Count + " miss=" + located.Misses.Count);
            sw.Restart();
            List<int[]> interiors = StructureSpanLister.TocInteriors(display);
            int nextGroup = 1;
            foreach (StructureSpanLister.LocatedSpan piece in located.Spans)
            {
                if (piece.Style != null && piece.Style.Structure == "body")
                {
                    continue;
                }

                if (StructureSpanLister.InsideToc(piece.Start, piece.End, interiors))
                {
                    continue;
                }

                if (windows != null && !InsideAny(piece.Start, piece.End, windows))
                {
                    continue;
                }

                List<StructureSpanLister.StyledSpan> parts = StructureSpanLister.BindPiece(display, piece);
                if (parts.Count == 0)
                {
                    continue;
                }

                if (parts.Count > 1)
                {
                    int group = nextGroup++;
                    foreach (StructureSpanLister.StyledSpan part in parts)
                    {
                        part.Group = group;
                    }
                }

                chosen.AddRange(parts);
            }

            EasyWriteDiagnostics.LogTiming(
                "style_clear.filter_bind",
                sw.ElapsedMilliseconds,
                "chosen=" + chosen.Count + " toc=" + interiors.Count);
            return chosen;
        }

        private static Dictionary<string, object> Unlink(
            Word.Document doc,
            List<StructureSpanLister.StyledSpan> chosen,
            IList<StructureStyleRecord> styles,
            bool disallowBackendApi)
        {
            var ranges = new List<Word.Range>();
            var links = new List<bool>();
            Stopwatch sw = Stopwatch.StartNew();
            if (chosen != null && chosen.Count > 0)
            {
                var locators = new List<Dictionary<string, object>>();
                foreach (StructureSpanLister.StyledSpan hit in chosen)
                {
                    locators.Add(LocatorFor(hit));
                    links.Add(hit.ParagraphLink);
                }

                InsertTargetLocatorHelper.LocateSpanBatchResult batch =
                    InsertTargetLocatorHelper.TryLocateSpans(doc, locators, disallowBackendApi);
                if (batch == null || batch.BatchBlocked || batch.Items == null || batch.Items.Count != chosen.Count)
                {
                    EasyWriteDiagnostics.LogTiming(
                        "style_clear.find_range",
                        sw.ElapsedMilliseconds,
                        "fail batch");
                    return Fail(batch == null || string.IsNullOrEmpty(batch.Error) ? "定位失败。" : batch.Error);
                }

                for (int i = 0; i < batch.Items.Count; i++)
                {
                    InsertTargetLocatorHelper.LocateSpanItem located = batch.Items[i];
                    if (located == null || !located.Success || located.Range == null)
                    {
                        EasyWriteDiagnostics.LogTiming(
                            "style_clear.find_range",
                            sw.ElapsedMilliseconds,
                            "ok=" + ranges.Count + " fail=" + (chosen[i].Code ?? ""));
                        return Fail(located == null || string.IsNullOrEmpty(located.Error) ? "定位失败。" : located.Error);
                    }

                    ranges.Add(located.Range);
                }
            }

            Assemble(doc, chosen, ranges, links);
            EasyWriteDiagnostics.LogTiming(
                "style_clear.find_range",
                sw.ElapsedMilliseconds,
                "n=" + ranges.Count);

            sw.Restart();
            try
            {
                for (int i = 0; i < ranges.Count; i++)
                {
                    if (links[i])
                    {
                        ClearParagraph(doc, ranges[i], styles);
                    }
                    else
                    {
                        ClearCharacter(doc, ranges[i]);
                    }
                }
            }
            catch (Exception ex)
            {
                EasyWriteDiagnostics.LogTiming(
                    "style_clear.set_style",
                    sw.ElapsedMilliseconds,
                    "fail n=" + ranges.Count);
                return Fail(ex.Message);
            }

            EasyWriteDiagnostics.LogTiming(
                "style_clear.set_style",
                sw.ElapsedMilliseconds,
                "n=" + ranges.Count);

            sw.Restart();
            List<Dictionary<string, object>> cleared = ReadBack(ranges);
            EasyWriteDiagnostics.LogTiming(
                "style_clear.read_back",
                sw.ElapsedMilliseconds,
                "runs=" + (cleared == null ? 0 : cleared.Count));
            return new Dictionary<string, object>
            {
                ["ok"] = true,
                ["cleared"] = cleared
            };
        }

        private static void Assemble(
            Word.Document doc,
            List<StructureSpanLister.StyledSpan> chosen,
            List<Word.Range> ranges,
            List<bool> links)
        {
            if (chosen == null || ranges == null || ranges.Count == 0)
            {
                return;
            }

            bool anyGroup = false;
            for (int i = 0; i < chosen.Count && i < ranges.Count; i++)
            {
                if (chosen[i].Group != 0)
                {
                    anyGroup = true;
                    break;
                }
            }

            if (!anyGroup)
            {
                return;
            }

            var assembled = new List<Word.Range>();
            var assembledLinks = new List<bool>();
            int index = 0;
            while (index < ranges.Count)
            {
                int group = index < chosen.Count ? chosen[index].Group : 0;
                if (group == 0)
                {
                    assembled.Add(ranges[index]);
                    assembledLinks.Add(index < links.Count && links[index]);
                    index++;
                    continue;
                }

                int start = ranges[index].Start;
                int end = ranges[index].End;
                bool link = index < links.Count && links[index];
                while (index < ranges.Count && index < chosen.Count && chosen[index].Group == group)
                {
                    if (ranges[index].Start < start)
                    {
                        start = ranges[index].Start;
                    }

                    if (ranges[index].End > end)
                    {
                        end = ranges[index].End;
                    }

                    index++;
                }

                assembled.Add(doc.Range(start, end));
                assembledLinks.Add(link);
            }

            ranges.Clear();
            ranges.AddRange(assembled);
            links.Clear();
            links.AddRange(assembledLinks);
        }

        private static Dictionary<string, object> LocatorFor(StructureSpanLister.StyledSpan hit)
        {
            var locator = new Dictionary<string, object>();
            bool paragraph = FormatContextHelper.IsParagraphCode(hit.Code);
            locator[paragraph ? "paragraph_codes" : "codes"] = hit.Code ?? "";
            if (!string.IsNullOrEmpty(hit.TextValue))
            {
                if (hit.HasIndex)
                {
                    locator["text"] = new Dictionary<string, object>
                    {
                        ["value"] = hit.TextValue,
                        ["index"] = hit.TextIndex
                    };
                }
                else
                {
                    locator["text"] = hit.TextValue;
                }
            }

            return locator;
        }

        private static void ClearCharacter(Word.Document doc, Word.Range range)
        {
            Word.Style font = doc.Styles[Word.WdBuiltinStyle.wdStyleDefaultParagraphFont];
            object styleRef = font;
            range.set_Style(ref styleRef);
        }

        private static void ClearParagraph(
            Word.Document doc,
            Word.Range range,
            IList<StructureStyleRecord> styles)
        {
            StructureStyleRecord body = null;
            if (styles != null)
            {
                foreach (StructureStyleRecord row in styles)
                {
                    if (row.Structure == "body")
                    {
                        body = row;
                        break;
                    }
                }
            }

            var names = new List<string> { "正文", "Normal" };
            if (body != null)
            {
                foreach (string name in StructureStyleCatalog.LookupNames(body))
                {
                    if (!names.Contains(name))
                    {
                        names.Add(name);
                    }
                }
            }

            foreach (Word.Paragraph paragraph in range.Paragraphs)
            {
                if (paragraph == null || paragraph.Range == null)
                {
                    continue;
                }

                if (paragraph.Range.Start >= range.End || paragraph.Range.End <= range.Start)
                {
                    continue;
                }

                if (AlreadyBody(paragraph))
                {
                    continue;
                }

                Word.Style found = null;
                foreach (string name in names)
                {
                    try
                    {
                        found = doc.Styles[name];
                    }
                    catch (Exception)
                    {
                        found = null;
                    }

                    if (found != null && found.Type != Word.WdStyleType.wdStyleTypeCharacter)
                    {
                        break;
                    }

                    found = null;
                }

                if (found == null)
                {
                    throw new InvalidOperationException("文档里找不到正文样式。");
                }

                object styleRef = found;
                paragraph.set_Style(ref styleRef);
            }
        }

        private static bool AlreadyBody(Word.Paragraph paragraph)
        {
            try
            {
                object raw = paragraph.get_Style();
                Word.Style style = raw as Word.Style;
                string name = style == null ? "" : (style.NameLocal ?? "");
                return name == "正文" || name == "Normal";
            }
            catch (Exception)
            {
                return false;
            }
        }

        private static List<Dictionary<string, object>> ReadBack(List<Word.Range> ranges)
        {
            var runs = new List<Dictionary<string, object>>();
            foreach (Word.Range range in ranges)
            {
                if (!CharFormatRunReader.TryRead(range, out List<Dictionary<string, object>> read, out string _))
                {
                    continue;
                }

                if (read == null)
                {
                    continue;
                }

                foreach (Dictionary<string, object> run in read)
                {
                    runs.Add(run);
                }
            }

            var merged = new List<Dictionary<string, object>>();
            foreach (Dictionary<string, object> run in runs)
            {
                if (merged.Count == 0)
                {
                    merged.Add(run);
                    continue;
                }

                Dictionary<string, object> previous = merged[merged.Count - 1];
                object left = previous.ContainsKey("char_format") ? previous["char_format"] : null;
                object right = run.ContainsKey("char_format") ? run["char_format"] : null;
                if (left is Dictionary<string, object> leftFormat
                    && right is Dictionary<string, object> rightFormat
                    && CharFormatCatalog.SameSnapshot(leftFormat, rightFormat))
                {
                    previous["text"] = (previous.ContainsKey("text") ? previous["text"]?.ToString() : "")
                        + (run.ContainsKey("text") ? run["text"]?.ToString() : "");
                    continue;
                }

                merged.Add(run);
            }

            return merged;
        }

        private static bool InsideAny(int start, int end, List<int[]> windows)
        {
            foreach (int[] window in windows)
            {
                if (start >= window[0] && end <= window[1])
                {
                    return true;
                }
            }

            return false;
        }

        private static string Validate(Word.Document doc, List<List<string>> runs, bool paragraph)
        {
            foreach (List<string> run in runs)
            {
                if (run != null && run.Count == 1 && EmptyCellCode.IsEmptyCellCode(run[0]))
                {
                    continue;
                }

                if (paragraph)
                {
                    AmbiguityCheckResult ambiguity = ParagraphFormatAmbiguityHelper.RunPrecheck(
                        new ApplyFormatScopeArgs(),
                        run,
                        "clear",
                        out TableScopeIndex _,
                        out string cacheError);
                    if (cacheError != null)
                    {
                        return cacheError;
                    }

                    if (ambiguity != null && ambiguity.HasBlockingError)
                    {
                        return ambiguity.SummaryMessage ?? "定位歧义";
                    }
                }
                else
                {
                    var anchor = new DocumentActionAmbiguityAnchor
                    {
                        ActionIndex = 0,
                        Field = "target",
                        Codes = string.Join(",", run)
                    };
                    AmbiguityCheckResult ambiguity = SentenceCodeAmbiguityHelper.CheckBatch(
                        new[] { anchor },
                        DocumentState.Snapshot,
                        ConfigManager.GetDocumentActionsAmbiguitySettings(),
                        null);
                    if (ambiguity != null && ambiguity.HasBlockingError)
                    {
                        return ambiguity.SummaryMessage ?? "定位歧义";
                    }
                }
            }

            return null;
        }

        private static string Prepare(Word.Document doc, bool disallowBackendApi)
        {
            return InsertTargetLocatorHelper.TryProcessDocumentForLocate(doc, disallowBackendApi);
        }

        private static string ReadCodes(
            Dictionary<string, object> item,
            out List<string> codes,
            out bool paragraph,
            out TextSpan span)
        {
            codes = SplitCodes(Arg(item, "target_codes"));
            paragraph = false;
            span = null;
            if (codes.Count == 0)
            {
                return "缺少 target_codes。";
            }

            string sibling = TextSpanArg.RejectSiblingIndex(item);
            if (sibling != null)
            {
                return sibling;
            }

            bool narrowing = TextSpanArg.HasNarrowingText(item);
            if (narrowing && codes.Count != 1)
            {
                return TextSpanArg.SingleCodeError;
            }

            if (item != null && item.ContainsKey("text"))
            {
                if (!TextSpanArg.TryRead(item, out span, out string parseError))
                {
                    return parseError ?? "text 无法解析。";
                }
            }

            bool anyP = false;
            bool anyS = false;
            bool anyT = false;
            foreach (string one in codes)
            {
                if (EmptyCellCode.IsEmptyCellCode(one))
                {
                    anyT = true;
                }
                else if (FormatContextHelper.IsParagraphCode(one))
                {
                    anyP = true;
                }
                else if (FormatContextHelper.IsSentenceCode(one))
                {
                    anyS = true;
                }
            }

            if (anyT && (anyP || anyS))
            {
                return "同一项不能把几何码和 S_、P_ 写在一起。";
            }

            if (anyT && codes.Count != 1)
            {
                return "几何码一次一条。";
            }

            if (anyP && anyS)
            {
                return "同一项混用 S_ 与 P_，未套用";
            }

            if (!anyP && !anyS && !anyT)
            {
                return "locate_failed: 无法定位编码序列";
            }

            paragraph = anyP;
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

        private static string Arg(Dictionary<string, object> args, string key)
        {
            if (args == null || !args.TryGetValue(key, out object value) || value == null)
            {
                return "";
            }

            return value.ToString().Trim();
        }

        private static Dictionary<string, object> Fail(string error)
        {
            return new Dictionary<string, object>
            {
                ["ok"] = false,
                ["error"] = error ?? "清除失败。"
            };
        }

    }
}
