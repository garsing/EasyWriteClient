using System;
using System.Collections.Generic;
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
            string codeError = ReadCodes(item, out List<string> codes, out bool paragraph, out TextSpan span);
            if (codeError != null)
            {
                return Fail(codeError);
            }

            string prepared = Prepare(doc, disallowBackendApi);
            if (prepared != null)
            {
                return Fail(prepared);
            }

            string snapshot = paragraph
                ? DocumentState.ParagraphSegSnapshot
                : DocumentState.SegSnapshot;
            List<List<string>> runs = FormatApplyBatchCompiler.SplitRuns(codes, snapshot);
            if (runs.Count == 0)
            {
                runs.Add(codes);
            }

            string check = Validate(doc, runs, paragraph);
            if (check != null)
            {
                return Fail(check);
            }

            display = F_DocumentStructureTool.ReadLocateDisplay();
            if (string.IsNullOrEmpty(display))
            {
                return Fail("没有定位用的句子 display。请先 F_get_document_content。");
            }

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

            return Unlink(doc, Choose(display, body, styles, windows), styles);
        }

        public static Dictionary<string, object> ClearAll(
            Word.Document doc,
            IList<StructureStyleRecord> styles,
            string display,
            XElement body,
            bool disallowBackendApi)
        {
            string prepared = Prepare(doc, disallowBackendApi);
            if (prepared != null)
            {
                return Fail(prepared);
            }

            display = F_DocumentStructureTool.ReadLocateDisplay();
            if (string.IsNullOrEmpty(display))
            {
                return Fail("没有定位用的句子 display。请先 F_get_document_content。");
            }

            return Unlink(doc, Choose(display, body, styles, null), styles);
        }

        private static List<StructureSpanLister.StyledSpan> Choose(
            string display,
            XElement body,
            IList<StructureStyleRecord> styles,
            List<int[]> windows)
        {
            var chosen = new List<StructureSpanLister.StyledSpan>();
            StructureSpanLister.LocateStyledResult located = StructureSpanLister.LocateStyled(display, body, styles);
            List<int[]> interiors = StructureSpanLister.TocInteriors(display);
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

                if (!StructureSpanLister.TryBindClear(display, piece, out StructureSpanLister.StyledSpan span))
                {
                    continue;
                }

                chosen.Add(span);
            }

            return chosen;
        }

        private static Dictionary<string, object> Unlink(
            Word.Document doc,
            List<StructureSpanLister.StyledSpan> chosen,
            IList<StructureStyleRecord> styles)
        {
            var ranges = new List<Word.Range>();
            var links = new List<bool>();
            foreach (StructureSpanLister.StyledSpan hit in chosen)
            {
                if (!TryRange(doc, hit, out Word.Range range, out string error))
                {
                    return Fail(error);
                }

                ranges.Add(range);
                links.Add(hit.ParagraphLink);
            }

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
                return Fail(ex.Message);
            }

            return new Dictionary<string, object>
            {
                ["ok"] = true,
                ["cleared"] = ReadBack(ranges)
            };
        }

        private static bool TryRange(
            Word.Document doc,
            StructureSpanLister.StyledSpan hit,
            out Word.Range range,
            out string error)
        {
            range = null;
            error = null;
            var options = new DisplaySequenceLocateOptions
            {
                Snapshot = DocumentState.Snapshot,
                SegSnapshot = DocumentState.SegSnapshot
            };
            DisplaySequenceLocateResult located = DisplaySequenceLocator.TryResolve(
                new List<string> { hit.Code },
                options);
            if (located.MatchCount != 1 || !located.DisplayStartPosition.HasValue)
            {
                error = "无法定位 " + hit.Code + "。";
                return false;
            }

            range = DisplayPositionRangeResolver.ResolveSpan(
                doc,
                new List<string> { hit.Code },
                located.DisplayStartPosition.Value,
                null,
                null,
                null,
                "clear");
            if (range == null)
            {
                error = "无法定位 " + hit.Code + "。";
                return false;
            }

            if (string.IsNullOrEmpty(hit.TextValue))
            {
                return true;
            }

            var text = new TextSpan
            {
                Value = hit.TextValue,
                HasIndex = hit.HasIndex,
                IndexIsInteger = true,
                Index = hit.TextIndex
            };
            InsertTargetLocatorHelper.ResolveRangeResult narrowed = TextSpanNarrower.TryNarrow(
                doc,
                range,
                text,
                hit.Code);
            if (narrowed == null || !narrowed.Success || narrowed.Range == null)
            {
                error = narrowed == null ? "无法定位 " + hit.Code + "。" : narrowed.Error;
                return false;
            }

            range = narrowed.Range;
            return true;
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
            foreach (string one in codes)
            {
                if (FormatContextHelper.IsParagraphCode(one))
                {
                    anyP = true;
                }
                else if (FormatContextHelper.IsSentenceCode(one))
                {
                    anyS = true;
                }
            }

            if (anyP && anyS)
            {
                return "同一项混用 S_ 与 P_，未套用";
            }

            if (!anyP && !anyS)
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
