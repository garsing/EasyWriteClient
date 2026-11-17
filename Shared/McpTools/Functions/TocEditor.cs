using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    internal static class TocFormatTrace
    {
        private static readonly object Gate = new object();

        internal static void Write(string message)
        {
            string line = DateTime.Now.ToString("HH:mm:ss.fff ") + (message ?? "");
            System.Diagnostics.Debug.WriteLine("[toc-format] " + line);
            try
            {
                lock (Gate)
                {
                    System.IO.File.AppendAllText(
                        System.IO.Path.Combine(System.IO.Path.GetTempPath(), "ew-toc-format.log"),
                        line + Environment.NewLine);
                }
            }
            catch (Exception)
            {
            }
        }
    }

    /// <summary>
    /// 一份目录的域指令、结果文字和这一份结果上的直接格式。
    /// 目录 1–9 和目录标题仍是文档共用的内置样式。标题按 wdStyleTocHeading 取，本地名可能是「目录标题」「TOC 标题」或 TOC Heading。
    /// </summary>
    public static class TocEditor
    {
        private const string DefaultLevels = "1-3";
        private static readonly Regex LevelsPattern = new Regex(
            @"^([1-9])-([1-9])$",
            RegexOptions.Compiled | RegexOptions.CultureInvariant);
        private static readonly Regex ParagraphCode = new Regex(
            @"^P_[A-Za-z0-9]+$",
            RegexOptions.Compiled | RegexOptions.CultureInvariant);

        public static ToolResult Run(Word.Document doc, string action, Dictionary<string, object> args)
        {
            if (doc == null)
            {
                return Fail("没有打开的文档");
            }

            string xml = F_TocTool.Arg(args, "xml");
            TocSpec spec = null;
            if (xml.Trim().Length > 0 && !TocXml.TryParse(xml, out spec, out string parseError))
            {
                return Fail(parseError ?? "目录 XML 无法解析");
            }

            if (action == "insert")
            {
                return Insert(doc, F_TocTool.Arg(args, "at"), spec ?? TocSpec.InsertDefault());
            }

            string tocId = F_TocTool.Arg(args, "toc_id").Trim();
            if (!TryLocate(doc, tocId, out TocHit hit, out string locateError))
            {
                return Fail(locateError ?? "这份目录已经不在");
            }

            if (action == "get")
            {
                return Get(doc, hit);
            }

            if (action == "delete")
            {
                return Delete(doc, hit);
            }

            if (action == "reset")
            {
                return Reset(doc, hit);
            }

            return Update(doc, hit, spec);
        }

        private static ToolResult Insert(Word.Document doc, string at, TocSpec spec)
        {
            if (!TryInsertPoint(doc, at, out Word.Range point, out string error))
            {
                return Fail(error);
            }

            string bookmark = null;
            if (spec.HasRange)
            {
                if (!TryCreateRangeBookmark(doc, spec.From, spec.To, out bookmark, out error))
                {
                    return Fail(error);
                }
            }

            spec.Bookmark = bookmark;
            Word.Field field = AddField(doc, point, TocXml.BuildInstruction(spec));
            if (field == null)
            {
                return Fail("插入目录域失败");
            }

            UpdateField(field);
            string id = AddBookmarkAtCode(doc, field, null);
            if (spec.HasTitle && !string.IsNullOrEmpty(spec.Title))
            {
                InsertTitleBefore(doc, id, spec.Title);
                field = FindField(doc, id) ?? field;
            }

            List<Dictionary<string, object>> warnings = ApplyOverlays(doc, field, spec, reset: false);
            return WithWarnings(Ok(id, "已插入目录 " + id, TocXml.Write(spec, id, includeDefaults: false), LinesOf(doc, field, false)), warnings);
        }

        private static ToolResult Get(Word.Document doc, TocHit hit)
        {
            if (hit.Field == null)
            {
                var placeholder = TocSpec.InsertDefault();
                var lines = new List<Dictionary<string, object>>
                {
                    new Dictionary<string, object>
                    {
                        ["level"] = "",
                        ["text"] = PlaceholderText(hit),
                        ["page"] = "",
                        ["placeholder"] = true
                    }
                };
                return Ok(hit.Id, "目录 " + hit.Id + " 还是占位，没有生成条目", TocXml.Write(placeholder, hit.Id, includeDefaults: false), lines, empty: true);
            }

            if (!TocXml.TryReadInstruction(hit.Field.Code.Text ?? "", out TocSpec spec, out string error))
            {
                return Fail(error ?? "读不到这份目录的指令");
            }

            if (!string.IsNullOrEmpty(spec.Bookmark))
            {
                if (!TryBookmarkToCodes(doc, spec.Bookmark, out string from, out string to, out error))
                {
                    return Fail(error ?? "范围已经不在");
                }

                spec.From = from;
                spec.To = to;
                spec.HasRange = true;
            }

            ReadDirect(doc, hit.Field, spec);
            string title = TitleText(doc, hit.Field);
            if (!string.IsNullOrWhiteSpace(title))
            {
                spec.HasTitle = true;
                spec.Title = title;
            }

            List<Dictionary<string, object>> body = LinesOf(doc, hit.Field, false);
            bool empty = body.Count == 0;
            return Ok(
                hit.Id,
                empty ? "目录 " + hit.Id + " 没有条目" : "目录 " + hit.Id,
                TocXml.Write(spec, hit.Id, includeDefaults: false),
                body,
                empty);
        }

        private static ToolResult Update(Word.Document doc, TocHit hit, TocSpec patch)
        {
            if (patch == null)
            {
                return Fail("更新目录须提供 xml");
            }

            if (hit.Field == null)
            {
                return ReplacePlaceholder(doc, hit, patch);
            }

            if (!TocXml.TryReadInstruction(hit.Field.Code.Text ?? "", out TocSpec current, out string error))
            {
                return Fail(error ?? "读不到这份目录的指令");
            }

            current.Bookmark = ExtractSwitch(hit.Field.Code.Text ?? "", "b");
            ReadDirect(doc, hit.Field, current);
            string title = TitleText(doc, hit.Field);
            current.HasTitle = title != null;
            current.Title = title ?? "";
            bool instructionChanged = TocXml.Merge(current, patch, out error);
            if (error != null)
            {
                return Fail(error);
            }

            if (patch.Reset)
            {
                current.Reset = true;
                current.LevelsFormat.Clear();
            }

            if (current.HasRange)
            {
                if (!TryCreateRangeBookmark(doc, current.From, current.To, out string bookmark, out error))
                {
                    return Fail(error);
                }

                ForgetRangeBookmark(doc, current.Bookmark);
                current.Bookmark = bookmark;
                instructionChanged = true;
            }
            else if (patch.ClearRange)
            {
                ForgetRangeBookmark(doc, current.Bookmark);
                current.Bookmark = null;
                instructionChanged = true;
            }

            if (instructionChanged)
            {
                hit.Field.Code.Text = " " + TocXml.BuildInstruction(current) + " ";
                UpdateField(hit.Field);
            }

            if (patch.HasTitle)
            {
                if (string.IsNullOrEmpty(patch.Title))
                {
                    DeleteTitle(doc, hit.Field);
                    current.HasTitle = false;
                    current.Title = null;
                }
                else if (TitleText(doc, hit.Field) == null)
                {
                    InsertTitleBefore(doc, hit.Id, patch.Title);
                    hit.Field = FindField(doc, hit.Id) ?? hit.Field;
                    current.HasTitle = true;
                    current.Title = patch.Title;
                }
                else
                {
                    ReplaceTitle(doc, hit.Field, patch.Title);
                    current.HasTitle = true;
                    current.Title = patch.Title;
                }
            }

            List<Dictionary<string, object>> warnings = ApplyOverlays(doc, hit.Field, current, current.Reset);
            AddBookmarkAtCode(doc, hit.Field, hit.Id);
            return WithWarnings(Ok(hit.Id, "已更新目录 " + hit.Id, TocXml.Write(current, hit.Id, includeDefaults: false), LinesOf(doc, hit.Field, false)), warnings);
        }

        private static ToolResult ReplacePlaceholder(Word.Document doc, TocHit hit, TocSpec patch)
        {
            TocSpec spec = TocSpec.InsertDefault();
            string error;
            TocXml.Merge(spec, patch, out error);
            if (error != null)
            {
                return Fail(error);
            }

            if (spec.HasRange)
            {
                if (!TryCreateRangeBookmark(doc, spec.From, spec.To, out string bookmark, out error))
                {
                    return Fail(error);
                }

                spec.Bookmark = bookmark;
            }

            int start = hit.Start;
            DeleteRange(doc, hit.Start, hit.End);
            Word.Range point = doc.Range(start, start);
            Word.Field field = AddField(doc, point, TocXml.BuildInstruction(spec));
            if (field == null)
            {
                return Fail("把占位目录换成目录域失败");
            }

            UpdateField(field);
            AddBookmarkAtCode(doc, field, hit.Id);
            if (spec.HasTitle && !string.IsNullOrEmpty(spec.Title))
            {
                InsertTitleBefore(doc, hit.Id, spec.Title);
                field = FindField(doc, hit.Id) ?? field;
            }

            List<Dictionary<string, object>> warnings = ApplyOverlays(doc, field, spec, spec.Reset);
            return WithWarnings(Ok(hit.Id, "已把占位目录换成目录域 " + hit.Id, TocXml.Write(spec, hit.Id, includeDefaults: false), LinesOf(doc, field, false)), warnings);
        }

        private static ToolResult Delete(Word.Document doc, TocHit hit)
        {
            if (hit.Gallery)
            {
                DeleteRange(doc, hit.Start, hit.End);
            }
            else if (hit.Field != null)
            {
                hit.Field.Delete();
            }
            else
            {
                return Fail("这份目录已经不在");
            }

            return Ok(hit.Id, "已删除目录 " + hit.Id, "", new List<Dictionary<string, object>>());
        }

        private static ToolResult Reset(Word.Document doc, TocHit hit)
        {
            if (hit.Field == null)
            {
                return Ok(hit.Id, "占位目录没有直接格式", "", LinesOf(doc, null, true));
            }

            ClearDirect(doc, hit.Field);
            return Ok(hit.Id, "已复位目录 " + hit.Id + " 的直接格式", "", LinesOf(doc, hit.Field, false));
        }

        private static bool TryLocate(Word.Document doc, string id, out TocHit hit, out string error)
        {
            hit = null;
            error = null;
            List<TocSpan> galleries = ListGalleries(doc);
            Word.Field field = FindField(doc, id);
            if (field != null)
            {
                if (!InMainStory(doc, field.Code.Start))
                {
                    error = "这份目录已经不在";
                    return false;
                }

                hit = new TocHit { Id = id, Field = field, Start = field.Code.Start, End = field.Result.End };
                foreach (TocSpan gallery in galleries)
                {
                    if (field.Code.Start >= gallery.Start && field.Code.Start < gallery.End)
                    {
                        hit.Gallery = true;
                        hit.Start = gallery.Start;
                        hit.End = gallery.End;
                        break;
                    }
                }

                return true;
            }

            foreach (TocSpan gallery in galleries)
            {
                if (!string.Equals(gallery.Id, id, StringComparison.Ordinal))
                {
                    continue;
                }

                if (!InMainStory(doc, gallery.Start))
                {
                    error = "这份目录已经不在";
                    return false;
                }

                hit = new TocHit
                {
                    Id = id,
                    Gallery = true,
                    Start = gallery.Start,
                    End = gallery.End
                };
                return true;
            }

            error = "这份目录已经不在";
            return false;
        }

        private static bool TryInsertPoint(Word.Document doc, string at, out Word.Range point, out string error)
        {
            point = null;
            error = null;
            string where = (at ?? "").Trim();
            if (where.Length == 0 || string.Equals(where, "start", StringComparison.OrdinalIgnoreCase))
            {
                point = InsertTargetLocatorHelper.RangeAtDocumentStart(doc);
                return point != null || FailPoint(out error);
            }

            if (where.StartsWith("ewC_", StringComparison.Ordinal))
            {
                if (!TryLocate(doc, where, out TocHit hit, out error))
                {
                    return false;
                }

                point = doc.Range(hit.Start, hit.Start);
                return true;
            }

            if (!ParagraphCode.IsMatch(where))
            {
                error = "at 只能是 start、段编码或目录编号";
                return false;
            }

            if (TocFieldStamp.IsInterior(where))
            {
                error = TocFieldStamp.InteriorMessage;
                return false;
            }

            var located = InsertTargetLocatorHelper.TryResolveCodeSpan(
                doc,
                new[] { where },
                null,
                InsertTargetLocatorHelper.LocateRangeKind.ContentSpan,
                paragraph: true,
                disallowBackendApi: true);
            if (!located.Success || located.Range == null)
            {
                error = string.IsNullOrEmpty(located.Error) ? "找不到插入位置" : located.Error;
                return false;
            }

            point = doc.Range(located.Range.Start, located.Range.Start);
            return true;
        }

        private static bool FailPoint(out string error)
        {
            error = "找不到插入位置";
            return false;
        }

        private static Word.Field AddField(Word.Document doc, Word.Range point, string instruction)
        {
            string switches = instruction ?? "";
            if (switches.StartsWith("TOC", StringComparison.OrdinalIgnoreCase))
            {
                switches = switches.Substring(3).TrimStart();
            }

            object type = Word.WdFieldType.wdFieldTOC;
            object text = switches;
            object preserve = false;
            return doc.Fields.Add(point, ref type, ref text, ref preserve);
        }

        private static void UpdateField(Word.Field field)
        {
            Word.Application app = null;
            Word.WdAlertLevel alerts = Word.WdAlertLevel.wdAlertsNone;
            bool restore = false;
            try
            {
                app = field.Application;
                alerts = app.DisplayAlerts;
                app.DisplayAlerts = Word.WdAlertLevel.wdAlertsNone;
                restore = true;
            }
            catch (Exception)
            {
                restore = false;
            }

            try
            {
                field.Update();
            }
            finally
            {
                if (restore && app != null)
                {
                    try
                    {
                        app.DisplayAlerts = alerts;
                    }
                    catch (Exception)
                    {
                    }
                }
            }
        }

        private static List<Dictionary<string, object>> ApplyOverlays(Word.Document doc, Word.Field field, TocSpec spec, bool reset)
        {
            var warnings = new List<Dictionary<string, object>>();
            if (field == null)
            {
                return warnings;
            }

            if (reset)
            {
                ClearDirect(doc, field);
                return warnings;
            }

            Word.Paragraph title = TitleParagraph(doc, field);
            if (title != null && spec.LevelsFormat.TryGetValue("title", out TocLevelFormat titleFormat))
            {
                WriteLevel(title.Range, titleFormat, warnings);
            }

            if (spec.WriteTabs)
            {
                EnsureLeaderTabs(doc, field, spec);
            }

            foreach (Word.Paragraph paragraph in field.Result.Paragraphs)
            {
                string level = LevelKey(StyleName(paragraph));
                if (level == null || !spec.LevelsFormat.TryGetValue(level, out TocLevelFormat format))
                {
                    continue;
                }

                WriteLevel(paragraph.Range, format, warnings);
                if (format.Font.ContainsKey("reflection"))
                {
                    TocFormatTrace.Write("after WriteLevel Reflection.Type=" + ReadReflectionType(paragraph.Range));
                }
            }

            return warnings;
        }

        private static void EnsureLeaderTabs(Word.Document doc, Word.Field field, TocSpec spec)
        {
            if (!spec.AlignRight)
            {
                return;
            }

            foreach (Word.Paragraph paragraph in field.Result.Paragraphs)
            {
                string level = LevelKey(StyleName(paragraph));
                if (level == null)
                {
                    continue;
                }

                if (!spec.LevelsFormat.TryGetValue(level, out TocLevelFormat format))
                {
                    format = new TocLevelFormat();
                    spec.LevelsFormat[level] = format;
                }

                if (format.Paragraph.ContainsKey("tab_stops"))
                {
                    continue;
                }

                float position = RightTabPosition(paragraph);
                format.Paragraph["tab_stops"] = new List<Dictionary<string, object>>
                {
                    new Dictionary<string, object>
                    {
                        ["position"] = position,
                        ["alignment"] = "wdAlignTabRight",
                        ["leader"] = LeaderEnum(spec.Leader)
                    }
                };
            }
        }

        private static void WriteLevel(Word.Range range, TocLevelFormat format, List<Dictionary<string, object>> warnings)
        {
            if (format.Font.Count > 0)
            {
                Dictionary<string, object> snapshot = FormatInheritHelper.BuildSnapshotFromCharFormat(format.Font);
                FormatInheritHelper.ResetFormatWarnings();
                FormatInheritHelper.ApplySnapshot(range, snapshot, charFormatOnly: true);
                List<Dictionary<string, object>> taken = FormatInheritHelper.TakeFormatWarnings();
                if (taken != null)
                {
                    foreach (Dictionary<string, object> warning in taken)
                    {
                        warning.TryGetValue("key", out object key);
                        warning.TryGetValue("message", out object message);
                        TocFormatTrace.Write("char warning " + key + " " + message);
                    }

                    if (warnings != null)
                    {
                        warnings.AddRange(taken);
                    }
                }

                if (snapshot.ContainsKey("reflection"))
                {
                    TocFormatTrace.Write("after ApplySnapshot Reflection.Type=" + ReadReflectionType(range));
                }
            }

            if (format.Paragraph.Count > 0)
            {
                ParaFormatWriter.Apply(range, format.Paragraph, out _, out _);
            }
        }

        private static void ClearDirect(Word.Document doc, Word.Field field)
        {
            Word.Paragraph title = TitleParagraph(doc, field);
            if (title != null)
            {
                ResetParagraph(title);
            }

            foreach (Word.Paragraph paragraph in field.Result.Paragraphs)
            {
                ResetParagraph(paragraph);
            }
        }

        private static void ResetParagraph(Word.Paragraph paragraph)
        {
            try
            {
                paragraph.Range.Font.Reset();
            }
            catch (Exception)
            {
            }

            try
            {
                paragraph.Range.ParagraphFormat.Reset();
            }
            catch (Exception)
            {
            }
        }

        private static void ReadDirect(Word.Document doc, Word.Field field, TocSpec spec)
        {
            Word.Paragraph title = TitleParagraph(doc, field);
            if (title != null)
            {
                TocLevelFormat format = DiffParagraph(title);
                if (format.Font.Count > 0 || format.Paragraph.Count > 0)
                {
                    spec.LevelsFormat["title"] = format;
                }
            }

            bool seenTab = false;
            foreach (Word.Paragraph paragraph in field.Result.Paragraphs)
            {
                string level = LevelKey(StyleName(paragraph));
                if (level == null || spec.LevelsFormat.ContainsKey(level))
                {
                    continue;
                }

                TocLevelFormat format = DiffParagraph(paragraph);
                if (format.Paragraph.TryGetValue("tab_stops", out object tabs) && tabs != null)
                {
                    seenTab = true;
                }

                if (format.Font.Count > 0 || format.Paragraph.Count > 0)
                {
                    spec.LevelsFormat[level] = format;
                }
            }

            spec.WriteTabs = false;
            if (!seenTab && spec.Separator == null)
            {
                spec.AlignRight = true;
                spec.Leader = "dot";
                spec.HasLeader = false;
                spec.HasAlign = false;
            }
        }

        private static TocLevelFormat DiffParagraph(Word.Paragraph paragraph)
        {
            var format = new TocLevelFormat();
            Word.Style style = paragraph.get_Style() as Word.Style;
            Word.Font font = paragraph.Range.Font;
            Word.Font styleFont = style == null ? null : style.Font;
            DiffFont(format.Font, font, styleFont);
            ReadFontDetail(format.Font, paragraph.Range, font, styleFont);
            Word.ParagraphFormat para = paragraph.Range.ParagraphFormat;
            Word.ParagraphFormat stylePara = style == null ? null : style.ParagraphFormat;
            DiffPara(format.Paragraph, para, stylePara);
            try
            {
                Dictionary<string, object> extracted = ParaFormatReader.Extract(paragraph.Range);
                foreach (KeyValuePair<string, object> pair in extracted)
                {
                    if (pair.Key == "paragraph_style" || pair.Value == null)
                    {
                        continue;
                    }

                    format.Paragraph[pair.Key] = pair.Value;
                }
            }
            catch (Exception)
            {
            }

            List<Dictionary<string, object>> tabs = CustomTabs(para);
            if (tabs.Count > 0)
            {
                format.Paragraph["tab_stops"] = tabs;
            }

            return format;
        }

        private static void ReadFontDetail(Dictionary<string, object> target, Word.Range range, Word.Font font, Word.Font style)
        {
            if (font == null)
            {
                return;
            }

            DiffText(target, "name", SafeString(() => font.Name), style == null ? null : SafeString(() => style.Name));
            DiffText(target, "name_ascii", SafeString(() => font.NameAscii), style == null ? null : SafeString(() => style.NameAscii));
            DiffText(target, "name_bi", SafeString(() => font.NameBi), style == null ? null : SafeString(() => style.NameBi));
            DiffText(target, "name_other", SafeString(() => font.NameOther), style == null ? null : SafeString(() => style.NameOther));
            DiffFloat(target, "size_bi", SafeFloat(() => font.SizeBi), style == null ? float.NaN : SafeFloat(() => style.SizeBi));
            DiffColor(target, "font_color", SafeInt(() => (int)font.Color), style == null ? 9999999 : SafeInt(() => (int)style.Color));
            DiffText(target, "color_index", EnumName(() => font.ColorIndex), style == null ? null : EnumName(() => style.ColorIndex));
            DiffText(target, "color_index_bi", EnumName(() => font.ColorIndexBi), style == null ? null : EnumName(() => style.ColorIndexBi));
            DiffColor(target, "underline_color", SafeInt(() => (int)font.UnderlineColor), style == null ? 9999999 : SafeInt(() => (int)style.UnderlineColor));
            DiffColor(target, "diacritic_color", SafeInt(() => (int)font.DiacriticColor), style == null ? 9999999 : SafeInt(() => (int)style.DiacriticColor));
            DiffTri(target, "bold_bi", SafeInt(() => font.BoldBi), style == null ? 9999999 : SafeInt(() => style.BoldBi));
            DiffTri(target, "italic_bi", SafeInt(() => font.ItalicBi), style == null ? 9999999 : SafeInt(() => style.ItalicBi));
            DiffText(target, "underline", EnumName(() => font.Underline), style == null ? null : EnumName(() => style.Underline));
            if (target.ContainsKey("underline") && target["underline"].ToString().IndexOf("None", StringComparison.Ordinal) < 0)
            {
                target["has_underline"] = true;
            }

            DiffTri(target, "has_strikethrough", SafeInt(() => font.StrikeThrough), style == null ? 9999999 : SafeInt(() => style.StrikeThrough));
            DiffTri(target, "double_strikethrough", SafeInt(() => font.DoubleStrikeThrough), style == null ? 9999999 : SafeInt(() => style.DoubleStrikeThrough));
            DiffTri(target, "is_subscript", SafeInt(() => font.Subscript), style == null ? 9999999 : SafeInt(() => style.Subscript));
            DiffTri(target, "is_superscript", SafeInt(() => font.Superscript), style == null ? 9999999 : SafeInt(() => style.Superscript));
            DiffTri(target, "shadow", SafeInt(() => font.Shadow), style == null ? 9999999 : SafeInt(() => style.Shadow));
            DiffTri(target, "outline", SafeInt(() => font.Outline), style == null ? 9999999 : SafeInt(() => style.Outline));
            DiffTri(target, "emboss", SafeInt(() => font.Emboss), style == null ? 9999999 : SafeInt(() => style.Emboss));
            DiffTri(target, "engrave", SafeInt(() => font.Engrave), style == null ? 9999999 : SafeInt(() => style.Engrave));
            DiffTri(target, "small_caps", SafeInt(() => font.SmallCaps), style == null ? 9999999 : SafeInt(() => style.SmallCaps));
            DiffTri(target, "all_caps", SafeInt(() => font.AllCaps), style == null ? 9999999 : SafeInt(() => style.AllCaps));
            DiffText(target, "emphasis_mark", EnumName(() => font.EmphasisMark), style == null ? null : EnumName(() => style.EmphasisMark));
            DiffFloat(target, "scaling", SafeFloat(() => font.Scaling), style == null ? float.NaN : SafeFloat(() => style.Scaling));
            DiffFloat(target, "spacing", SafeFloat(() => font.Spacing), style == null ? float.NaN : SafeFloat(() => style.Spacing));
            DiffFloat(target, "position", SafeFloat(() => font.Position), style == null ? float.NaN : SafeFloat(() => style.Position));
            DiffFloat(target, "kerning", SafeFloat(() => font.Kerning), style == null ? float.NaN : SafeFloat(() => style.Kerning));
            DiffBool(target, "disable_character_space_grid", SafeBool(() => font.DisableCharacterSpaceGrid), style != null && SafeBool(() => style.DisableCharacterSpaceGrid));
            DiffText(target, "ligatures", EnumName(() => font.Ligatures), style == null ? null : EnumName(() => style.Ligatures));
            DiffText(target, "number_form", EnumName(() => font.NumberForm), style == null ? null : EnumName(() => style.NumberForm));
            DiffText(target, "number_spacing", EnumName(() => font.NumberSpacing), style == null ? null : EnumName(() => style.NumberSpacing));
            DiffText(target, "stylistic_set", EnumName(() => font.StylisticSet), style == null ? null : EnumName(() => style.StylisticSet));
            DiffTri(target, "contextual_alternates", SafeInt(() => font.ContextualAlternates), style == null ? 9999999 : SafeInt(() => style.ContextualAlternates));
            DiffText(target, "highlight_color_index", EnumName(() => range.HighlightColorIndex), null);
            DiffText(target, "language_id", EnumName(() => range.LanguageID), null);
            DiffText(target, "language_id_far_east", EnumName(() => range.LanguageIDFarEast), null);
            int proof = SafeInt(() => range.NoProofing);
            if (proof == 1 || proof == -1)
            {
                target["no_proofing"] = true;
            }
            DiffFloat(target, "fit_text_width", SafeFloat(() => range.FitTextWidth), float.NaN);
            ReadShading(target, "shading", SafeShading(() => font.Shading), style == null ? null : SafeShading(() => style.Shading));
            ReadColorObject(target, "text_color", SafeColorRgb(() => font.TextColor), style == null ? int.MinValue : SafeColorRgb(() => style.TextColor));
            ReadBorders(target, font, style);
            ReadEffect(target, "fill", SafeBool(() => font.Fill.Visible == Microsoft.Office.Core.MsoTriState.msoTrue), "visible");
            ReadEffectNumber(target, "glow", SafeFloat(() => font.Glow.Radius), "radius");
            ReadEffectNumber(target, "line", SafeFloat(() => font.Line.Weight), "weight");
            string reflectionNow = EnumName(() => font.Reflection.Type);
            float reflectionSize = SafeFloat(() => font.Reflection.Size);
            TocFormatTrace.Write("get Reflection.Type=" + (reflectionNow ?? "null") + " Size=" + reflectionSize);
            bool reflectionNone = string.IsNullOrEmpty(reflectionNow)
                || reflectionNow.IndexOf("None", StringComparison.OrdinalIgnoreCase) >= 0
                || reflectionNow.StartsWith("read-failed", StringComparison.Ordinal);
            if (!reflectionNone)
            {
                target["reflection"] = new Dictionary<string, object> { ["type"] = reflectionNow };
            }
            else if (!float.IsNaN(reflectionSize) && reflectionSize > 0f && reflectionSize <= 100f)
            {
                target["reflection"] = new Dictionary<string, object>
                {
                    ["type"] = "msoReflectionTypeNone",
                    ["size"] = reflectionSize
                };
            }
            ReadEffectNumber(target, "text_shadow", SafeFloat(() => font.TextShadow.OffsetX), "offset_x");
        }

        private static void DiffFont(Dictionary<string, object> target, Word.Font font, Word.Font style)
        {
            if (font == null)
            {
                return;
            }

            string name = SafeString(() => font.NameFarEast);
            string styleName = style == null ? null : SafeString(() => style.NameFarEast);
            if (!string.IsNullOrEmpty(name) && !string.Equals(name, styleName, StringComparison.Ordinal))
            {
                target["font_name"] = name;
            }

            float size = SafeFloat(() => font.Size);
            float styleSize = style == null ? float.NaN : SafeFloat(() => style.Size);
            if (size > 0 && size < 9999 && (float.IsNaN(styleSize) || Math.Abs(size - styleSize) > 0.05f))
            {
                target["font_size"] = size;
            }

            int bold = SafeInt(() => font.Bold);
            int styleBold = style == null ? 9999999 : SafeInt(() => style.Bold);
            if (bold == 0 || bold == 1 || bold == -1)
            {
                if (bold != styleBold)
                {
                    target["is_bold"] = bold != 0;
                }
            }

            int italic = SafeInt(() => font.Italic);
            int styleItalic = style == null ? 9999999 : SafeInt(() => style.Italic);
            if ((italic == 0 || italic == 1 || italic == -1) && italic != styleItalic)
            {
                target["is_italic"] = italic != 0;
            }
        }

        private static void DiffPara(
            Dictionary<string, object> target,
            Word.ParagraphFormat para,
            Word.ParagraphFormat style)
        {
            if (para == null)
            {
                return;
            }

            string align = AlignName(SafeInt(() => (int)para.Alignment));
            string styleAlign = style == null ? null : AlignName(SafeInt(() => (int)style.Alignment));
            if (align != null && !string.Equals(align, styleAlign, StringComparison.Ordinal))
            {
                target["alignment"] = align;
            }

            DiffPoints(target, "left_indent", SafeFloat(() => para.LeftIndent), style == null ? float.NaN : SafeFloat(() => style.LeftIndent));
            DiffPoints(target, "right_indent", SafeFloat(() => para.RightIndent), style == null ? float.NaN : SafeFloat(() => style.RightIndent));
            DiffPoints(target, "first_line_indent", SafeFloat(() => para.FirstLineIndent), style == null ? float.NaN : SafeFloat(() => style.FirstLineIndent));
            DiffPoints(target, "space_before", SafeFloat(() => para.SpaceBefore), style == null ? float.NaN : SafeFloat(() => style.SpaceBefore));
            DiffPoints(target, "space_after", SafeFloat(() => para.SpaceAfter), style == null ? float.NaN : SafeFloat(() => style.SpaceAfter));
            DiffPoints(target, "line_spacing", SafeFloat(() => para.LineSpacing), style == null ? float.NaN : SafeFloat(() => style.LineSpacing));
        }

        private static void DiffPoints(Dictionary<string, object> target, string key, float value, float style)
        {
            if (float.IsNaN(value) || value > 999999)
            {
                return;
            }

            if (float.IsNaN(style) || Math.Abs(value - style) > 0.05f)
            {
                target[key] = value;
            }
        }

        private static List<Dictionary<string, object>> CustomTabs(Word.ParagraphFormat para)
        {
            var tabs = new List<Dictionary<string, object>>();
            if (para == null)
            {
                return tabs;
            }

            foreach (Word.TabStop tab in para.TabStops)
            {
                bool custom;
                try
                {
                    custom = tab.CustomTab;
                }
                catch (Exception)
                {
                    custom = false;
                }

                if (!custom)
                {
                    continue;
                }

                tabs.Add(new Dictionary<string, object>
                {
                    ["position"] = tab.Position,
                    ["alignment"] = TabAlignEnum(tab.Alignment),
                    ["leader"] = LeaderEnum(LeaderName(tab.Leader))
                });
            }

            return tabs;
        }

        private static List<Dictionary<string, object>> LinesOf(Word.Document doc, Word.Field field, bool placeholder)
        {
            var lines = new List<Dictionary<string, object>>();
            if (field == null)
            {
                return lines;
            }

            foreach (Word.Paragraph paragraph in field.Result.Paragraphs)
            {
                string raw = VisibleText((paragraph.Range.Text ?? "").TrimEnd('\r', '\a', '\n', '\u0007'));
                if (raw.Length == 0)
                {
                    continue;
                }

                SplitEntry(raw, out string text, out string page);
                if (IsEmptyResult(text))
                {
                    continue;
                }

                bool hold = placeholder || text.IndexOf("键入章标题", StringComparison.Ordinal) >= 0
                    || text.IndexOf("Type chapter title", StringComparison.OrdinalIgnoreCase) >= 0
                    || text.IndexOf("Type the chapter title", StringComparison.OrdinalIgnoreCase) >= 0;
                lines.Add(new Dictionary<string, object>
                {
                    ["level"] = LevelKey(StyleName(paragraph)) ?? "",
                    ["text"] = text,
                    ["page"] = hold ? "" : page,
                    ["placeholder"] = hold
                });
            }

            return lines;
        }

        private static void SplitEntry(string raw, out string text, out string page)
        {
            page = "";
            int tab = raw.LastIndexOf('\t');
            text = tab >= 0 ? raw.Substring(0, tab) : raw;
            if (tab >= 0 && tab + 1 < raw.Length)
            {
                page = raw.Substring(tab + 1).Trim();
            }

            text = text.Trim().TrimEnd('.', '·', '…', '。');
        }

        private static bool IsEmptyResult(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
            {
                return true;
            }

            return text.IndexOf("No table of contents entries", StringComparison.OrdinalIgnoreCase) >= 0
                || text.IndexOf("没有找到目录项", StringComparison.Ordinal) >= 0
                || text.IndexOf("未找到目录项", StringComparison.Ordinal) >= 0;
        }

        private static string PlaceholderText(TocHit hit)
        {
            return "键入章标题";
        }

        private static void InsertTitle(Word.Document doc, Word.Field field, string text)
        {
            string id = FindBookmarkAt(doc, field.Code.Start);
            if (string.IsNullOrEmpty(id))
            {
                id = AddBookmarkAtCode(doc, field, null);
            }

            InsertTitleBefore(doc, id, text);
        }

        private static void InsertTitleBefore(Word.Document doc, string bookmarkId, string text)
        {
            Word.Bookmark mark = doc.Bookmarks[bookmarkId];
            int pos = mark.Range.Start;
            if (pos > doc.Content.Start)
            {
                pos -= 1;
            }

            string body = text ?? "";
            Word.Range at = doc.Range(pos, pos);
            at.InsertBefore(body + "\r");
            Word.Paragraph paragraph = doc.Range(pos, pos).Paragraphs[1];
            ApplyTitleStyle(paragraph);
        }

        private static string FindBookmarkAt(Word.Document doc, int start)
        {
            foreach (Word.Bookmark mark in doc.Bookmarks)
            {
                string name = mark.Name ?? "";
                if (name.StartsWith("ewC_", StringComparison.Ordinal) && Math.Abs(mark.Range.Start - start) <= 1)
                {
                    return name;
                }
            }

            return null;
        }

        private static string VisibleText(string raw)
        {
            if (string.IsNullOrEmpty(raw) || raw.IndexOf('\u0013') < 0)
            {
                return raw ?? "";
            }

            var buffer = new System.Text.StringBuilder();
            int i = 0;
            while (i < raw.Length)
            {
                if (raw[i] == '\u0013')
                {
                    int sep = raw.IndexOf('\u0014', i);
                    int end = raw.IndexOf('\u0015', i);
                    if (sep >= 0 && end > sep)
                    {
                        buffer.Append(raw, sep + 1, end - sep - 1);
                        i = end + 1;
                        continue;
                    }
                }

                if (raw[i] != '\u0014' && raw[i] != '\u0015')
                {
                    buffer.Append(raw[i]);
                }

                i++;
            }

            return buffer.ToString();
        }

        private static void ReplaceTitle(Word.Document doc, Word.Field field, string text)
        {
            Word.Paragraph paragraph = TitleParagraph(doc, field);
            if (paragraph == null)
            {
                InsertTitle(doc, field, text);
                return;
            }

            paragraph.Range.Text = (text ?? "") + "\r";
            ApplyTitleStyle(paragraph);
        }

        private static void DeleteTitle(Word.Document doc, Word.Field field)
        {
            Word.Paragraph paragraph = TitleParagraph(doc, field);
            if (paragraph != null)
            {
                paragraph.Range.Delete();
            }
        }

        private static string TitleText(Word.Document doc, Word.Field field)
        {
            Word.Paragraph paragraph = TitleParagraph(doc, field);
            if (paragraph == null)
            {
                return null;
            }

            return (paragraph.Range.Text ?? "").TrimEnd('\r', '\a', '\n');
        }

        private static Word.Paragraph TitleParagraph(Word.Document doc, Word.Field field)
        {
            if (field == null || field.Code.Start <= doc.Content.Start)
            {
                return null;
            }

            int fieldBegin = field.Code.Start - 1;
            if (fieldBegin <= doc.Content.Start)
            {
                return null;
            }

            Word.Range before = doc.Range(fieldBegin - 1, fieldBegin - 1);
            Word.Paragraph paragraph = before.Paragraphs[1];
            if (paragraph == null || paragraph.Range.End > fieldBegin + 1)
            {
                return null;
            }

            if (IsTitleStyle(paragraph))
            {
                return paragraph;
            }

            return null;
        }

        private static bool TryCreateRangeBookmark(
            Word.Document doc,
            string from,
            string to,
            out string name,
            out string error)
        {
            name = null;
            error = null;
            if (string.IsNullOrEmpty(from) || string.IsNullOrEmpty(to))
            {
                error = "起止要成对";
                return false;
            }

            if (TocFieldStamp.IsInterior(from) || TocFieldStamp.IsInterior(to))
            {
                error = TocFieldStamp.InteriorMessage;
                return false;
            }

            var located = InsertTargetLocatorHelper.TryResolveCodeSpan(
                doc,
                new[] { from, to },
                null,
                InsertTargetLocatorHelper.LocateRangeKind.ContentSpan,
                paragraph: true,
                disallowBackendApi: true);
            if (!located.Success || located.Range == null)
            {
                error = string.IsNullOrEmpty(located.Error) ? "范围已经不在" : located.Error;
                return false;
            }

            name = NewRangeName(doc);
            object anchor = located.Range;
            doc.Bookmarks.Add(name, ref anchor);
            return true;
        }

        private static bool TryBookmarkToCodes(
            Word.Document doc,
            string bookmark,
            out string from,
            out string to,
            out string error)
        {
            from = null;
            to = null;
            error = null;
            Word.Bookmark mark;
            try
            {
                mark = doc.Bookmarks[bookmark];
            }
            catch (Exception)
            {
                error = "范围已经不在";
                return false;
            }

            if (mark == null || mark.Range == null)
            {
                error = "范围已经不在";
                return false;
            }

            Word.Paragraphs paragraphs = mark.Range.Paragraphs;
            from = CodeOfParagraph(paragraphs[1]);
            to = CodeOfParagraph(paragraphs[paragraphs.Count]);
            if (string.IsNullOrEmpty(from) || string.IsNullOrEmpty(to))
            {
                error = "范围已经不在";
                return false;
            }

            return true;
        }

        private static string CodeOfParagraph(Word.Paragraph paragraph)
        {
            if (paragraph == null)
            {
                return null;
            }

            string text = (paragraph.Range.Text ?? "").TrimEnd('\r', '\a', '\n', '\u0007');
            if (text.Length == 0)
            {
                return null;
            }

            string found = null;
            foreach (IReadOnlyList<string> mapping in DocumentState.ParagraphNameMapping)
            {
                if (mapping.Count < 3 || mapping[2] != text)
                {
                    continue;
                }

                if (found != null && !string.Equals(found, mapping[0], StringComparison.Ordinal))
                {
                    return null;
                }

                found = mapping[0];
            }

            return found;
        }

        private static void ForgetRangeBookmark(Word.Document doc, string name)
        {
            if (string.IsNullOrEmpty(name) || !name.StartsWith("ewR_", StringComparison.Ordinal))
            {
                return;
            }

            try
            {
                if (doc.Bookmarks.Exists(name))
                {
                    doc.Bookmarks[name].Delete();
                }
            }
            catch (Exception)
            {
            }
        }

        private static string AddBookmarkAtCode(Word.Document doc, Word.Field field, string id)
        {
            Word.Range point = field.Code.Duplicate;
            point.Collapse(Word.WdCollapseDirection.wdCollapseStart);
            if (string.IsNullOrEmpty(id))
            {
                id = NewTocId(doc);
            }
            else
            {
                ForgetBookmark(doc, id);
            }

            object anchor = point;
            doc.Bookmarks.Add(id, ref anchor);
            return id;
        }

        private static void ForgetBookmark(Word.Document doc, string name)
        {
            try
            {
                if (doc.Bookmarks.Exists(name))
                {
                    doc.Bookmarks[name].Delete();
                }
            }
            catch (Exception)
            {
            }
        }

        private static Word.Field FindField(Word.Document doc, string id)
        {
            Word.Bookmark mark = null;
            try
            {
                if (doc.Bookmarks.Exists(id))
                {
                    mark = doc.Bookmarks[id];
                }
            }
            catch (Exception)
            {
                mark = null;
            }

            if (mark == null)
            {
                return null;
            }

            int at = mark.Range.Start;
            foreach (Word.Field field in doc.Fields)
            {
                if (!IsTocField(field))
                {
                    continue;
                }

                int start = field.Code.Start;
                if (start == at || start == at - 1 || start == at + 1)
                {
                    return field;
                }
            }

            return null;
        }

        private static bool IsTocField(Word.Field field)
        {
            try
            {
                if (field.Type == Word.WdFieldType.wdFieldTOC)
                {
                    return true;
                }
            }
            catch (Exception)
            {
            }

            try
            {
                return TocFieldStamp.IsTocInstruction(field.Code.Text);
            }
            catch (Exception)
            {
                return false;
            }
        }

        private static void DeleteRange(Word.Document doc, int start, int end)
        {
            if (end < start)
            {
                int swap = start;
                start = end;
                end = swap;
            }

            doc.Range(start, end).Delete();
        }

        private static bool InMainStory(Word.Document doc, int start)
        {
            return start >= doc.Content.Start && start < doc.Content.End;
        }

        private static List<TocSpan> ListGalleries(Word.Document doc)
        {
            var spans = new List<TocSpan>();
            string xml;
            try
            {
                xml = doc.Content.WordOpenXML;
            }
            catch (Exception)
            {
                return spans;
            }

            if (string.IsNullOrEmpty(xml))
            {
                return spans;
            }

            XDocument package;
            try
            {
                package = XDocument.Parse(xml);
            }
            catch (Exception)
            {
                return spans;
            }

            XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
            XElement body = package.Descendants(w + "body").FirstOrDefault();
            if (body == null)
            {
                return spans;
            }

            int index = 0;
            WalkGallery(body, w, false, null, ref index, spans);
            foreach (TocSpan span in spans)
            {
                if (span.FirstPara <= 0 || doc.Paragraphs.Count < span.LastPara)
                {
                    continue;
                }

                span.Start = doc.Paragraphs[span.FirstPara].Range.Start;
                span.End = doc.Paragraphs[span.LastPara].Range.End;
            }

            spans.RemoveAll(span => span.Start == 0 && span.End == 0);
            return spans;
        }

        private static void WalkGallery(
            XElement element,
            XNamespace w,
            bool inDeleted,
            TocSpan gallery,
            ref int index,
            List<TocSpan> spans)
        {
            foreach (XElement child in element.Elements())
            {
                string name = child.Name.LocalName;
                if (name == "txbxContent")
                {
                    continue;
                }

                if (name == "sdt")
                {
                    TocSpan opened = null;
                    if (!inDeleted && gallery == null && TocFieldStamp.IsPlainTocGallery(child))
                    {
                        opened = new TocSpan();
                        opened.Id = (string)child.Element(w + "sdtPr")?.Element(w + "tag")?.Attribute(w + "val");
                    }

                    WalkGallery(child, w, inDeleted, opened ?? gallery, ref index, spans);
                    if (opened != null)
                    {
                        spans.Add(opened);
                    }

                    continue;
                }

                if (name == "p")
                {
                    index++;
                    if (gallery != null)
                    {
                        if (gallery.FirstPara == 0)
                        {
                            gallery.FirstPara = index;
                        }

                        gallery.LastPara = index;
                    }

                    continue;
                }

                if (name == "tr")
                {
                    WalkGallery(child, w, inDeleted, gallery, ref index, spans);
                    index++;
                    continue;
                }

                if (name == "del")
                {
                    WalkGallery(child, w, true, gallery, ref index, spans);
                    continue;
                }

                if (name.EndsWith("Pr", StringComparison.Ordinal))
                {
                    continue;
                }

                WalkGallery(child, w, inDeleted, gallery, ref index, spans);
            }
        }

        private static void ApplyTitleStyle(Word.Paragraph paragraph)
        {
            if (paragraph == null)
            {
                return;
            }

            Word.Style builtin = TocHeadingStyle(paragraph.Range.Document);
            if (builtin != null)
            {
                try
                {
                    object style = builtin;
                    paragraph.set_Style(ref style);
                    return;
                }
                catch (Exception)
                {
                }
            }

            TrySetParagraphStyle(paragraph, "目录标题", "TOC 标题", "TOC Heading");
        }

        private static bool IsTitleStyle(Word.Paragraph paragraph)
        {
            Word.Style style = paragraph == null ? null : paragraph.get_Style() as Word.Style;
            Word.Style builtin = paragraph == null ? null : TocHeadingStyle(paragraph.Range.Document);
            if (style != null && builtin != null && SameStyle(style, builtin))
            {
                return true;
            }

            string name = StyleName(paragraph);
            return name != null
                && (name.IndexOf("目录标题", StringComparison.Ordinal) >= 0
                    || name.IndexOf("TOC 标题", StringComparison.OrdinalIgnoreCase) >= 0
                    || string.Equals(name, "TOC Heading", StringComparison.OrdinalIgnoreCase));
        }

        private static Word.Style TocHeadingStyle(Word.Document doc)
        {
            if (doc == null)
            {
                return null;
            }

            try
            {
                return doc.Styles[Word.WdBuiltinStyle.wdStyleTocHeading];
            }
            catch (Exception)
            {
                return null;
            }
        }

        private static bool SameStyle(Word.Style left, Word.Style right)
        {
            try
            {
                string localLeft = left.NameLocal ?? "";
                string localRight = right.NameLocal ?? "";
                return localLeft.Length > 0 && string.Equals(localLeft, localRight, StringComparison.OrdinalIgnoreCase);
            }
            catch (Exception)
            {
                return false;
            }
        }

        private static void TrySetParagraphStyle(Word.Paragraph paragraph, params string[] names)
        {
            foreach (string name in names)
            {
                try
                {
                    object style = name;
                    paragraph.set_Style(ref style);
                    return;
                }
                catch (Exception)
                {
                }
            }
        }

        private static string StyleName(Word.Paragraph paragraph)
        {
            try
            {
                Word.Style style = paragraph.get_Style() as Word.Style;
                if (style == null)
                {
                    return null;
                }

                string local = style.NameLocal;
                return string.IsNullOrEmpty(local) ? style.NameLocal : local;
            }
            catch (Exception)
            {
                return null;
            }
        }

        private static string LevelKey(string styleName)
        {
            if (string.IsNullOrEmpty(styleName))
            {
                return null;
            }

            Match match = Regex.Match(styleName, @"(?:目录|TOC)\s*([1-9])\b", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
            if (match.Success)
            {
                return match.Groups[1].Value;
            }

            return null;
        }

        private static float RightTabPosition(Word.Paragraph paragraph)
        {
            try
            {
                Word.Style style = paragraph.get_Style() as Word.Style;
                if (style != null)
                {
                    foreach (Word.TabStop tab in style.ParagraphFormat.TabStops)
                    {
                        if (tab.Alignment == Word.WdTabAlignment.wdAlignTabRight)
                        {
                            return tab.Position;
                        }
                    }
                }
            }
            catch (Exception)
            {
            }

            return 415f;
        }

        private static string ExtractSwitch(string code, string letter)
        {
            Match match = Regex.Match(
                code ?? "",
                @"\\" + letter + "\\s+\"([^\"]*)\"",
                RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
            return match.Success ? match.Groups[1].Value : null;
        }

        private static string NewTocId(Word.Document doc)
        {
            for (int i = 0; i < 8; i++)
            {
                string name = "ewC_" + Guid.NewGuid().ToString("N").Substring(0, 8);
                try
                {
                    if (!doc.Bookmarks.Exists(name))
                    {
                        return name;
                    }
                }
                catch (Exception)
                {
                    return name;
                }
            }

            return "ewC_" + Guid.NewGuid().ToString("N").Substring(0, 12);
        }

        private static string NewRangeName(Word.Document doc)
        {
            for (int i = 0; i < 8; i++)
            {
                string name = "ewR_" + Guid.NewGuid().ToString("N").Substring(0, 8);
                try
                {
                    if (!doc.Bookmarks.Exists(name))
                    {
                        return name;
                    }
                }
                catch (Exception)
                {
                    return name;
                }
            }

            return "ewR_" + Guid.NewGuid().ToString("N").Substring(0, 12);
        }

        private static string AlignName(int value)
        {
            if (value == (int)Word.WdParagraphAlignment.wdAlignParagraphLeft)
            {
                return "left";
            }

            if (value == (int)Word.WdParagraphAlignment.wdAlignParagraphCenter)
            {
                return "center";
            }

            if (value == (int)Word.WdParagraphAlignment.wdAlignParagraphRight)
            {
                return "right";
            }

            if (value == (int)Word.WdParagraphAlignment.wdAlignParagraphJustify)
            {
                return "justify";
            }

            return null;
        }

        private static string TabAlignEnum(Word.WdTabAlignment alignment)
        {
            if (alignment == Word.WdTabAlignment.wdAlignTabRight)
            {
                return "wdAlignTabRight";
            }

            if (alignment == Word.WdTabAlignment.wdAlignTabCenter)
            {
                return "wdAlignTabCenter";
            }

            if (alignment == Word.WdTabAlignment.wdAlignTabDecimal)
            {
                return "wdAlignTabDecimal";
            }

            return "wdAlignTabLeft";
        }

        internal static string LeaderEnum(string leader)
        {
            switch ((leader ?? "").Trim())
            {
                case "dot":
                    return "wdTabLeaderDots";
                case "dash":
                    return "wdTabLeaderDashes";
                case "line":
                    return "wdTabLeaderLines";
                case "heavy":
                    return "wdTabLeaderHeavy";
                case "middleDot":
                    return "wdTabLeaderMiddleDot";
                case "none":
                    return "wdTabLeaderSpaces";
                default:
                    return string.IsNullOrEmpty(leader) ? "wdTabLeaderDots" : leader;
            }
        }

        private static string LeaderName(Word.WdTabLeader leader)
        {
            if (leader == Word.WdTabLeader.wdTabLeaderDots)
            {
                return "dot";
            }

            if (leader == Word.WdTabLeader.wdTabLeaderDashes)
            {
                return "dash";
            }

            if (leader == Word.WdTabLeader.wdTabLeaderLines)
            {
                return "line";
            }

            if (leader == Word.WdTabLeader.wdTabLeaderHeavy)
            {
                return "heavy";
            }

            if (leader == Word.WdTabLeader.wdTabLeaderMiddleDot)
            {
                return "middleDot";
            }

            return "none";
        }

        private static void DiffText(Dictionary<string, object> target, string key, string value, string style)
        {
            if (string.IsNullOrEmpty(value) || value.IndexOf("Undefined", StringComparison.Ordinal) >= 0 || value == "9999999")
            {
                return;
            }

            if (!string.Equals(value, style, StringComparison.Ordinal))
            {
                target[key] = value;
            }
        }

        private static void DiffFloat(Dictionary<string, object> target, string key, float value, float style)
        {
            if (float.IsNaN(value) || value > 999998 || (float.IsNaN(style) && Math.Abs(value) < 0.05f))
            {
                return;
            }

            if (float.IsNaN(style) || Math.Abs(value - style) > 0.05f)
            {
                target[key] = value;
            }
        }

        private static void DiffColor(Dictionary<string, object> target, string key, int value, int style)
        {
            if (value == 9999999 || value == style)
            {
                return;
            }

            string hex = WordColorToHex(value);
            if (hex != null)
            {
                target[key] = hex;
            }
        }

        private static void DiffTri(Dictionary<string, object> target, string key, int value, int style)
        {
            if (value != 0 && value != 1 && value != -1)
            {
                return;
            }

            if (value != style)
            {
                target[key] = value != 0;
            }
        }

        private static void DiffBool(Dictionary<string, object> target, string key, bool value, bool style)
        {
            if (value != style)
            {
                target[key] = value;
            }
        }

        private static string WordColorToHex(int wordColor)
        {
            if (wordColor < 0 || wordColor > 0xFFFFFF)
            {
                return null;
            }

            int red = wordColor & 0xFF;
            int green = (wordColor >> 8) & 0xFF;
            int blue = (wordColor >> 16) & 0xFF;
            return "#" + red.ToString("X2") + green.ToString("X2") + blue.ToString("X2");
        }

        private static string EnumName(Func<object> read)
        {
            try
            {
                object value = read();
                return value == null ? null : value.ToString();
            }
            catch (Exception)
            {
                return null;
            }
        }

        private static bool SafeBool(Func<bool> read)
        {
            try
            {
                return read();
            }
            catch (Exception)
            {
                return false;
            }
        }

        private static Word.Shading SafeShading(Func<Word.Shading> read)
        {
            try
            {
                return read();
            }
            catch (Exception)
            {
                return null;
            }
        }

        private static int SafeColorRgb(Func<Word.ColorFormat> read)
        {
            try
            {
                Word.ColorFormat color = read();
                return color == null ? int.MinValue : color.RGB;
            }
            catch (Exception)
            {
                return int.MinValue;
            }
        }

        private static void ReadShading(Dictionary<string, object> target, string key, Word.Shading mine, Word.Shading style)
        {
            if (mine == null)
            {
                return;
            }

            string texture = EnumName(() => mine.Texture);
            string styleTexture = style == null ? null : EnumName(() => style.Texture);
            if (string.IsNullOrEmpty(texture) || string.Equals(texture, styleTexture, StringComparison.Ordinal))
            {
                return;
            }

            target[key] = new Dictionary<string, object>
            {
                ["texture"] = texture,
                ["background_pattern_color"] = WordColorToHex(SafeInt(() => (int)mine.BackgroundPatternColor)) ?? ""
            };
        }

        private static void ReadColorObject(Dictionary<string, object> target, string key, int rgb, int styleRgb)
        {
            if (rgb == int.MinValue || rgb == styleRgb)
            {
                return;
            }

            string hex = WordColorToHex(rgb);
            if (hex != null)
            {
                target[key] = new Dictionary<string, object> { ["rgb"] = hex };
            }
        }

        private static void ReadBorders(Dictionary<string, object> target, Word.Font font, Word.Font style)
        {
            string mine = EnumName(() => font.Borders.OutsideLineStyle);
            string other = style == null ? null : EnumName(() => style.Borders.OutsideLineStyle);
            if (string.IsNullOrEmpty(mine) || mine.IndexOf("None", StringComparison.Ordinal) >= 0 || string.Equals(mine, other, StringComparison.Ordinal))
            {
                return;
            }

            target["borders"] = new Dictionary<string, object> { ["outside_line_style"] = mine };
        }

        private static string ReadReflectionType(Word.Range range)
        {
            try
            {
                return range.Font.Reflection.Type.ToString();
            }
            catch (Exception ex)
            {
                return "read-failed " + ex.GetType().Name + " " + ex.Message;
            }
        }

        private static void ReadEffect(Dictionary<string, object> target, string key, bool on, string token)
        {
            if (!on || string.IsNullOrEmpty(token))
            {
                return;
            }

            target[key] = new Dictionary<string, object> { ["value"] = token };
        }

        private static void ReadEffectNumber(Dictionary<string, object> target, string key, float value, string token)
        {
            if (float.IsNaN(value) || value <= 0 || value > 999998)
            {
                return;
            }

            target[key] = new Dictionary<string, object> { [token] = value };
        }

        private static string SafeString(Func<string> read)
        {
            try
            {
                return read();
            }
            catch (Exception)
            {
                return null;
            }
        }

        private static float SafeFloat(Func<float> read)
        {
            try
            {
                return read();
            }
            catch (Exception)
            {
                return float.NaN;
            }
        }

        private static int SafeInt(Func<int> read)
        {
            try
            {
                return read();
            }
            catch (Exception)
            {
                return 9999999;
            }
        }

        private static ToolResult WithWarnings(ToolResult result, List<Dictionary<string, object>> warnings)
        {
            if (result != null && result.Success && warnings != null && warnings.Count > 0
                && result.Data is Dictionary<string, object> data)
            {
                data["format_warnings"] = warnings;
            }

            return result;
        }

        private static ToolResult Ok(
            string id,
            string message,
            string xml,
            List<Dictionary<string, object>> lines,
            bool empty = false)
        {
            return new ToolResult
            {
                Success = true,
                Data = new Dictionary<string, object>
                {
                    ["toc_id"] = id,
                    ["message"] = message,
                    ["xml"] = xml ?? "",
                    ["lines"] = lines ?? new List<Dictionary<string, object>>(),
                    ["empty"] = empty || lines == null || lines.Count == 0
                }
            };
        }

        private static ToolResult Fail(string error)
        {
            return new ToolResult { Success = false, Error = error ?? "目录操作失败" };
        }

        private sealed class TocHit
        {
            public string Id;
            public Word.Field Field;
            public bool Gallery;
            public int Start;
            public int End;
        }

        private sealed class TocSpan
        {
            public string Id;
            public int FirstPara;
            public int LastPara;
            public int Start;
            public int End;
        }
    }

    public sealed class TocLevelFormat
    {
        public Dictionary<string, object> Font = new Dictionary<string, object>(StringComparer.Ordinal);
        public Dictionary<string, object> Paragraph = new Dictionary<string, object>(StringComparer.Ordinal);
    }

    public sealed class TocSpec
    {
        public string Levels = "1-3";
        public bool HasLevels;
        public List<TocStyleRef> Styles = new List<TocStyleRef>();
        public bool HasStyles;
        public bool ClearStyles;
        public string From;
        public string To;
        public bool HasRange;
        public bool ClearRange;
        public string Bookmark;
        public bool PageNumbers = true;
        public string OmitPages;
        public bool HasPageNumbers;
        public bool AlignRight = true;
        public bool HasAlign;
        public string Leader = "dot";
        public bool HasLeader;
        public bool WriteTabs;
        public string Separator;
        public bool Hyperlink = true;
        public bool HasHyperlink;
        public bool WebHide = true;
        public bool HasWebHide;
        public string ChapterSeq;
        public string ChapterSep;
        public bool HasChapter;
        public bool ClearChapter;
        public bool KeepTabs;
        public bool HasKeepTabs;
        public bool KeepBreaks;
        public bool HasKeepBreaks;
        public bool HasTitle;
        public string Title;
        public bool Reset;
        public Dictionary<string, TocLevelFormat> LevelsFormat =
            new Dictionary<string, TocLevelFormat>(StringComparer.Ordinal);

        public static TocSpec InsertDefault()
        {
            return new TocSpec { Levels = "1-3", HasLevels = true, Hyperlink = true, WebHide = true, AlignRight = true, Leader = "dot" };
        }
    }

    public sealed class TocStyleRef
    {
        public string Name;
        public string Level;
    }

    public static class TocXml
    {
        public static bool TryParse(string xml, out TocSpec spec, out string error)
        {
            spec = new TocSpec();
            error = null;
            XDocument document;
            try
            {
                document = XDocument.Parse(xml);
            }
            catch (Exception)
            {
                error = "目录 XML 无法解析";
                return false;
            }

            XElement root = document.Root;
            if (root == null || !string.Equals(root.Name.LocalName, "toc", StringComparison.Ordinal))
            {
                error = "目录 XML 须以 toc 为根";
                return false;
            }

            XElement content = root.Element("content");
            if (content != null)
            {
                XElement levels = content.Element("levels");
                if (levels != null)
                {
                    spec.HasLevels = true;
                    spec.Levels = (levels.Value ?? "").Trim();
                    if (!ValidLevels(spec.Levels, out error))
                    {
                        return false;
                    }
                }

                XElement styles = content.Element("styles");
                if (styles != null)
                {
                    spec.HasStyles = true;
                    spec.ClearStyles = !styles.Elements("style").Any();
                    foreach (XElement style in styles.Elements("style"))
                    {
                        string name = (string)style.Attribute("name") ?? "";
                        string level = ((string)style.Attribute("level") ?? "").Trim();
                        if (name.Trim().Length == 0)
                        {
                            error = "样式名不能为空";
                            return false;
                        }

                        if (name.IndexOf(',') >= 0 || name.IndexOf('"') >= 0)
                        {
                            error = "样式名不能含逗号或引号";
                            return false;
                        }

                        if (!ValidOneLevel(level))
                        {
                            error = "样式收进的级别须为 1 到 9";
                            return false;
                        }

                        spec.Styles.Add(new TocStyleRef { Name = name.Trim(), Level = level });
                    }
                }

                XElement range = content.Element("range");
                if (range != null)
                {
                    string from = (string)range.Attribute("from");
                    string to = (string)range.Attribute("to");
                    bool hasFrom = !string.IsNullOrWhiteSpace(from);
                    bool hasTo = !string.IsNullOrWhiteSpace(to);
                    if (hasFrom != hasTo)
                    {
                        error = "起止要成对";
                        return false;
                    }

                    if (!hasFrom)
                    {
                        spec.ClearRange = true;
                    }
                    else
                    {
                        spec.HasRange = true;
                        spec.From = from.Trim();
                        spec.To = to.Trim();
                    }
                }
            }

            XElement format = root.Element("format");
            if (format != null)
            {
                spec.Reset = IsTrue((string)format.Attribute("reset"));
                foreach (XElement page in format.Elements("pageNumbers"))
                {
                    spec.HasPageNumbers = true;
                    string levels = (string)page.Attribute("levels");
                    bool show = !string.Equals((string)page.Attribute("show"), "false", StringComparison.OrdinalIgnoreCase);
                    if (string.IsNullOrWhiteSpace(levels))
                    {
                        spec.PageNumbers = show;
                        if (show)
                        {
                            spec.OmitPages = null;
                        }
                        else
                        {
                            spec.OmitPages = "1-9";
                        }
                    }
                    else
                    {
                        levels = levels.Trim();
                        if (!ValidLevels(levels, out error))
                        {
                            return false;
                        }

                        if (!show)
                        {
                            spec.OmitPages = levels;
                            spec.PageNumbers = true;
                        }
                    }
                }

                XElement align = format.Element("align");
                if (align != null)
                {
                    spec.HasAlign = true;
                    spec.AlignRight = !string.Equals(align.Value.Trim(), "after", StringComparison.OrdinalIgnoreCase)
                        && !string.Equals(align.Value.Trim(), "left", StringComparison.OrdinalIgnoreCase);
                    spec.WriteTabs = true;
                }

                XElement leader = format.Element("leader");
                if (leader != null)
                {
                    spec.HasLeader = true;
                    spec.Leader = leader.Value.Trim();
                    if (!KnownLeader(spec.Leader))
                    {
                        error = "前导符只能是 none、dot、dash、line、heavy、middleDot";
                        return false;
                    }

                    spec.WriteTabs = true;
                }

                XElement separator = format.Element("separator");
                if (separator != null)
                {
                    spec.Separator = separator.Value ?? "";
                    if (spec.Separator.Length > 5)
                    {
                        error = "分隔符最多 5 个字符";
                        return false;
                    }
                }

                spec.Hyperlink = ReadBool(format.Element("hyperlink"), spec.Hyperlink, out bool hasHyperlink);
                spec.HasHyperlink = hasHyperlink;
                spec.WebHide = ReadBool(format.Element("webHidePageNumbers"), spec.WebHide, out bool hasWeb);
                spec.HasWebHide = hasWeb;
                spec.KeepTabs = ReadBool(format.Element("keepTabs"), spec.KeepTabs, out bool hasTabs);
                spec.HasKeepTabs = hasTabs;
                spec.KeepBreaks = ReadBool(format.Element("keepBreaks"), spec.KeepBreaks, out bool hasBreaks);
                spec.HasKeepBreaks = hasBreaks;

                XElement chapter = format.Element("chapter");
                if (chapter != null)
                {
                    spec.HasChapter = true;
                    spec.ChapterSeq = (string)chapter.Attribute("seq");
                    spec.ChapterSep = (string)chapter.Attribute("separator");
                    spec.ClearChapter = string.IsNullOrWhiteSpace(spec.ChapterSeq);
                }

                XElement title = format.Element("title");
                if (title != null)
                {
                    spec.HasTitle = true;
                    spec.Title = (string)title.Attribute("text") ?? "";
                }

                if (!spec.Reset)
                {
                    foreach (XElement level in format.Elements("level"))
                    {
                        string n = ((string)level.Attribute("n") ?? "").Trim();
                        if (!string.Equals(n, "title", StringComparison.OrdinalIgnoreCase) && !ValidOneLevel(n))
                        {
                            error = "级别只能是 title 或 1 到 9";
                            return false;
                        }

                        var row = new TocLevelFormat();
                        XElement font = level.Element("font");
                        if (font != null && !ReadKeys(font, row.Font, character: true, out error))
                        {
                            return false;
                        }

                        XElement paragraph = level.Element("paragraph");
                        if (paragraph != null && !ReadKeys(paragraph, row.Paragraph, character: false, out error))
                        {
                            return false;
                        }

                        spec.LevelsFormat[string.Equals(n, "title", StringComparison.OrdinalIgnoreCase) ? "title" : n] = row;
                    }
                }
            }

            if (!spec.HasLevels)
            {
                spec.Levels = "1-3";
            }

            return true;
        }

        public static bool Merge(TocSpec current, TocSpec patch, out string error)
        {
            error = null;
            if (patch == null)
            {
                return false;
            }

            bool instruction = false;
            if (patch.HasLevels)
            {
                current.Levels = patch.Levels;
                current.HasLevels = true;
                instruction = true;
            }

            if (patch.HasStyles)
            {
                current.Styles = patch.Styles;
                current.HasStyles = patch.Styles.Count > 0;
                current.ClearStyles = patch.ClearStyles;
                instruction = true;
            }

            if (patch.HasRange)
            {
                current.From = patch.From;
                current.To = patch.To;
                current.HasRange = true;
                current.ClearRange = false;
                instruction = true;
            }
            else if (patch.ClearRange)
            {
                current.HasRange = false;
                current.ClearRange = true;
                current.From = null;
                current.To = null;
                instruction = true;
            }

            if (patch.HasPageNumbers)
            {
                current.PageNumbers = patch.PageNumbers;
                current.OmitPages = patch.OmitPages;
                current.HasPageNumbers = true;
                instruction = true;
            }

            if (patch.HasAlign)
            {
                current.AlignRight = patch.AlignRight;
                current.HasAlign = true;
                current.WriteTabs = patch.AlignRight;
                if (!patch.AlignRight && current.Separator == null && patch.Separator == null)
                {
                    current.Separator = " ";
                }

                instruction = !patch.AlignRight || current.Separator != null;
            }

            if (patch.Separator != null)
            {
                current.Separator = patch.Separator;
                instruction = true;
            }

            if (patch.HasLeader)
            {
                current.Leader = patch.Leader;
                current.HasLeader = true;
                current.WriteTabs = true;
            }

            if (patch.HasHyperlink)
            {
                current.Hyperlink = patch.Hyperlink;
                current.HasHyperlink = true;
                instruction = true;
            }

            if (patch.HasWebHide)
            {
                current.WebHide = patch.WebHide;
                current.HasWebHide = true;
                instruction = true;
            }

            if (patch.HasChapter)
            {
                current.ChapterSeq = patch.ClearChapter ? null : patch.ChapterSeq;
                current.ChapterSep = patch.ClearChapter ? null : patch.ChapterSep;
                current.HasChapter = !patch.ClearChapter;
                current.ClearChapter = patch.ClearChapter;
                instruction = true;
            }

            if (patch.HasKeepTabs)
            {
                current.KeepTabs = patch.KeepTabs;
                current.HasKeepTabs = true;
                instruction = true;
            }

            if (patch.HasKeepBreaks)
            {
                current.KeepBreaks = patch.KeepBreaks;
                current.HasKeepBreaks = true;
                instruction = true;
            }

            if (patch.HasTitle)
            {
                current.HasTitle = true;
                current.Title = patch.Title ?? "";
            }

            if (!patch.Reset)
            {
                foreach (KeyValuePair<string, TocLevelFormat> pair in patch.LevelsFormat)
                {
                    if (!current.LevelsFormat.TryGetValue(pair.Key, out TocLevelFormat mine))
                    {
                        mine = new TocLevelFormat();
                        current.LevelsFormat[pair.Key] = mine;
                    }

                    foreach (KeyValuePair<string, object> key in pair.Value.Font)
                    {
                        mine.Font[key.Key] = key.Value;
                    }

                    foreach (KeyValuePair<string, object> key in pair.Value.Paragraph)
                    {
                        mine.Paragraph[key.Key] = key.Value;
                    }
                }
            }

            current.Reset = patch.Reset;
            return instruction;
        }

        public static string BuildInstruction(TocSpec spec)
        {
            var parts = new List<string> { "TOC" };
            if (!string.IsNullOrEmpty(spec.Bookmark))
            {
                parts.Add("\\b \"" + spec.Bookmark + "\"");
            }

            parts.Add("\\o \"" + (string.IsNullOrEmpty(spec.Levels) ? "1-3" : spec.Levels) + "\"");
            if (spec.Hyperlink)
            {
                parts.Add("\\h");
            }

            if (spec.WebHide)
            {
                parts.Add("\\z");
            }

            parts.Add("\\u");
            if (!spec.PageNumbers)
            {
                parts.Add("\\n \"1-9\"");
            }
            else if (!string.IsNullOrEmpty(spec.OmitPages))
            {
                parts.Add("\\n \"" + spec.OmitPages + "\"");
            }

            if (spec.Styles.Count > 0)
            {
                var pairs = new List<string>();
                foreach (TocStyleRef style in spec.Styles)
                {
                    pairs.Add(style.Name);
                    pairs.Add(style.Level);
                }

                parts.Add("\\t \"" + string.Join(",", pairs) + "\"");
            }

            if (!spec.AlignRight && spec.Separator != null)
            {
                parts.Add("\\p \"" + spec.Separator + "\"");
            }

            if (!string.IsNullOrEmpty(spec.ChapterSeq))
            {
                parts.Add("\\s \"" + spec.ChapterSeq + "\"");
                if (!string.IsNullOrEmpty(spec.ChapterSep))
                {
                    parts.Add("\\d \"" + spec.ChapterSep + "\"");
                }
            }

            if (spec.KeepTabs)
            {
                parts.Add("\\w");
            }

            if (spec.KeepBreaks)
            {
                parts.Add("\\x");
            }

            return string.Join(" ", parts);
        }

        public static bool TryReadInstruction(string code, out TocSpec spec, out string error)
        {
            spec = TocSpec.InsertDefault();
            error = null;
            string text = code ?? "";
            if (!TocFieldStamp.IsTocInstruction(text))
            {
                error = "这不是目录域";
                return false;
            }

            spec.Levels = Quoted(text, "o") ?? "1-9";
            if (!ValidLevels(spec.Levels, out error))
            {
                spec.Levels = "1-9";
                error = null;
            }

            spec.HasLevels = true;
            spec.Bookmark = Quoted(text, "b");
            spec.Hyperlink = HasFlag(text, "h");
            spec.HasHyperlink = !spec.Hyperlink;
            spec.WebHide = HasFlag(text, "z");
            spec.HasWebHide = !spec.WebHide;
            spec.KeepTabs = HasFlag(text, "w");
            spec.HasKeepTabs = spec.KeepTabs;
            spec.KeepBreaks = HasFlag(text, "x");
            spec.HasKeepBreaks = spec.KeepBreaks;
            spec.OmitPages = Quoted(text, "n");
            if (spec.OmitPages == null && HasFlag(text, "n"))
            {
                spec.OmitPages = "1-9";
            }

            spec.PageNumbers = string.IsNullOrEmpty(spec.OmitPages);
            spec.HasPageNumbers = !spec.PageNumbers;
            string styles = Quoted(text, "t");
            if (!string.IsNullOrEmpty(styles))
            {
                string[] bits = styles.Split(',');
                for (int i = 0; i + 1 < bits.Length; i += 2)
                {
                    spec.Styles.Add(new TocStyleRef { Name = bits[i], Level = bits[i + 1] });
                }

                spec.HasStyles = spec.Styles.Count > 0;
            }

            spec.Separator = Quoted(text, "p");
            spec.AlignRight = spec.Separator == null;
            spec.HasAlign = !spec.AlignRight;
            spec.ChapterSeq = Quoted(text, "s");
            spec.ChapterSep = Quoted(text, "d");
            spec.HasChapter = !string.IsNullOrEmpty(spec.ChapterSeq);
            spec.Leader = "dot";
            return true;
        }

        public static string Write(TocSpec spec, string id, bool includeDefaults)
        {
            var toc = new XElement("toc");
            var content = new XElement("content");
            content.Add(new XElement("levels", spec.Levels ?? "1-3"));
            if (spec.Styles.Count > 0)
            {
                var styles = new XElement("styles");
                foreach (TocStyleRef style in spec.Styles)
                {
                    styles.Add(new XElement(
                        "style",
                        new XAttribute("name", style.Name ?? ""),
                        new XAttribute("level", style.Level ?? "")));
                }

                content.Add(styles);
            }

            if (spec.HasRange && !string.IsNullOrEmpty(spec.From) && !string.IsNullOrEmpty(spec.To))
            {
                content.Add(new XElement(
                    "range",
                    new XAttribute("from", spec.From),
                    new XAttribute("to", spec.To)));
            }

            toc.Add(content);
            var format = new XElement("format");
            bool any = false;
            if (includeDefaults || spec.HasPageNumbers)
            {
                if (!spec.PageNumbers)
                {
                    format.Add(new XElement("pageNumbers", new XAttribute("show", "false")));
                    any = true;
                }
                else if (!string.IsNullOrEmpty(spec.OmitPages))
                {
                    format.Add(new XElement(
                        "pageNumbers",
                        new XAttribute("levels", spec.OmitPages),
                        new XAttribute("show", "false")));
                    any = true;
                }
            }

            if (includeDefaults || spec.HasAlign)
            {
                format.Add(new XElement("align", spec.AlignRight ? "right" : "after"));
                any = true;
            }

            if (includeDefaults || spec.HasLeader)
            {
                format.Add(new XElement("leader", spec.Leader ?? "dot"));
                any = true;
            }

            if (spec.Separator != null && !spec.AlignRight)
            {
                format.Add(new XElement("separator", spec.Separator));
                any = true;
            }

            if (includeDefaults || spec.HasHyperlink)
            {
                format.Add(new XElement("hyperlink", spec.Hyperlink ? "true" : "false"));
                any = true;
            }

            if (includeDefaults || spec.HasWebHide)
            {
                format.Add(new XElement("webHidePageNumbers", spec.WebHide ? "true" : "false"));
                any = true;
            }

            if (spec.HasChapter && !string.IsNullOrEmpty(spec.ChapterSeq))
            {
                var chapter = new XElement("chapter", new XAttribute("seq", spec.ChapterSeq));
                if (!string.IsNullOrEmpty(spec.ChapterSep))
                {
                    chapter.Add(new XAttribute("separator", spec.ChapterSep));
                }

                format.Add(chapter);
                any = true;
            }

            if (spec.HasKeepTabs && spec.KeepTabs)
            {
                format.Add(new XElement("keepTabs", "true"));
                any = true;
            }

            if (spec.HasKeepBreaks && spec.KeepBreaks)
            {
                format.Add(new XElement("keepBreaks", "true"));
                any = true;
            }

            if (spec.HasTitle)
            {
                format.Add(new XElement("title", new XAttribute("text", spec.Title ?? "")));
                any = true;
            }

            foreach (KeyValuePair<string, TocLevelFormat> pair in spec.LevelsFormat)
            {
                if (pair.Value.Font.Count == 0 && pair.Value.Paragraph.Count == 0)
                {
                    continue;
                }

                var level = new XElement("level", new XAttribute("n", pair.Key));
                if (pair.Value.Font.Count > 0)
                {
                    level.Add(WriteKeys("font", pair.Value.Font));
                }

                if (pair.Value.Paragraph.Count > 0)
                {
                    level.Add(WriteKeys("paragraph", pair.Value.Paragraph));
                }

                format.Add(level);
                any = true;
            }

            if (any)
            {
                toc.Add(format);
            }

            return toc.ToString(SaveOptions.DisableFormatting);
        }

        private static XElement WriteKeys(string name, Dictionary<string, object> keys)
        {
            var element = new XElement(name);
            foreach (KeyValuePair<string, object> pair in keys)
            {
                if (pair.Key == "tab_stops")
                {
                    var tabs = new XElement("tab_stops");
                    if (pair.Value is IEnumerable<Dictionary<string, object>> rows)
                    {
                        foreach (Dictionary<string, object> row in rows)
                        {
                            var tab = new XElement("tab");
                            foreach (KeyValuePair<string, object> cell in row)
                            {
                                tab.Add(new XAttribute(cell.Key, Convert.ToString(cell.Value, CultureInfo.InvariantCulture) ?? ""));
                            }

                            tabs.Add(tab);
                        }
                    }

                    element.Add(tabs);
                    continue;
                }

                if (pair.Value is System.Collections.IDictionary nested)
                {
                    element.Add(DictElement(pair.Key, nested));
                    continue;
                }

                element.Add(new XAttribute(pair.Key, Convert.ToString(pair.Value, CultureInfo.InvariantCulture) ?? ""));
            }

            return element;
        }

        private static XElement DictElement(string name, System.Collections.IDictionary row)
        {
            var element = new XElement(name);
            foreach (System.Collections.DictionaryEntry entry in row)
            {
                string key = Convert.ToString(entry.Key) ?? "";
                if (entry.Value is System.Collections.IDictionary child)
                {
                    element.Add(DictElement(key, child));
                }
                else
                {
                    element.Add(new XAttribute(key, Convert.ToString(entry.Value, CultureInfo.InvariantCulture) ?? ""));
                }
            }

            return element;
        }

        private static bool ReadKeys(XElement element, Dictionary<string, object> target, bool character, out string error)
        {
            error = null;
            foreach (XAttribute attribute in element.Attributes())
            {
                if (character)
                {
                    if (Array.IndexOf(CharFormatCatalog.ApplyOrder, attribute.Name.LocalName) < 0)
                    {
                        error = "不认识的字符格式 " + attribute.Name.LocalName;
                        return false;
                    }
                }
                else if (!ParaFormatCatalog.IsKnownKey(attribute.Name.LocalName)
                    || attribute.Name.LocalName == "paragraph_style")
                {
                    error = "不认识的段落格式 " + attribute.Name.LocalName;
                    return false;
                }

                target[attribute.Name.LocalName] = attribute.Value;
            }

            foreach (XElement child in element.Elements())
            {
                if (!character && child.Name.LocalName == "tab_stops")
                {
                    continue;
                }

                if (character)
                {
                    if (Array.IndexOf(CharFormatCatalog.ApplyOrder, child.Name.LocalName) < 0)
                    {
                        error = "不认识的字符格式 " + child.Name.LocalName;
                        return false;
                    }
                }
                else if (!ParaFormatCatalog.IsKnownKey(child.Name.LocalName)
                    || child.Name.LocalName == "paragraph_style")
                {
                    error = "不认识的段落格式 " + child.Name.LocalName;
                    return false;
                }

                target[child.Name.LocalName] = NodeValue(child);
            }

            if (character)
            {
                return true;
            }

            XElement tabs = element.Element("tab_stops");
            if (tabs == null)
            {
                return true;
            }

            var rows = new List<Dictionary<string, object>>();
            foreach (XElement tab in tabs.Elements("tab"))
            {
                var row = new Dictionary<string, object>(StringComparer.Ordinal);
                foreach (XAttribute attribute in tab.Attributes())
                {
                    string value = attribute.Value;
                    if (attribute.Name.LocalName == "alignment")
                    {
                        value = MapTabAlign(value);
                    }
                    else if (attribute.Name.LocalName == "leader")
                    {
                        value = TocEditor.LeaderEnum(value);
                    }

                    row[attribute.Name.LocalName] = value;
                }

                rows.Add(row);
            }

            target["tab_stops"] = rows;
            return true;
        }

        private static object NodeValue(XElement element)
        {
            bool nested = element.HasAttributes || element.HasElements;
            if (!nested)
            {
                return element.Value ?? "";
            }

            var row = new Dictionary<string, object>(StringComparer.Ordinal);
            foreach (XAttribute attribute in element.Attributes())
            {
                row[attribute.Name.LocalName] = attribute.Value;
            }

            foreach (XElement child in element.Elements())
            {
                row[child.Name.LocalName] = NodeValue(child);
            }

            if (row.Count == 0)
            {
                return element.Value ?? "";
            }

            return row;
        }

        private static string MapTabAlign(string value)
        {
            switch ((value ?? "").Trim().ToLowerInvariant())
            {
                case "left":
                    return "wdAlignTabLeft";
                case "center":
                    return "wdAlignTabCenter";
                case "right":
                    return "wdAlignTabRight";
                case "decimal":
                    return "wdAlignTabDecimal";
                case "bar":
                    return "wdAlignTabBar";
                default:
                    return value;
            }
        }

        private static bool ReadBool(XElement element, bool current, out bool present)
        {
            present = element != null;
            if (element == null)
            {
                return current;
            }

            return IsTrue(element.Value);
        }

        private static bool IsTrue(string value)
        {
            return string.Equals((value ?? "").Trim(), "true", StringComparison.OrdinalIgnoreCase)
                || (value ?? "").Trim() == "1";
        }

        private static bool KnownLeader(string leader)
        {
            switch (leader)
            {
                case "none":
                case "dot":
                case "dash":
                case "line":
                case "heavy":
                case "middleDot":
                    return true;
                default:
                    return false;
            }
        }

        private static bool ValidOneLevel(string level)
        {
            return level != null && level.Length == 1 && level[0] >= '1' && level[0] <= '9';
        }

        private static bool ValidLevels(string levels, out string error)
        {
            error = null;
            Match match = Regex.Match(levels ?? "", @"^([1-9])-([1-9])$");
            if (!match.Success || match.Groups[1].Value[0] > match.Groups[2].Value[0])
            {
                error = "级别须为 1 到 9，且起点不大于终点";
                return false;
            }

            return true;
        }

        private static string Quoted(string code, string letter)
        {
            Match match = Regex.Match(
                code ?? "",
                @"\\" + letter + "\\s+\"([^\"]*)\"",
                RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
            return match.Success ? match.Groups[1].Value : null;
        }

        private static bool HasFlag(string code, string letter)
        {
            return Regex.IsMatch(
                code ?? "",
                @"\\" + letter + "(?![a-zA-Z])",
                RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
        }
    }
}
