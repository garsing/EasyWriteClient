using System;
using System.Collections.Generic;
using System.Diagnostics;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    /// <summary>
    /// 节级 PageSetup 与页眉页脚 Story 应用（从 DocumentRestoreHelper 抽取，供粗迁/恢复共用）。
    /// format_transfer 开启时对 PageSetup / Story COM 子步骤打点。
    /// </summary>
    public static class SectionLayoutHelper
    {
        private static bool FtDebugEnabled => EasyWriteDiagnostics.IsEnabled(DebugCategory.FormatTransfer);

        private static void FtLog(string message)
        {
            EasyWriteDiagnostics.Log(DebugCategory.FormatTransfer, $"[SectionLayout] {message}");
        }

        public static void CopyPageSetup(Word.PageSetup src, Word.PageSetup dst)
        {
            if (src == null || dst == null)
            {
                return;
            }

            try
            {
                dst.Orientation = src.Orientation;
                dst.PageWidth = src.PageWidth;
                dst.PageHeight = src.PageHeight;
                dst.TopMargin = src.TopMargin;
                dst.BottomMargin = src.BottomMargin;
                dst.LeftMargin = src.LeftMargin;
                dst.RightMargin = src.RightMargin;
                dst.HeaderDistance = src.HeaderDistance;
                dst.FooterDistance = src.FooterDistance;
                dst.DifferentFirstPageHeaderFooter = src.DifferentFirstPageHeaderFooter;
                dst.OddAndEvenPagesHeaderFooter = src.OddAndEvenPagesHeaderFooter;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[SectionLayoutHelper] CopyPageSetup warning: {ex.Message}");
            }
        }

        public static void ApplyPageSetupFromDict(
            Word.PageSetup pageSetup,
            Dictionary<string, object> setupDict,
            bool skipIfNearlyEqual = true)
        {
            if (pageSetup == null || setupDict == null || setupDict.Count == 0)
            {
                return;
            }

            var totalSw = FtDebugEnabled ? Stopwatch.StartNew() : null;
            FtLog($"ApplyPageSetupFromDict ENTER keys={setupDict.Count} skipIfNearlyEqual={skipIfNearlyEqual}");

            try
            {
                string orientation = GetString(setupDict, "orientation");
                if (string.Equals(orientation, "landscape", StringComparison.OrdinalIgnoreCase))
                {
                    FtLog("set Orientation=landscape…");
                    pageSetup.Orientation = Word.WdOrientation.wdOrientLandscape;
                }
                else if (!string.IsNullOrEmpty(orientation))
                {
                    FtLog("set Orientation=portrait…");
                    pageSetup.Orientation = Word.WdOrientation.wdOrientPortrait;
                }

                TrySetFloat(pageSetup, "PageWidth", setupDict, "page_width", skipIfNearlyEqual);
                TrySetFloat(pageSetup, "PageHeight", setupDict, "page_height", skipIfNearlyEqual);
                TrySetFloat(pageSetup, "TopMargin", setupDict, "top_margin", skipIfNearlyEqual);
                TrySetFloat(pageSetup, "BottomMargin", setupDict, "bottom_margin", skipIfNearlyEqual);
                TrySetFloat(pageSetup, "LeftMargin", setupDict, "left_margin", skipIfNearlyEqual);
                TrySetFloat(pageSetup, "RightMargin", setupDict, "right_margin", skipIfNearlyEqual);
                TrySetFloat(pageSetup, "HeaderDistance", setupDict, "header_distance", skipIfNearlyEqual);
                TrySetFloat(pageSetup, "FooterDistance", setupDict, "footer_distance", skipIfNearlyEqual);

                if (setupDict.TryGetValue("different_first_page_header_footer", out object diffFirst))
                {
                    FtLog($"set DifferentFirstPageHeaderFooter={ToBool(diffFirst)}…");
                    pageSetup.DifferentFirstPageHeaderFooter = ToWordIntBool(ToBool(diffFirst));
                }

                if (setupDict.TryGetValue("odd_and_even_pages_header_footer", out object oddEven))
                {
                    FtLog($"set OddAndEvenPagesHeaderFooter={ToBool(oddEven)}…");
                    pageSetup.OddAndEvenPagesHeaderFooter = ToWordIntBool(ToBool(oddEven));
                }
            }
            catch (Exception ex)
            {
                FtLog($"ApplyPageSetupFromDict exception: {ex.Message}");
                System.Diagnostics.Debug.WriteLine($"[SectionLayoutHelper] ApplyPageSetupFromDict warning: {ex.Message}");
            }

            if (totalSw != null)
            {
                totalSw.Stop();
                FtLog($"ApplyPageSetupFromDict EXIT elapsed_ms={totalSw.ElapsedMilliseconds}");
            }
        }

        public static void ApplyCharFormatToRange(Word.Range range, Dictionary<string, object> charFormat)
        {
            if (range == null || charFormat == null || charFormat.Count == 0)
            {
                return;
            }

            Dictionary<string, object> snapshot = FormatInheritHelper.BuildSnapshotFromCharFormat(charFormat);
            if (snapshot == null || snapshot.Count == 0)
            {
                return;
            }

            FormatInheritHelper.ApplySnapshot(range, snapshot, charFormatOnly: true);
        }

        public static bool StoryHasContent(Word.Range range)
        {
            if (range == null)
            {
                return false;
            }

            var sw = FtDebugEnabled ? Stopwatch.StartNew() : null;
            try
            {
                FtLog("StoryHasContent read Text…");
                string text = (range.Text ?? "").Replace("\r", "").Replace("\a", "").Trim();
                if (!string.IsNullOrEmpty(text))
                {
                    if (sw != null)
                    {
                        sw.Stop();
                        FtLog($"StoryHasContent=true via text len={text.Length} elapsed_ms={sw.ElapsedMilliseconds}");
                    }

                    return true;
                }

                FtLog("StoryHasContent check Fields…");
                if (range.Fields != null && range.Fields.Count > 0)
                {
                    if (sw != null)
                    {
                        sw.Stop();
                        FtLog($"StoryHasContent=true via fields elapsed_ms={sw.ElapsedMilliseconds}");
                    }

                    return true;
                }

                FtLog("StoryHasContent check InlineShapes…");
                if (range.InlineShapes != null && range.InlineShapes.Count > 0)
                {
                    if (sw != null)
                    {
                        sw.Stop();
                        FtLog($"StoryHasContent=true via shapes elapsed_ms={sw.ElapsedMilliseconds}");
                    }

                    return true;
                }
            }
            catch (Exception ex)
            {
                FtLog($"StoryHasContent warning: {ex.Message}");
                System.Diagnostics.Debug.WriteLine($"[SectionLayoutHelper] StoryHasContent warning: {ex.Message}");
            }

            if (sw != null)
            {
                sw.Stop();
                FtLog($"StoryHasContent=false elapsed_ms={sw.ElapsedMilliseconds}");
            }

            return false;
        }

        public static void EnsureUnlinked(Word.Section section, Word.WdHeaderFooterIndex index)
        {
            if (section == null)
            {
                return;
            }

            try
            {
                FtLog($"EnsureUnlinked index={index} header…");
                var sw = FtDebugEnabled ? Stopwatch.StartNew() : null;
                section.Headers[index].LinkToPrevious = false;
                if (sw != null)
                {
                    sw.Stop();
                    FtLog($"EnsureUnlinked header done elapsed_ms={sw.ElapsedMilliseconds}");
                }

                FtLog($"EnsureUnlinked index={index} footer…");
                sw = FtDebugEnabled ? Stopwatch.StartNew() : null;
                section.Footers[index].LinkToPrevious = false;
                if (sw != null)
                {
                    sw.Stop();
                    FtLog($"EnsureUnlinked footer done elapsed_ms={sw.ElapsedMilliseconds}");
                }
            }
            catch (Exception ex)
            {
                FtLog($"EnsureUnlinked warning: {ex.Message}");
                System.Diagnostics.Debug.WriteLine($"[SectionLayoutHelper] EnsureUnlinked warning: {ex.Message}");
            }
        }

        public static void ClearStoryRange(Word.Range range)
        {
            if (range == null)
            {
                return;
            }

            try
            {
                int start = range.Start;
                int end = range.End;
                FtLog($"ClearStoryRange start={start} end={end}…");
                var sw = FtDebugEnabled ? Stopwatch.StartNew() : null;
                if (start < end)
                {
                    range.Delete();
                }

                if (sw != null)
                {
                    sw.Stop();
                    FtLog($"ClearStoryRange done deleted={start < end} elapsed_ms={sw.ElapsedMilliseconds}");
                }
            }
            catch (Exception ex)
            {
                FtLog($"ClearStoryRange warning: {ex.Message}");
                System.Diagnostics.Debug.WriteLine($"[SectionLayoutHelper] ClearStoryRange warning: {ex.Message}");
            }
        }

        /// <summary>
        /// 页眉/页脚 Story 须保留至少一个段落；用于 restore 前清空，避免 FormattedText 叠加多余 ¶。
        /// </summary>
        public static void ClearStoryKeepingOneParagraph(Word.Range storyRange)
        {
            if (storyRange == null)
            {
                return;
            }

            try
            {
                while (storyRange.Paragraphs.Count > 1)
                {
                    storyRange.Paragraphs[storyRange.Paragraphs.Count].Range.Delete();
                }

                storyRange.Paragraphs[1].Range.Text = "";
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine(
                    $"[SectionLayoutHelper] ClearStoryKeepingOneParagraph warning: {ex.Message}");
            }
        }

        /// <summary>
        /// 复制页眉/页脚 Story：先清空目标，再写入源内容（排除源 Story 末尾段落符，避免每次 restore 多一行）。
        /// </summary>
        public static void CopyStoryFormattedText(Word.Range sourceRange, Word.Range targetRange)
        {
            if (sourceRange == null || targetRange == null)
            {
                return;
            }

            ClearStoryKeepingOneParagraph(targetRange);

            Word.Range srcDup = sourceRange.Duplicate;
            if (srcDup.End > srcDup.Start)
            {
                srcDup.MoveEnd(Word.WdUnits.wdCharacter, -1);
            }

            if (srcDup.End <= srcDup.Start)
            {
                return;
            }

            try
            {
                Word.Range dstInsert = targetRange.Duplicate;
                dstInsert.Collapse(Word.WdCollapseDirection.wdCollapseStart);
                dstInsert.FormattedText = srcDup.FormattedText;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine(
                    $"[SectionLayoutHelper] CopyStoryFormattedText warning: {ex.Message}");
            }
        }

        public static string ApplyHeaderFooterStory(
            Word.Range headerRange,
            Word.Range footerRange,
            Dictionary<string, object> headerStory,
            Dictionary<string, object> footerStory,
            string contentMode)
        {
            string appliedMode = "none";
            if (headerStory != null && headerStory.Count > 0)
            {
                appliedMode = ApplySingleStory(headerRange, headerStory, contentMode);
            }

            if (footerStory != null && footerStory.Count > 0)
            {
                string footerApplied = ApplySingleStory(footerRange, footerStory, contentMode);
                if (appliedMode == "none")
                {
                    appliedMode = footerApplied;
                }
            }

            return appliedMode;
        }

        public static string ResolveStoryApplyMode(
            Word.Range range,
            Dictionary<string, object> storyDict,
            string contentMode)
        {
            string mode = (contentMode ?? "auto").Trim().ToLowerInvariant();
            if (mode == "format_only")
            {
                return "format_only";
            }

            if (mode == "full")
            {
                return "full";
            }

            return StoryHasContent(range) ? "format_only" : "full";
        }

        public static string ApplySingleStory(
            Word.Range range,
            Dictionary<string, object> storyDict,
            string contentMode,
            string debugTag = null)
        {
            if (range == null || storyDict == null || storyDict.Count == 0)
            {
                return "none";
            }

            string tag = string.IsNullOrEmpty(debugTag) ? "story" : debugTag;
            var totalSw = FtDebugEnabled ? Stopwatch.StartNew() : null;
            FtLog($"ApplySingleStory ENTER {tag} content_mode={contentMode ?? "auto"}");

            FtLog($"ApplySingleStory {tag} ResolveStoryApplyMode / StoryHasContent…");
            var resolveSw = FtDebugEnabled ? Stopwatch.StartNew() : null;
            string applyMode = ResolveStoryApplyMode(range, storyDict, contentMode);
            if (resolveSw != null)
            {
                resolveSw.Stop();
                FtLog($"ApplySingleStory {tag} resolve mode={applyMode} elapsed_ms={resolveSw.ElapsedMilliseconds}");
            }
            else
            {
                FtLog($"ApplySingleStory {tag} resolve mode={applyMode}");
            }

            Dictionary<string, object> charFormat = GetDict(storyDict, "char_format");

            if (applyMode == "full")
            {
                FtLog($"ApplySingleStory {tag} ClearStoryRange…");
                ClearStoryRange(range);

                FtLog($"ApplySingleStory {tag} WriteContentSegments…");
                var writeSw = FtDebugEnabled ? Stopwatch.StartNew() : null;
                WriteContentSegments(range, storyDict);
                if (writeSw != null)
                {
                    writeSw.Stop();
                    FtLog($"ApplySingleStory {tag} WriteContentSegments done elapsed_ms={writeSw.ElapsedMilliseconds}");
                }

                if (charFormat != null && charFormat.Count > 0)
                {
                    FtLog($"ApplySingleStory {tag} ApplyCharFormatToRange…");
                    var fmtSw = FtDebugEnabled ? Stopwatch.StartNew() : null;
                    ApplyCharFormatToRange(range, charFormat);
                    if (fmtSw != null)
                    {
                        fmtSw.Stop();
                        FtLog($"ApplySingleStory {tag} ApplyCharFormatToRange done elapsed_ms={fmtSw.ElapsedMilliseconds}");
                    }
                }

                if (totalSw != null)
                {
                    totalSw.Stop();
                    FtLog($"ApplySingleStory EXIT {tag} mode=full elapsed_ms={totalSw.ElapsedMilliseconds}");
                }

                return "full";
            }

            if (charFormat != null && charFormat.Count > 0)
            {
                FtLog($"ApplySingleStory {tag} ApplyCharFormatToRange (format_only)…");
                var fmtSw = FtDebugEnabled ? Stopwatch.StartNew() : null;
                ApplyCharFormatToRange(range, charFormat);
                if (fmtSw != null)
                {
                    fmtSw.Stop();
                    FtLog($"ApplySingleStory {tag} ApplyCharFormatToRange done elapsed_ms={fmtSw.ElapsedMilliseconds}");
                }

                if (totalSw != null)
                {
                    totalSw.Stop();
                    FtLog($"ApplySingleStory EXIT {tag} mode=format_only elapsed_ms={totalSw.ElapsedMilliseconds}");
                }

                return "format_only";
            }

            if (totalSw != null)
            {
                totalSw.Stop();
                FtLog($"ApplySingleStory EXIT {tag} mode=none elapsed_ms={totalSw.ElapsedMilliseconds}");
            }

            return "none";
        }

        public static void WriteContentSegments(Word.Range range, Dictionary<string, object> storyDict)
        {
            if (range == null || storyDict == null)
            {
                return;
            }

            object segmentsObj = GetObject(storyDict, "content_segments");
            if (segmentsObj is List<object> segmentList && segmentList.Count > 0)
            {
                FtLog($"WriteContentSegments via List count={segmentList.Count}");
                WriteSegmentList(range, segmentList);
                return;
            }

            if (segmentsObj is Newtonsoft.Json.Linq.JArray jArray && jArray.Count > 0)
            {
                FtLog($"WriteContentSegments via JArray count={jArray.Count}");
                WriteSegmentList(range, jArray.ToObject<List<object>>());
                return;
            }

            string plain = GetString(storyDict, "content_plain");
            if (!string.IsNullOrEmpty(plain))
            {
                FtLog($"WriteContentSegments plain len={plain.Length}…");
                var sw = FtDebugEnabled ? Stopwatch.StartNew() : null;
                range.Text = plain;
                if (sw != null)
                {
                    sw.Stop();
                    FtLog($"WriteContentSegments plain done elapsed_ms={sw.ElapsedMilliseconds}");
                }
            }
            else
            {
                FtLog("WriteContentSegments empty (no segments / plain)");
            }
        }

        private static void WriteSegmentList(Word.Range range, List<object> segments)
        {
            if (range == null || segments == null)
            {
                return;
            }

            Word.Range insertRange = range.Duplicate;
            insertRange.Collapse(Word.WdCollapseDirection.wdCollapseStart);

            foreach (object segObj in segments)
            {
                Dictionary<string, object> seg = NormalizeDict(segObj);
                if (seg == null || seg.Count == 0)
                {
                    continue;
                }

                string segType = GetString(seg, "type");
                if (segType == "text")
                {
                    string value = GetString(seg, "value");
                    if (!string.IsNullOrEmpty(value))
                    {
                        insertRange.InsertAfter(value);
                        insertRange.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                    }
                }
                else if (segType == "field")
                {
                    InsertFieldInstruction(insertRange, GetString(seg, "instruction"));
                    insertRange.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                }
            }
        }

        private static void InsertFieldInstruction(Word.Range range, string instruction)
        {
            if (range == null || string.IsNullOrWhiteSpace(instruction))
            {
                return;
            }

            string normalized = instruction.Trim().ToUpperInvariant();
            try
            {
                if (normalized.Contains("NUMPAGES"))
                {
                    range.Fields.Add(range, Word.WdFieldType.wdFieldNumPages);
                    return;
                }

                if (normalized.Contains("PAGE"))
                {
                    range.Fields.Add(range, Word.WdFieldType.wdFieldPage);
                    return;
                }

                object fieldType = Word.WdFieldType.wdFieldEmpty;
                range.Fields.Add(range, ref fieldType, instruction.Trim(), false);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[SectionLayoutHelper] InsertFieldInstruction warning: {ex.Message}");
            }
        }

        private static void TrySetFloat(
            Word.PageSetup pageSetup,
            string propertyName,
            Dictionary<string, object> setupDict,
            string jsonKey,
            bool skipIfNearlyEqual = true)
        {
            if (!setupDict.TryGetValue(jsonKey, out object value) || value == null)
            {
                return;
            }

            if (!TryToFloat(value, out float f))
            {
                return;
            }

            try
            {
                if (skipIfNearlyEqual && TryGetCurrentFloat(pageSetup, propertyName, out float current)
                    && Math.Abs(current - f) < 0.5f)
                {
                    FtLog($"set {propertyName} SKIPPED near-equal current={current} target={f}");
                    return;
                }

                FtLog($"set {propertyName}={f}…");
                var sw = FtDebugEnabled ? Stopwatch.StartNew() : null;
                switch (propertyName)
                {
                    case "PageWidth":
                        pageSetup.PageWidth = f;
                        break;
                    case "PageHeight":
                        pageSetup.PageHeight = f;
                        break;
                    case "TopMargin":
                        pageSetup.TopMargin = f;
                        break;
                    case "BottomMargin":
                        pageSetup.BottomMargin = f;
                        break;
                    case "LeftMargin":
                        pageSetup.LeftMargin = f;
                        break;
                    case "RightMargin":
                        pageSetup.RightMargin = f;
                        break;
                    case "HeaderDistance":
                        pageSetup.HeaderDistance = f;
                        break;
                    case "FooterDistance":
                        pageSetup.FooterDistance = f;
                        break;
                }

                if (sw != null)
                {
                    sw.Stop();
                    FtLog($"set {propertyName} done elapsed_ms={sw.ElapsedMilliseconds}");
                }
            }
            catch (Exception ex)
            {
                FtLog($"TrySetFloat {propertyName} warning: {ex.Message}");
                System.Diagnostics.Debug.WriteLine(
                    $"[SectionLayoutHelper] TrySetFloat {propertyName} warning: {ex.Message}");
            }
        }

        private static bool TryGetCurrentFloat(Word.PageSetup pageSetup, string propertyName, out float value)
        {
            value = 0f;
            try
            {
                switch (propertyName)
                {
                    case "PageWidth":
                        value = pageSetup.PageWidth;
                        return true;
                    case "PageHeight":
                        value = pageSetup.PageHeight;
                        return true;
                    case "TopMargin":
                        value = pageSetup.TopMargin;
                        return true;
                    case "BottomMargin":
                        value = pageSetup.BottomMargin;
                        return true;
                    case "LeftMargin":
                        value = pageSetup.LeftMargin;
                        return true;
                    case "RightMargin":
                        value = pageSetup.RightMargin;
                        return true;
                    case "HeaderDistance":
                        value = pageSetup.HeaderDistance;
                        return true;
                    case "FooterDistance":
                        value = pageSetup.FooterDistance;
                        return true;
                    default:
                        return false;
                }
            }
            catch
            {
                return false;
            }
        }

        private static bool TryToFloat(object value, out float result)
        {
            result = 0f;
            if (value == null)
            {
                return false;
            }

            if (value is float f)
            {
                result = f;
                return true;
            }

            if (value is double d)
            {
                result = (float)d;
                return true;
            }

            if (value is long l)
            {
                result = l;
                return true;
            }

            if (value is int i)
            {
                result = i;
                return true;
            }

            return float.TryParse(value.ToString(), out result);
        }

        private static bool ToBool(object value)
        {
            if (value is bool b)
            {
                return b;
            }

            if (value == null)
            {
                return false;
            }

            return string.Equals(value.ToString(), "true", StringComparison.OrdinalIgnoreCase)
                || value.ToString() == "1";
        }

        /// <summary>Word COM 布尔属性为 int：0=false，-1=true。</summary>
        public static int ToWordIntBool(bool value)
        {
            return value ? -1 : 0;
        }

        private static Dictionary<string, object> GetDict(Dictionary<string, object> dict, string key)
        {
            if (dict == null || !dict.TryGetValue(key, out object o))
            {
                return null;
            }

            return NormalizeDict(o);
        }

        private static object GetObject(Dictionary<string, object> dict, string key)
        {
            if (dict == null || !dict.TryGetValue(key, out object o))
            {
                return null;
            }

            return o;
        }

        private static string GetString(Dictionary<string, object> dict, string key)
        {
            if (dict == null || !dict.TryGetValue(key, out object o) || o == null)
            {
                return "";
            }

            return o.ToString() ?? "";
        }

        private static Dictionary<string, object> NormalizeDict(object obj)
        {
            if (obj == null)
            {
                return null;
            }

            if (obj is Dictionary<string, object> d)
            {
                return d;
            }

            if (obj is Newtonsoft.Json.Linq.JObject jo)
            {
                return jo.ToObject<Dictionary<string, object>>();
            }

            return null;
        }
    }
}
