using System;
using System.Collections.Generic;
using System.Diagnostics;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    public sealed class PageLayoutApplyResult
    {
        public bool Success { get; set; }
        public int SectionsProcessed { get; set; }
        public int HeadersFootersUpdated { get; set; }
        public int HfContentFullCount { get; set; }
        public int HfContentFormatOnlyCount { get; set; }
        public List<string> HfContentModes { get; set; } = new List<string>();
        public List<string> Warnings { get; set; } = new List<string>();
        public bool PageSetupApplied { get; set; }
        public bool PrinterSwitched { get; set; }
        public bool PrinterReady { get; set; } = true;
    }

    public sealed class PageLayoutApplyOptions
    {
        public bool ApplyPageSetup { get; set; } = true;
        public bool ApplyHeadersFooters { get; set; } = true;
        public int? SectionIndex { get; set; }
    }

    /// <summary>
    /// 将 KB section_page_layout.mode 应用到当前稿每一节（粗迁先页后段）。
    /// PageSetup 经 DocumentPageSetupApplier（PDF 保障）；Story 复用现网逻辑。
    /// </summary>
    public static class PageLayoutApplyHelper
    {
        private static readonly string[] StoryKeys = { "primary", "first", "even" };

        private static bool FtDebugEnabled => EasyWriteDiagnostics.IsEnabled(DebugCategory.FormatTransfer);

        private static void FtLog(string message)
        {
            EasyWriteDiagnostics.Log(DebugCategory.FormatTransfer, $"[PageLayout] {message}");
        }

        public static PageLayoutApplyResult ApplyModeToDocument(
            Word.Document doc,
            Dictionary<string, object> mode,
            string contentMode = "auto",
            PageLayoutApplyOptions options = null)
        {
            var result = new PageLayoutApplyResult();
            if (options == null)
            {
                options = new PageLayoutApplyOptions();
            }

            if (doc == null || mode == null || mode.Count == 0)
            {
                result.Warnings.Add("mode 为空，跳过页布局");
                FtLog("ABORT mode empty");
                return result;
            }

            if (!options.ApplyPageSetup && !options.ApplyHeadersFooters)
            {
                result.Warnings.Add("apply_page_setup 与 apply_headers_footers 均为 false");
                FtLog("ABORT both apply flags false");
                return result;
            }

            Dictionary<string, object> pageSetup = GetDict(mode, "page_setup");
            Dictionary<string, object> headers = GetDict(mode, "headers");
            Dictionary<string, object> footers = GetDict(mode, "footers");
            bool enableFirst = pageSetup != null
                && GetBool(pageSetup, "different_first_page_header_footer");
            bool enableEven = pageSetup != null
                && GetBool(pageSetup, "odd_and_even_pages_header_footer");

            int sectionCount;
            try
            {
                sectionCount = doc.Sections.Count;
            }
            catch (Exception ex)
            {
                FtLog($"ABORT doc.Sections.Count failed: {ex.Message}");
                result.Warnings.Add($"doc.Sections.Count failed: {ex.Message}");
                return result;
            }

            FtLog(
                $"BEGIN sections={sectionCount} content_mode={contentMode ?? "auto"} " +
                $"applyPs={options.ApplyPageSetup} applyHf={options.ApplyHeadersFooters} " +
                $"sectionIndex={(options.SectionIndex.HasValue ? options.SectionIndex.Value.ToString() : "all")} " +
                $"enableFirst={enableFirst} enableEven={enableEven} " +
                $"has_page_setup={pageSetup != null && pageSetup.Count > 0} " +
                $"has_headers={headers != null && headers.Count > 0} " +
                $"has_footers={footers != null && footers.Count > 0}");

            bool printerReadyForSync = true;
            string previousPrinter = null;
            Word.Application app = null;
            bool prevScreenUpdating = true;
            try
            {
                app = doc.Application;
                if (app != null)
                {
                    try
                    {
                        previousPrinter = app.ActivePrinter;
                    }
                    catch
                    {
                        previousPrinter = null;
                    }

                    prevScreenUpdating = app.ScreenUpdating;
                    app.ScreenUpdating = false;
                    FtLog($"ScreenUpdating false (was {prevScreenUpdating})");
                }
            }
            catch (Exception ex)
            {
                FtLog($"ScreenUpdating disable skipped: {ex.Message}");
                app = null;
            }

            // SyncStorySwitches 也会写 PageSetup：整段页布局期间保持 PDF，结束再 Restore
            if (options.ApplyPageSetup && pageSetup != null && pageSetup.Count > 0)
            {
                FtLog("DocumentPageSetupApplier.ApplyPageSetupToDocument…");
                PageSetupApplyResult psResult = DocumentPageSetupApplier.ApplyPageSetupToDocument(
                    doc,
                    pageSetup,
                    new PageSetupApplyOptions
                    {
                        SectionIndex = options.SectionIndex,
                        SkipIfNearlyEqual = true,
                        RestorePrinterAfter = false,
                    });
                result.PrinterReady = psResult.PrinterReady;
                result.PrinterSwitched = psResult.PrinterSwitched;
                result.PageSetupApplied = psResult.SectionsSucceeded > 0;
                printerReadyForSync = psResult.PrinterReady;
                if (psResult.Warnings != null && psResult.Warnings.Count > 0)
                {
                    result.Warnings.AddRange(psResult.Warnings);
                }

                FtLog(
                    $"DocumentPageSetupApplier done ready={psResult.PrinterReady} " +
                    $"succeeded={psResult.SectionsSucceeded}/{psResult.SectionsAttempted}");
            }
            else if (options.ApplyPageSetup)
            {
                FtLog("ApplyPageSetup skipped (empty page_setup)");
            }
            else if (options.ApplyHeadersFooters)
            {
                printerReadyForSync = WordActivePrinterHelper.EnsurePreferredActivePrinter(app);
                result.PrinterReady = printerReadyForSync;
                if (!printerReadyForSync)
                {
                    result.Warnings.Add("page_setup_skipped_no_printer");
                    FtLog("HF-only: Ensure printer failed; SyncStorySwitches will be skipped");
                }
            }

            int start = 1;
            int end = sectionCount;
            if (options.SectionIndex.HasValue)
            {
                int idx = options.SectionIndex.Value;
                if (idx < 1 || idx > sectionCount)
                {
                    result.Warnings.Add($"section_index={idx} 超出范围 1..{sectionCount}");
                    result.Success = result.PageSetupApplied || result.HeadersFootersUpdated > 0;
                    return result;
                }

                start = idx;
                end = idx;
            }

            var totalSw = FtDebugEnabled ? Stopwatch.StartNew() : null;
            try
            {
                for (int i = start; i <= end; i++)
                {
                    FtLog($"section[{i}/{sectionCount}] ENTER");
                    var secSw = FtDebugEnabled ? Stopwatch.StartNew() : null;

                    Word.Section section;
                    try
                    {
                        section = doc.Sections[i];
                    }
                    catch (Exception ex)
                    {
                        FtLog($"section[{i}/{sectionCount}] get Section failed: {ex.Message}");
                        result.Warnings.Add($"section {i} get failed: {ex.Message}");
                        continue;
                    }

                    try
                    {
                        bool needSync = options.ApplyPageSetup || options.ApplyHeadersFooters;
                        if (needSync && printerReadyForSync)
                        {
                            FtLog($"section[{i}/{sectionCount}] SyncStorySwitches…");
                            var syncSw = FtDebugEnabled ? Stopwatch.StartNew() : null;
                            SyncStorySwitches(section, enableFirst, enableEven, result.Warnings, i, sectionCount);
                            if (syncSw != null)
                            {
                                syncSw.Stop();
                                FtLog($"section[{i}/{sectionCount}] SyncStorySwitches done elapsed_ms={syncSw.ElapsedMilliseconds}");
                            }
                        }
                        else if (needSync)
                        {
                            FtLog($"section[{i}/{sectionCount}] SyncStorySwitches SKIPPED (no printer)");
                        }

                        if (options.ApplyHeadersFooters)
                        {
                            FtLog($"section[{i}/{sectionCount}] ApplyStoriesForSection…");
                            var storiesSw = FtDebugEnabled ? Stopwatch.StartNew() : null;
                            ApplyStoriesForSection(
                                section, headers, footers, enableFirst, enableEven, contentMode, result, i, sectionCount);
                            if (storiesSw != null)
                            {
                                storiesSw.Stop();
                                FtLog($"section[{i}/{sectionCount}] ApplyStoriesForSection done elapsed_ms={storiesSw.ElapsedMilliseconds}");
                            }
                        }

                        result.SectionsProcessed++;
                    }
                    catch (Exception ex)
                    {
                        FtLog($"section[{i}/{sectionCount}] EXCEPTION: {ex.GetType().Name}: {ex.Message}");
                        result.Warnings.Add($"section {i} apply failed: {ex.Message}");
                    }

                    if (secSw != null)
                    {
                        secSw.Stop();
                        FtLog($"section[{i}/{sectionCount}] EXIT elapsed_ms={secSw.ElapsedMilliseconds}");
                    }
                    else
                    {
                        FtLog($"section[{i}/{sectionCount}] EXIT");
                    }
                }
            }
            finally
            {
                if (app != null)
                {
                    try
                    {
                        app.ScreenUpdating = prevScreenUpdating;
                        FtLog($"ScreenUpdating restored={prevScreenUpdating}");
                    }
                    catch (Exception ex)
                    {
                        FtLog($"ScreenUpdating restore failed: {ex.Message}");
                    }

                    WordActivePrinterHelper.RestoreActivePrinter(app, previousPrinter);
                }
            }

            bool anyWork =
                result.PageSetupApplied
                || result.HeadersFootersUpdated > 0
                || (options.ApplyHeadersFooters && result.SectionsProcessed > 0);
            result.Success = anyWork || result.SectionsProcessed > 0;
            if (totalSw != null)
            {
                totalSw.Stop();
                FtLog(
                    $"END success={result.Success} sections={result.SectionsProcessed} " +
                    $"ps_applied={result.PageSetupApplied} hf_updated={result.HeadersFootersUpdated} " +
                    $"total_elapsed_ms={totalSw.ElapsedMilliseconds}");
            }
            else
            {
                FtLog(
                    $"END success={result.Success} sections={result.SectionsProcessed} " +
                    $"ps_applied={result.PageSetupApplied}");
            }

            return result;
        }

        private static void SyncStorySwitches(
            Word.Section section,
            bool enableFirst,
            bool enableEven,
            List<string> warnings,
            int sectionIndex = 0,
            int sectionCount = 0)
        {
            string tag = sectionIndex > 0 ? $"section[{sectionIndex}/{sectionCount}]" : "section";
            try
            {
                FtLog($"{tag} set DifferentFirstPageHeaderFooter={enableFirst}");
                section.PageSetup.DifferentFirstPageHeaderFooter = SectionLayoutHelper.ToWordIntBool(enableFirst);
                FtLog($"{tag} set OddAndEvenPagesHeaderFooter={enableEven}");
                section.PageSetup.OddAndEvenPagesHeaderFooter = SectionLayoutHelper.ToWordIntBool(enableEven);

                if (!enableFirst)
                {
                    FtLog($"{tag} ClearStory first header…");
                    var sw = FtDebugEnabled ? Stopwatch.StartNew() : null;
                    ClearStory(section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterFirstPage].Range);
                    if (sw != null)
                    {
                        sw.Stop();
                        FtLog($"{tag} ClearStory first header done elapsed_ms={sw.ElapsedMilliseconds}");
                    }

                    FtLog($"{tag} ClearStory first footer…");
                    sw = FtDebugEnabled ? Stopwatch.StartNew() : null;
                    ClearStory(section.Footers[Word.WdHeaderFooterIndex.wdHeaderFooterFirstPage].Range);
                    if (sw != null)
                    {
                        sw.Stop();
                        FtLog($"{tag} ClearStory first footer done elapsed_ms={sw.ElapsedMilliseconds}");
                    }
                }

                if (!enableEven)
                {
                    FtLog($"{tag} ClearStory even header…");
                    var sw = FtDebugEnabled ? Stopwatch.StartNew() : null;
                    ClearStory(section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterEvenPages].Range);
                    if (sw != null)
                    {
                        sw.Stop();
                        FtLog($"{tag} ClearStory even header done elapsed_ms={sw.ElapsedMilliseconds}");
                    }

                    FtLog($"{tag} ClearStory even footer…");
                    sw = FtDebugEnabled ? Stopwatch.StartNew() : null;
                    ClearStory(section.Footers[Word.WdHeaderFooterIndex.wdHeaderFooterEvenPages].Range);
                    if (sw != null)
                    {
                        sw.Stop();
                        FtLog($"{tag} ClearStory even footer done elapsed_ms={sw.ElapsedMilliseconds}");
                    }
                }
            }
            catch (Exception ex)
            {
                FtLog($"{tag} SyncStorySwitches warning: {ex.Message}");
                warnings.Add($"sync story switches warning: {ex.Message}");
            }
        }

        private static void ApplyStoriesForSection(
            Word.Section section,
            Dictionary<string, object> headers,
            Dictionary<string, object> footers,
            bool enableFirst,
            bool enableEven,
            string contentMode,
            PageLayoutApplyResult result,
            int sectionIndex = 0,
            int sectionCount = 0)
        {
            string tag = sectionIndex > 0 ? $"section[{sectionIndex}/{sectionCount}]" : "section";

            foreach (string storyKey in StoryKeys)
            {
                if (storyKey == "first" && !enableFirst)
                {
                    FtLog($"{tag} story={storyKey} skipped (enableFirst=false)");
                    continue;
                }

                if (storyKey == "even" && !enableEven)
                {
                    FtLog($"{tag} story={storyKey} skipped (enableEven=false)");
                    continue;
                }

                Word.WdHeaderFooterIndex index = MapStoryKey(storyKey);
                FtLog($"{tag} story={storyKey} EnsureUnlinked…");
                var unlinkSw = FtDebugEnabled ? Stopwatch.StartNew() : null;
                SectionLayoutHelper.EnsureUnlinked(section, index);
                if (unlinkSw != null)
                {
                    unlinkSw.Stop();
                    FtLog($"{tag} story={storyKey} EnsureUnlinked done elapsed_ms={unlinkSw.ElapsedMilliseconds}");
                }

                Dictionary<string, object> headerStory = GetStoryDict(headers, storyKey);
                Dictionary<string, object> footerStory = GetStoryDict(footers, storyKey);

                if (headerStory != null && headerStory.Count > 0)
                {
                    FtLog($"{tag} story={storyKey} ApplySingleStory header keys={headerStory.Count}…");
                    var sw = FtDebugEnabled ? Stopwatch.StartNew() : null;
                    string modeApplied = SectionLayoutHelper.ApplySingleStory(
                        section.Headers[index].Range,
                        headerStory,
                        contentMode,
                        $"{tag} header:{storyKey}");
                    if (sw != null)
                    {
                        sw.Stop();
                        FtLog($"{tag} story={storyKey} ApplySingleStory header done mode={modeApplied} elapsed_ms={sw.ElapsedMilliseconds}");
                    }

                    TrackStoryApply(modeApplied, storyKey, "header", result);
                }
                else
                {
                    FtLog($"{tag} story={storyKey} header skipped (empty)");
                }

                if (footerStory != null && footerStory.Count > 0)
                {
                    FtLog($"{tag} story={storyKey} ApplySingleStory footer keys={footerStory.Count}…");
                    var sw = FtDebugEnabled ? Stopwatch.StartNew() : null;
                    string modeApplied = SectionLayoutHelper.ApplySingleStory(
                        section.Footers[index].Range,
                        footerStory,
                        contentMode,
                        $"{tag} footer:{storyKey}");
                    if (sw != null)
                    {
                        sw.Stop();
                        FtLog($"{tag} story={storyKey} ApplySingleStory footer done mode={modeApplied} elapsed_ms={sw.ElapsedMilliseconds}");
                    }

                    TrackStoryApply(modeApplied, storyKey, "footer", result);
                }
                else
                {
                    FtLog($"{tag} story={storyKey} footer skipped (empty)");
                }
            }
        }

        private static void TrackStoryApply(
            string modeApplied,
            string storyKey,
            string storyKind,
            PageLayoutApplyResult result)
        {
            if (modeApplied == "none")
            {
                return;
            }

            result.HeadersFootersUpdated++;
            result.HfContentModes.Add($"{storyKind}:{storyKey}:{modeApplied}");
            if (modeApplied == "full")
            {
                result.HfContentFullCount++;
            }
            else if (modeApplied == "format_only")
            {
                result.HfContentFormatOnlyCount++;
            }
        }

        private static void ClearStory(Word.Range range)
        {
            SectionLayoutHelper.ClearStoryRange(range);
        }

        private static Word.WdHeaderFooterIndex MapStoryKey(string storyKey)
        {
            switch (storyKey)
            {
                case "first":
                    return Word.WdHeaderFooterIndex.wdHeaderFooterFirstPage;
                case "even":
                    return Word.WdHeaderFooterIndex.wdHeaderFooterEvenPages;
                default:
                    return Word.WdHeaderFooterIndex.wdHeaderFooterPrimary;
            }
        }

        private static Dictionary<string, object> GetStoryDict(
            Dictionary<string, object> container,
            string storyKey)
        {
            if (container == null || !container.TryGetValue(storyKey, out object o))
            {
                return null;
            }

            return NormalizeDict(o);
        }

        private static Dictionary<string, object> GetDict(Dictionary<string, object> dict, string key)
        {
            if (dict == null || !dict.TryGetValue(key, out object o))
            {
                return null;
            }

            return NormalizeDict(o);
        }

        private static bool GetBool(Dictionary<string, object> dict, string key)
        {
            if (dict == null || !dict.TryGetValue(key, out object o) || o == null)
            {
                return false;
            }

            if (o is bool b)
            {
                return b;
            }

            return string.Equals(o.ToString(), "true", StringComparison.OrdinalIgnoreCase)
                || o.ToString() == "1";
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
