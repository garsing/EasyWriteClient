using System;
using System.Collections.Generic;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    public sealed class PageSetupApplyOptions
    {
        /// <summary>null = 全节；1-based 单节。</summary>
        public int? SectionIndex { get; set; }

        public bool SkipIfNearlyEqual { get; set; } = true;

        public bool RestorePrinterAfter { get; set; } = true;
    }

    public sealed class PageSetupApplyResult
    {
        public bool PrinterReady { get; set; }
        public bool PrinterSwitched { get; set; }
        public string ActivePrinterBefore { get; set; }
        public string ActivePrinterAfter { get; set; }
        public int SectionsAttempted { get; set; }
        public int SectionsSucceeded { get; set; }
        public List<string> Warnings { get; set; } = new List<string>();
    }

    /// <summary>
    /// 写 Section.PageSetup 的唯一入口；每次写前 Ensure Word 活动打印机为 PDF/XPS。
    /// </summary>
    public static class DocumentPageSetupApplier
    {
        private const string NoPrinterWarning = "page_setup_skipped_no_printer";

        public static PageSetupApplyResult ApplyPageSetupToDocument(
            Word.Document doc,
            Dictionary<string, object> pageSetupDict,
            PageSetupApplyOptions options = null)
        {
            var result = new PageSetupApplyResult();
            if (options == null)
            {
                options = new PageSetupApplyOptions();
            }

            if (doc == null)
            {
                result.Warnings.Add("document 为空，跳过 page_setup");
                return result;
            }

            if (pageSetupDict == null || pageSetupDict.Count == 0)
            {
                result.Warnings.Add("page_setup 为空，跳过");
                return result;
            }

            Word.Application app = null;
            try
            {
                app = doc.Application;
            }
            catch (Exception ex)
            {
                result.Warnings.Add($"doc.Application failed: {ex.Message}");
                return result;
            }

            try
            {
                result.ActivePrinterBefore = app?.ActivePrinter;
            }
            catch
            {
                result.ActivePrinterBefore = null;
            }

            bool ready = WordActivePrinterHelper.EnsurePreferredActivePrinter(app);
            result.PrinterReady = ready;
            try
            {
                result.ActivePrinterAfter = app?.ActivePrinter;
            }
            catch
            {
                result.ActivePrinterAfter = null;
            }

            result.PrinterSwitched =
                !string.IsNullOrEmpty(result.ActivePrinterAfter)
                && !string.Equals(
                    result.ActivePrinterBefore ?? "",
                    result.ActivePrinterAfter,
                    StringComparison.OrdinalIgnoreCase);

            if (!ready)
            {
                result.Warnings.Add(NoPrinterWarning);
                return result;
            }

            int sectionCount;
            try
            {
                sectionCount = doc.Sections.Count;
            }
            catch (Exception ex)
            {
                result.Warnings.Add($"doc.Sections.Count failed: {ex.Message}");
                return result;
            }

            int start = 1;
            int end = sectionCount;
            if (options.SectionIndex.HasValue)
            {
                int idx = options.SectionIndex.Value;
                if (idx < 1 || idx > sectionCount)
                {
                    result.Warnings.Add($"section_index={idx} 超出范围 1..{sectionCount}");
                    return result;
                }

                start = idx;
                end = idx;
            }

            string previousForRestore = options.RestorePrinterAfter ? result.ActivePrinterBefore : null;
            try
            {
                for (int i = start; i <= end; i++)
                {
                    result.SectionsAttempted++;
                    try
                    {
                        Word.Section section = doc.Sections[i];
                        SectionLayoutHelper.ApplyPageSetupFromDict(
                            section.PageSetup,
                            pageSetupDict,
                            options.SkipIfNearlyEqual);
                        result.SectionsSucceeded++;
                    }
                    catch (Exception ex)
                    {
                        result.Warnings.Add($"section {i} PageSetup failed: {ex.Message}");
                    }
                }
            }
            finally
            {
                if (options.RestorePrinterAfter)
                {
                    WordActivePrinterHelper.RestoreActivePrinter(app, previousForRestore);
                }
            }

            return result;
        }
    }
}
