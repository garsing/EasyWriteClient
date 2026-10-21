using System;
using System.Diagnostics;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    /// <summary>
    /// 从 checkpoint .docx 按节复制 Story 与节属性到活动文档（D4 / I2）。
    /// </summary>
    public static class DocumentRestoreHelper
    {
        public static void RestoreActiveDocumentFromCheckpoint(
            string checkpointPath,
            Word.Application app,
            Word.Document document = null)
        {
            if (app == null)
            {
                throw new ArgumentNullException(nameof(app));
            }

            if (string.IsNullOrWhiteSpace(checkpointPath))
            {
                throw new ArgumentException("checkpointPath is required", nameof(checkpointPath));
            }

            Word.Document target = document;
            if (target == null)
            {
                try
                {
                    target = app.ActiveDocument;
                }
                catch (Exception)
                {
                    target = null;
                }
            }

            if (target == null)
            {
                throw new InvalidOperationException("没有活动的 Word 文档");
            }

            // 光标由 WordCheckpointKindAdapter.RestoreCaret 按建副本时记录恢复（D2）。
            Word.Document source = null;
            try
            {
                source = app.Documents.Open(
                    checkpointPath,
                    ReadOnly: true,
                    Visible: false);

                CopyAllStoriesAndSectionProps(source, target);
            }
            finally
            {
                if (source != null)
                {
                    try
                    {
                        source.Close(SaveChanges: false);
                    }
                    catch
                    {
                        // ignore close errors
                    }
                }
            }
        }

        private static void AlignSectionCount(Word.Document source, Word.Document target, int sourceCount)
        {
            while (target.Sections.Count < sourceCount)
            {
                Word.Range endRange = target.Content;
                endRange.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                endRange.InsertBreak(Word.WdBreakType.wdSectionBreakNextPage);
            }
        }

        private static void CopyAllStoriesAndSectionProps(Word.Document source, Word.Document target)
        {
            int sourceCount = source.Sections.Count;
            AlignSectionCount(source, target, sourceCount);

            int sectionsToCopy = Math.Min(sourceCount, target.Sections.Count);
            for (int i = 1; i <= sectionsToCopy; i++)
            {
                Word.Section srcSection = source.Sections[i];
                Word.Section dstSection = target.Sections[i];

                SectionLayoutHelper.CopyPageSetup(srcSection.PageSetup, dstSection.PageSetup);
                CopyHeaderFooterStory(srcSection, dstSection, Word.WdHeaderFooterIndex.wdHeaderFooterPrimary);
                CopyHeaderFooterStory(srcSection, dstSection, Word.WdHeaderFooterIndex.wdHeaderFooterFirstPage);
                CopyHeaderFooterStory(srcSection, dstSection, Word.WdHeaderFooterIndex.wdHeaderFooterEvenPages);
                CopySectionBody(srcSection, dstSection, i == sourceCount);
            }
        }

        private static void CopyPageSetup(Word.PageSetup src, Word.PageSetup dst)
        {
            SectionLayoutHelper.CopyPageSetup(src, dst);
        }

        private static void CopyHeaderFooterStory(
            Word.Section srcSection,
            Word.Section dstSection,
            Word.WdHeaderFooterIndex index)
        {
            try
            {
                SectionLayoutHelper.CopyStoryFormattedText(
                    srcSection.Headers[index].Range,
                    dstSection.Headers[index].Range);
                SectionLayoutHelper.CopyStoryFormattedText(
                    srcSection.Footers[index].Range,
                    dstSection.Footers[index].Range);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[Checkpoint] CopyHeaderFooter {index} warning: {ex.Message}");
            }
        }

        private static void CopySectionBody(Word.Section srcSection, Word.Section dstSection, bool isLastSection)
        {
            Word.Range srcRange = srcSection.Range;
            Word.Range dstRange = dstSection.Range;

            if (!isLastSection)
            {
                srcRange = srcRange.Duplicate;
                if (srcRange.End > srcRange.Start)
                {
                    srcRange.MoveEnd(Word.WdUnits.wdCharacter, -1);
                }
            }

            CopyRangeText(srcRange, dstRange);
        }

        private static void CopyRangeText(Word.Range sourceRange, Word.Range targetRange)
        {
            if (sourceRange == null || targetRange == null)
            {
                return;
            }

            try
            {
                targetRange.FormattedText = sourceRange.FormattedText;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[Checkpoint] CopyRangeText warning: {ex.Message}");
            }
        }
    }
}
