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
        private sealed class DocumentViewState
        {
            public int SelectionStart { get; set; }
            public int SelectionEnd { get; set; }
            public int DocumentEnd { get; set; }
            public int VerticalPercentScrolled { get; set; }
        }

        public static void RestoreActiveDocumentFromCheckpoint(string checkpointPath, Word.Application app)
        {
            if (app == null)
            {
                throw new ArgumentNullException(nameof(app));
            }

            if (string.IsNullOrWhiteSpace(checkpointPath))
            {
                throw new ArgumentException("checkpointPath is required", nameof(checkpointPath));
            }

            Word.Document target = app.ActiveDocument;
            if (target == null)
            {
                throw new InvalidOperationException("没有活动的 Word 文档");
            }

            DocumentViewState viewState = CaptureViewState(app, target);

            // 不读写 TrackRevisions（2026-08：取消审阅工具注册与默认关审阅）
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

                RestoreViewStateNear(app, target, viewState);
            }
        }

        /// <summary>
        /// 恢复前记录光标与纵向滚动，供整篇替换后回到大致阅读位置。
        /// </summary>
        private static DocumentViewState CaptureViewState(Word.Application app, Word.Document doc)
        {
            var state = new DocumentViewState();
            if (app == null || doc == null)
            {
                return state;
            }

            try
            {
                state.DocumentEnd = doc.Content?.End ?? 0;
                if (app.Selection != null)
                {
                    state.SelectionStart = app.Selection.Start;
                    state.SelectionEnd = app.Selection.End;
                }

                if (app.ActiveWindow != null)
                {
                    state.VerticalPercentScrolled = app.ActiveWindow.VerticalPercentScrolled;
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[Checkpoint] CaptureViewState warning: {ex.Message}");
            }

            return state;
        }

        /// <summary>
        /// 按恢复前光标在文档中的相对位置映射到新正文，并滚入视口。
        /// 光标在文首但已向下滚动时，回退使用 VerticalPercentScrolled。
        /// </summary>
        private static void RestoreViewStateNear(Word.Application app, Word.Document doc, DocumentViewState state)
        {
            if (app == null || doc == null || state == null)
            {
                return;
            }

            try
            {
                int docEnd = doc.Content?.End ?? 0;
                if (docEnd <= 1)
                {
                    return;
                }

                int maxPos = Math.Max(0, docEnd - 1);
                bool cursorAtDocStart = state.SelectionStart <= 1;
                bool hadScroll = state.VerticalPercentScrolled > 0;

                if (cursorAtDocStart && hadScroll && state.VerticalPercentScrolled >= 5)
                {
                    if (app.ActiveWindow != null)
                    {
                        app.ActiveWindow.VerticalPercentScrolled = state.VerticalPercentScrolled;
                    }

                    Debug.WriteLine(
                        $"[Checkpoint] RestoreViewState scroll-only percent={state.VerticalPercentScrolled}");
                    return;
                }

                int beforeSpan = Math.Max(1, state.DocumentEnd - 1);
                double ratio = Math.Max(0, Math.Min(1, (double)state.SelectionStart / beforeSpan));
                int targetStart = (int)Math.Round(ratio * maxPos);
                targetStart = Math.Max(0, Math.Min(targetStart, maxPos));

                int selectionLen = Math.Max(0, state.SelectionEnd - state.SelectionStart);
                int targetEnd = Math.Min(targetStart + selectionLen, maxPos);
                if (targetEnd < targetStart)
                {
                    targetEnd = targetStart;
                }

                if (app.Selection != null)
                {
                    app.Selection.SetRange(targetStart, targetEnd);
                }

                Word.Range scrollRange = doc.Range(targetStart, targetStart);
                WordRangeFinder.ScrollRangeIntoView(app, scrollRange, "checkpoint-restore");

                Debug.WriteLine(
                    $"[Checkpoint] RestoreViewState ratio={ratio:F3} " +
                    $"before={state.SelectionStart}/{state.DocumentEnd} after={targetStart}/{docEnd}");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[Checkpoint] RestoreViewStateNear warning: {ex.Message}");
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
