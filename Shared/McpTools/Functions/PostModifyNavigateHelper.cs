using System;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    /// <summary>
    /// 文档变更工具成功后：移光标到目标后沿（或新建对象换行后），再视配置滚视口。
    /// </summary>
    public static class PostModifyNavigateHelper
    {
        /// <summary>通则：光标落到 targetRange.End，再视配置滚视口。</summary>
        public static bool NavigateAfterEnd(Word.Application app, Word.Range targetRange, string debugTag = "")
        {
            if (app == null || targetRange == null)
            {
                LogFail(debugTag, "app 或 targetRange 为空");
                return false;
            }

            try
            {
                int pos = targetRange.End;
                if (app.Selection == null)
                {
                    LogFail(debugTag, "Selection 不可用");
                    return false;
                }

                app.Selection.SetRange(pos, pos);
                LogOk(debugTag, $"SetRange End={pos}");
                return MaybeScrollToSelection(app, debugTag);
            }
            catch (Exception ex)
            {
                LogFail(debugTag, ex.Message);
                return false;
            }
        }

        /// <summary>新建表/图：End 处换行 → 光标在换行符后 → 视配置滚视口。</summary>
        public static bool NavigateAfterInsertObject(Word.Application app, Word.Range objectRange, string debugTag = "")
        {
            if (app == null || objectRange == null)
            {
                LogFail(debugTag, "app 或 objectRange 为空");
                return false;
            }

            try
            {
                Word.Range range = objectRange.Duplicate;
                range.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                range.InsertParagraphAfter();
                range.Collapse(Word.WdCollapseDirection.wdCollapseEnd);

                if (app.Selection == null)
                {
                    LogFail(debugTag, "Selection 不可用");
                    return false;
                }

                app.Selection.SetRange(range.Start, range.End);
                LogOk(debugTag, $"InsertParagraphAfter SetRange={range.Start}");
                return MaybeScrollToSelection(app, debugTag);
            }
            catch (Exception ex)
            {
                LogFail(debugTag, ex.Message);
                return false;
            }
        }

        private static bool MaybeScrollToSelection(Word.Application app, string debugTag)
        {
            if (!ConfigManager.GetDocumentNavigateSettings().AutoNavigateAfterModify)
            {
                LogOk(debugTag, "AutoNavigateAfterModify=false，跳过 Scroll");
                return true;
            }

            if (app?.Selection?.Range == null)
            {
                LogFail(debugTag, "Selection.Range 不可用");
                return false;
            }

            return WordRangeFinder.ScrollRangeIntoView(app, app.Selection.Range, debugTag);
        }

        private static void LogOk(string debugTag, string message)
        {
            string tag = string.IsNullOrEmpty(debugTag) ? "" : $"({debugTag}) ";
            System.Diagnostics.Debug.WriteLine($"[PostModifyNavigate] {tag}{message}");
        }

        private static void LogFail(string debugTag, string message)
        {
            string tag = string.IsNullOrEmpty(debugTag) ? "" : $"({debugTag}) ";
            System.Diagnostics.Debug.WriteLine($"[PostModifyNavigate] {tag}失败: {message}");
        }
    }
}
