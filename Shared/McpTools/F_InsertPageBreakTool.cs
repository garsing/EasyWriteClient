using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using WordAddIn1.DocumentHost;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    /// <summary>
    /// 在文字渠道已有定位点后插入分页符，后面的内容从下一页开始。经 <see cref="DocumentHostAdapter"/>（Word/WPS）。
    /// </summary>
    public static class F_InsertPageBreakTool
    {
        private const string SuccessMessage = "已插入分页符，后面的内容从下一页开始。";
        private const string PptHint =
            "这是 Word 分页符，不要用 F_manage_ppt_slide，也不要往 insert 里写 \\f";
        private const int WdPageBreak = 7;

        public static void Register(
            Dictionary<string, Func<Dictionary<string, object>, Task<ToolResult>>> toolRegistry,
            object wordApplication)
        {
            toolRegistry["F_insert_page_break"] = async (args) =>
            {
                try
                {
                    if (!DocumentHostAdapter.TryResolveInteropDocument(
                            args,
                            wordApplication,
                            out InteropDocumentHandle docHandle,
                            out ToolResult resolveError))
                    {
                        return WithPptHint(resolveError);
                    }

                    Word.Document doc = docHandle.Document;
                    Word.Application app = TryGetApplication(doc, wordApplication);
                    System.Diagnostics.Debug.WriteLine(
                        $"[InsertPageBreak] host={docHandle.HostName}, channel_id={docHandle.ChannelId}");

                    var before = TryGetPageInfo(app, doc);
                    var locate = InsertTargetLocatorHelper.TryResolveInsertRange(doc, args);
                    if (!locate.Success)
                    {
                        return Fail(locate.Error);
                    }

                    string kind = locate.UsedCursor
                        ? "cursor"
                        : (locate.Spec != null ? locate.Spec.Kind.ToString() : "(none)");
                    System.Diagnostics.Debug.WriteLine(
                        $"[InsertPageBreak] used_cursor={locate.UsedCursor} target={kind} " +
                        $"total_pages_before={before.TotalPages}");

                    InsertPageBreak(locate.Range);

                    try
                    {
                        doc.Repaginate();
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine(
                            $"[InsertPageBreak] Repaginate: {ex.Message}");
                    }

                    var after = TryGetPageInfo(app, doc);
                    if (after.TotalPages <= 0 && before.TotalPages > 0)
                    {
                        after.TotalPages = before.TotalPages;
                        after.CursorPageNumber = before.CursorPageNumber;
                        after.DocTitle = before.DocTitle;
                    }

                    try
                    {
                        MoveCursorAfterBreak(app, locate.Range);
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine(
                            $"[InsertPageBreak] 移光标失败: {ex.Message}");
                    }

                    var data = new Dictionary<string, object>
                    {
                        ["total_pages"] = after.TotalPages,
                        ["cursor_page_number"] = after.CursorPageNumber,
                        ["channel_id"] = docHandle.ChannelId ?? "",
                        ["host"] = docHandle.HostName ?? "",
                        ["message"] = SuccessMessage
                    };
                    var echo = InsertTargetLocatorHelper.BuildTargetEcho(locate.Spec);
                    if (echo != null)
                    {
                        data["target_echo"] = echo;
                    }

                    System.Diagnostics.Debug.WriteLine(
                        $"[InsertPageBreak] ok used_cursor={locate.UsedCursor} target={kind} " +
                        $"total_pages_before={before.TotalPages} total_pages_after={after.TotalPages}");

                    await Task.CompletedTask;
                    return new ToolResult { Success = true, Data = data };
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"[InsertPageBreak] 失败: {ex.Message}");
                    return Fail($"插入分页符失败: {ex.Message}");
                }
            };
        }

        private static ToolResult Fail(string error)
        {
            return new ToolResult { Success = false, Error = error };
        }

        private static ToolResult WithPptHint(ToolResult error)
        {
            if (error == null)
            {
                return Fail(PptHint);
            }

            string text = error.Error ?? "";
            if (text.IndexOf("F_manage_ppt_slide", StringComparison.Ordinal) < 0)
            {
                error.Error = string.IsNullOrEmpty(text)
                    ? PptHint
                    : text.TrimEnd() + "。" + PptHint;
            }

            return error;
        }

        private static void InsertPageBreak(Word.Range range)
        {
            try
            {
                object breakType = Word.WdBreakType.wdPageBreak;
                range.InsertBreak(ref breakType);
                return;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine(
                    $"[InsertPageBreak] InsertBreak(enum) 失败: {ex.Message}");
            }

            dynamic dyn = range;
            dyn.InsertBreak(WdPageBreak);
        }

        private static void MoveCursorAfterBreak(Word.Application app, Word.Range range)
        {
            if (app?.Selection == null || range == null)
            {
                return;
            }

            range.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
            app.Selection.SetRange(range.Start, range.Start);
        }

        private static PageCaptureHelper.DocumentPageInfo TryGetPageInfo(
            Word.Application app,
            Word.Document doc)
        {
            try
            {
                return PageCaptureHelper.GetDocumentPageInfo(app, doc);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine(
                    $"[InsertPageBreak] 查页失败: {ex.Message}");
                return new PageCaptureHelper.DocumentPageInfo();
            }
        }

        private static Word.Application TryGetApplication(Word.Document doc, object wordApplication)
        {
            try
            {
                if (doc != null)
                {
                    return doc.Application;
                }
            }
            catch (Exception)
            {
            }

            return wordApplication as Word.Application;
        }
    }
}
