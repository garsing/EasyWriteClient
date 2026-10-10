using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using WordAddIn1.DocumentHost;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    /// <summary>
    /// 在文字渠道插入分页符，得到一页空白。经 <see cref="DocumentHostAdapter"/>（Word/WPS）。
    /// </summary>
    public static class F_InsertBlankPageTool
    {
        private const string SuccessMessage = "已新增一页空白，光标在新页开头，可以开始写内容。";
        private const string StartSuccessMessage = "已在第 1 页前新增一页空白，光标在新页开头，可以开始写内容。";
        private const string PptHint = "这是 Word 分页，不要用 F_manage_ppt_slide";
        private const int WdPageBreak = 7;

        public static void Register(
            Dictionary<string, Func<Dictionary<string, object>, Task<ToolResult>>> toolRegistry,
            object wordApplication)
        {
            toolRegistry["F_insert_blank_page"] = async (args) =>
            {
                try
                {
                    if (TryGetContractParseFailure(args, out ToolResult contractFail))
                    {
                        return contractFail;
                    }

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
                        $"[InsertBlankPage] host={docHandle.HostName}, channel_id={docHandle.ChannelId}");

                    string parseErr = TryParseOptionalPageNumber(args, out int? pageNumber);
                    if (parseErr != null)
                    {
                        return Fail(parseErr);
                    }

                    string positionErr = TryParseOptionalPosition(args, out bool atStart);
                    if (positionErr != null)
                    {
                        return Fail(positionErr);
                    }

                    if (atStart && pageNumber.HasValue)
                    {
                        return Fail("position 与 page_number 不能同时传");
                    }

                    var before = PageCaptureHelper.GetDocumentPageInfo(app, doc);
                    if (pageNumber.HasValue &&
                        (pageNumber.Value < 1 || pageNumber.Value > before.TotalPages))
                    {
                        return Fail($"页码须为 1～{before.TotalPages}");
                    }

                    int insertPos;
                    int expectedNewPage;
                    string successMessage;
                    if (atStart)
                    {
                        insertPos = Math.Max(0, doc.Content.Start);
                        expectedNewPage = 1;
                        successMessage = StartSuccessMessage;
                    }
                    else
                    {
                        insertPos = ResolveInsertPos(doc, pageNumber, before.TotalPages);
                        expectedNewPage = pageNumber.HasValue && pageNumber.Value < before.TotalPages
                            ? pageNumber.Value + 1
                            : before.TotalPages + 1;
                        successMessage = SuccessMessage;
                    }

                    System.Diagnostics.Debug.WriteLine(
                        $"[InsertBlankPage] page_number={pageNumber?.ToString() ?? "(none)"} " +
                        $"position={(atStart ? "start" : "(none)")} " +
                        $"insertPos={insertPos} total={before.TotalPages}");

                    Word.Range at = doc.Range(insertPos, insertPos);
                    InsertPageBreak(at);

                    try
                    {
                        doc.Repaginate();
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine(
                            $"[InsertBlankPage] Repaginate: {ex.Message}");
                    }

                    PageCaptureHelper.DocumentPageInfo after;
                    try
                    {
                        after = PageCaptureHelper.GetDocumentPageInfo(app, doc);
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine(
                            $"[InsertBlankPage] 插后查页失败: {ex.Message}");
                        after = new PageCaptureHelper.DocumentPageInfo
                        {
                            TotalPages = expectedNewPage,
                            CursorPageNumber = expectedNewPage,
                            DocTitle = before.DocTitle
                        };
                    }

                    try
                    {
                        MoveCursorToPage(app, expectedNewPage);
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine(
                            $"[InsertBlankPage] 移光标失败: {ex.Message}");
                    }

                    await Task.CompletedTask;
                    return new ToolResult
                    {
                        Success = true,
                        Data = new Dictionary<string, object>
                        {
                            ["total_pages"] = after.TotalPages,
                            ["cursor_page_number"] = after.CursorPageNumber,
                            ["channel_id"] = docHandle.ChannelId ?? "",
                            ["host"] = docHandle.HostName ?? "",
                            ["message"] = successMessage
                        }
                    };
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"[InsertBlankPage] 失败: {ex.Message}");
                    return Fail($"新增空白页失败: {ex.Message}");
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

        public static bool TryGetContractParseFailure(
            Dictionary<string, object> args,
            out ToolResult failure)
        {
            failure = null;
            string parseErr = TryParseOptionalPageNumber(args, out int? pageNumber);
            if (parseErr != null)
            {
                failure = DocumentCheckpointMutatingTools.Fail(parseErr);
                return true;
            }

            string positionErr = TryParseOptionalPosition(args, out bool atStart);
            if (positionErr != null)
            {
                failure = DocumentCheckpointMutatingTools.Fail(positionErr);
                return true;
            }

            if (atStart && pageNumber.HasValue)
            {
                failure = DocumentCheckpointMutatingTools.Fail("position 与 page_number 不能同时传");
                return true;
            }

            return false;
        }

        private static string TryParseOptionalPageNumber(
            Dictionary<string, object> args,
            out int? pageNumber)
        {
            pageNumber = null;
            if (args == null || !args.ContainsKey("page_number") || args["page_number"] == null)
            {
                return null;
            }

            object raw = args["page_number"];
            string text = Convert.ToString(raw)?.Trim();
            if (string.IsNullOrEmpty(text))
            {
                return null;
            }

            if (raw is int i)
            {
                pageNumber = i;
                return null;
            }

            if (raw is long l)
            {
                pageNumber = (int)l;
                return null;
            }

            if (int.TryParse(text, out int parsed))
            {
                pageNumber = parsed;
                return null;
            }

            return "页码须为整数";
        }

        private static string TryParseOptionalPosition(
            Dictionary<string, object> args,
            out bool atStart)
        {
            atStart = false;
            if (args == null || !args.ContainsKey("position") || args["position"] == null)
            {
                return null;
            }

            string text = Convert.ToString(args["position"])?.Trim();
            if (string.IsNullOrEmpty(text))
            {
                return null;
            }

            if (!InsertTargetLocatorHelper.IsStartPosition(text))
            {
                return InsertTargetLocatorHelper.PositionStartOnlyError;
            }

            atStart = true;
            return null;
        }

        private static int ResolveInsertPos(Word.Document doc, int? pageNumber, int totalPages)
        {
            int endPos = Math.Max(0, doc.Content.End - 1);
            if (!pageNumber.HasValue || pageNumber.Value >= totalPages)
            {
                return endPos;
            }

            Word.Range nextPage = doc.GoTo(
                Word.WdGoToItem.wdGoToPage,
                Word.WdGoToDirection.wdGoToAbsolute,
                pageNumber.Value + 1);
            int start = nextPage != null ? nextPage.Start : endPos;
            return Math.Max(0, Math.Min(start, endPos));
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
                    $"[InsertBlankPage] InsertBreak(enum) 失败: {ex.Message}");
            }

            dynamic dyn = range;
            dyn.InsertBreak(WdPageBreak);
        }

        private static void MoveCursorToPage(Word.Application app, int pageNumber)
        {
            if (app?.Selection == null || pageNumber < 1)
            {
                return;
            }

            object what = Word.WdGoToItem.wdGoToPage;
            object which = Word.WdGoToDirection.wdGoToAbsolute;
            object count = pageNumber;
            app.Selection.GoTo(ref what, ref which, ref count);
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
