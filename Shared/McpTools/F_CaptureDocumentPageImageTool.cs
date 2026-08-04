using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    /// <summary>
    /// 截取当前活动 Word 文档指定页（或光标页）渲染图像，供视觉大模型分析版式。
    /// </summary>
    public static class F_CaptureDocumentPageImageTool
    {
        public static void Register(
            Dictionary<string, Func<Dictionary<string, object>, Task<ToolResult>>> toolRegistry,
            object wordApplication)
        {
            toolRegistry["F_capture_document_page_image"] = async (args) =>
            {
                try
                {
                    System.Diagnostics.Debug.WriteLine("[F_capture_document_page_image] 开始执行");

                    var app = wordApplication as Word.Application;
                    if (app == null)
                    {
                        return new ToolResult { Success = false, Error = "Word应用程序不可用" };
                    }

                    int? requestedPage;
                    string pageErr = TryParsePageNumber(args, out requestedPage);
                    if (pageErr != null)
                    {
                        return new ToolResult { Success = false, Error = pageErr };
                    }

                    if (!ChannelDocument.TryResolve(args, wordApplication, out Word.Document doc, out ToolResult resolveError))
                    {
                        return resolveError;
                    }

                    var capture = PageCaptureHelper.CaptureDocumentPage(app, requestedPage);
                    return new ToolResult
                    {
                        Success = true,
                        Data = new Dictionary<string, object>
                        {
                            ["page_number"] = capture.PageNumber,
                            ["total_pages"] = capture.TotalPages,
                            ["format"] = capture.Format,
                            ["width_px"] = capture.WidthPx,
                            ["height_px"] = capture.HeightPx,
                            ["image_base64"] = capture.ImageBase64,
                            ["captured_at"] = capture.CapturedAt,
                            ["doc_title"] = capture.DocTitle
                        }
                    };
                }
                catch (PageCaptureHelper.PageOutOfRangeException ex)
                {
                    return new ToolResult
                    {
                        Success = false,
                        Error = ex.Message,
                        Data = new Dictionary<string, object>
                        {
                            ["page_number"] = ex.RequestedPage,
                            ["total_pages"] = ex.TotalPages
                        }
                    };
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"[F_capture_document_page_image] 失败: {ex.Message}");
                    return new ToolResult { Success = false, Error = $"截图失败: {ex.Message}" };
                }
            };
        }

        private static string TryParsePageNumber(
            IReadOnlyDictionary<string, object> args,
            out int? pageNumber)
        {
            pageNumber = null;
            if (args == null || !args.TryGetValue("page_number", out object raw) || raw == null)
            {
                return null;
            }

            if (raw is bool)
            {
                return "page_number 必须是整数";
            }

            try
            {
                int val = Convert.ToInt32(raw);
                if (val < 1)
                {
                    return "page_number 必须是 ≥1 的整数";
                }

                pageNumber = val;
                return null;
            }
            catch (Exception)
            {
                return "page_number 必须是整数";
            }
        }
    }
}
