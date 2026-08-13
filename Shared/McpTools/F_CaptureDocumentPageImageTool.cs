using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using WordAddIn1.DocumentHost;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    /// <summary>
    /// 截取渠道文档指定页（或光标页）渲染图像。经 <see cref="DocumentHostAdapter"/>。
    /// 截图需导出 PDF，默认会 Activate 文档（可能抢前台）。
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

                    int? requestedPage;
                    string pageErr = TryParsePageNumber(args, out requestedPage);
                    if (pageErr != null)
                    {
                        return new ToolResult { Success = false, Error = pageErr };
                    }

                    // 截图导出依赖文档/应用状态，允许 Activate
                    if (!DocumentHostAdapter.TryResolveInteropDocument(
                            args,
                            wordApplication,
                            out InteropDocumentHandle docHandle,
                            out ToolResult resolveError))
                    {
                        System.Diagnostics.Debug.WriteLine(
                            $"[F_capture_document_page_image] resolve failed: {resolveError?.Error}");
                        return resolveError;
                    }

                    Word.Document doc = docHandle.Document;
                    System.Diagnostics.Debug.WriteLine(
                        $"[F_capture_document_page_image] host={docHandle.HostName}, channel_id={docHandle.ChannelId}");

                    Word.Application app = null;
                    try
                    {
                        app = doc.Application;
                    }
                    catch (Exception)
                    {
                    }

                    if (app == null)
                    {
                        app = wordApplication as Word.Application;
                    }

                    if (app == null)
                    {
                        return new ToolResult { Success = false, Error = "无法获取文档 Application（Word/WPS）" };
                    }

                    var capture = PageCaptureHelper.CaptureDocumentPage(app, requestedPage, doc);
                    await Task.CompletedTask;
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
                            ["doc_title"] = capture.DocTitle,
                            ["channel_id"] = docHandle.ChannelId ?? "",
                            ["host"] = docHandle.HostName ?? ""
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
            if (args == null || !args.ContainsKey("page_number") || args["page_number"] == null)
            {
                return null;
            }

            object raw = args["page_number"];
            if (raw is int i)
            {
                pageNumber = i;
            }
            else if (raw is long l)
            {
                pageNumber = (int)l;
            }
            else if (!int.TryParse(raw.ToString(), out int parsed))
            {
                return "page_number 必须是整数";
            }
            else
            {
                pageNumber = parsed;
            }

            if (pageNumber.Value < 1)
            {
                return "page_number 必须是 ≥1 的整数";
            }

            return null;
        }
    }
}
