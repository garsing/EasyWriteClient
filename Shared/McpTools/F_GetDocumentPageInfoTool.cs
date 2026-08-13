using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using WordAddIn1.DocumentHost;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    /// <summary>
    /// 查询渠道文档分页信息（总页数、光标页），不渲染图像。
    /// 经 <see cref="DocumentHostAdapter"/>（Word/WPS）；只读不抢前台。
    /// </summary>
    public static class F_GetDocumentPageInfoTool
    {
        public static void Register(
            Dictionary<string, Func<Dictionary<string, object>, Task<ToolResult>>> toolRegistry,
            object wordApplication)
        {
            toolRegistry["F_get_document_page_info"] = async (args) =>
            {
                try
                {
                    System.Diagnostics.Debug.WriteLine("[F_get_document_page_info] 开始执行");

                    if (!DocumentHostAdapter.TryResolveInteropDocument(
                            args,
                            wordApplication,
                            out InteropDocumentHandle docHandle,
                            out ToolResult resolveError,
                            activateDocument: false))
                    {
                        System.Diagnostics.Debug.WriteLine(
                            $"[F_get_document_page_info] resolve failed: {resolveError?.Error}; " +
                            $"arg.channel_id={ChannelContext.TryGetChannelIdFromParameters(args) ?? "(null)"}; " +
                            $"default={ChannelRegistry.DefaultChannelId ?? "(null)"}");
                        return resolveError;
                    }

                    Word.Document doc = docHandle.Document;
                    System.Diagnostics.Debug.WriteLine(
                        $"[F_get_document_page_info] host={docHandle.HostName}, channel_id={docHandle.ChannelId}");

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

                    var info = PageCaptureHelper.GetDocumentPageInfo(app, doc);
                    await Task.CompletedTask;
                    return new ToolResult
                    {
                        Success = true,
                        Data = new Dictionary<string, object>
                        {
                            ["total_pages"] = info.TotalPages,
                            ["cursor_page_number"] = info.CursorPageNumber,
                            ["doc_title"] = info.DocTitle,
                            ["channel_id"] = docHandle.ChannelId ?? "",
                            ["host"] = docHandle.HostName ?? ""
                        }
                    };
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"[F_get_document_page_info] 失败: {ex.Message}");
                    return new ToolResult { Success = false, Error = $"查询文档页信息失败: {ex.Message}" };
                }
            };
        }
    }
}
