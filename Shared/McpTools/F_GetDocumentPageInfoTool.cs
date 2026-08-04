using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    /// <summary>
    /// 查询当前活动 Word 文档分页信息（总页数、光标页），不渲染图像。
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

                    var app = wordApplication as Word.Application;
                    if (app == null)
                    {
                        return new ToolResult { Success = false, Error = "Word应用程序不可用" };
                    }

                    if (!ChannelDocument.TryResolve(args, wordApplication, out Word.Document doc, out ToolResult resolveError))
                    {
                        return resolveError;
                    }

                    var info = PageCaptureHelper.GetDocumentPageInfo(app);
                    return new ToolResult
                    {
                        Success = true,
                        Data = new Dictionary<string, object>
                        {
                            ["total_pages"] = info.TotalPages,
                            ["cursor_page_number"] = info.CursorPageNumber,
                            ["doc_title"] = info.DocTitle
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
