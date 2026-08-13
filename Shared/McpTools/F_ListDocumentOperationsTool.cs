using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using WordAddIn1.DocumentHost;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    /// <summary>
    /// 列出当前渠道文档的操作/检查点记录（只读）。经 <see cref="DocumentHostAdapter"/>。
    /// </summary>
    public static class F_ListDocumentOperationsTool
    {
        public static void Register(
            Dictionary<string, Func<Dictionary<string, object>, Task<ToolResult>>> toolRegistry,
            object wordApplication)
        {
            toolRegistry["F_list_document_operations"] = (args) =>
            {
                try
                {
                    if (!DocumentHostAdapter.TryResolveInteropDocument(
                            args,
                            wordApplication,
                            out InteropDocumentHandle docHandle,
                            out ToolResult resolveError,
                            activateDocument: false))
                    {
                        return Task.FromResult(resolveError);
                    }

                    Word.Document doc = docHandle.Document;
                    System.Diagnostics.Debug.WriteLine(
                        $"[F_list_document_operations] host={docHandle.HostName}, channel_id={docHandle.ChannelId}");

                    DocumentState.EnsureCurrDocUuid(doc);

                    string text = DocumentCheckpointService.FormatListTextForCurrentDocument();
                    return Task.FromResult(new ToolResult
                    {
                        Success = true,
                        Data = text
                    });
                }
                catch (Exception ex)
                {
                    return Task.FromResult(new ToolResult { Success = false, Error = $"读取操作记录失败: {ex.Message}" });
                }
            };
        }
    }
}
