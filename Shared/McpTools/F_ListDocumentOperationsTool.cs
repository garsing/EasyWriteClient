using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
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
                    if (!ChannelDocument.TryResolve(args, wordApplication, out Word.Document doc, out ToolResult resolveError))
                    {
                        return Task.FromResult(resolveError);
                    }

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
