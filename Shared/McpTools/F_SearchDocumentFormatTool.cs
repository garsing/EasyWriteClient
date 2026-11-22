using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using WordAddIn1.DocumentHost;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    public static class F_SearchDocumentFormatTool
    {
        public static void Register(
            Dictionary<string, Func<Dictionary<string, object>, Task<ToolResult>>> toolRegistry,
            object wordApplication)
        {
            toolRegistry["F_search_document_format"] = (args) =>
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
                    return Task.FromResult(FormatDocumentSearch.Run(doc, args, docHandle.DisallowBackendApi));
                }
                catch (Exception ex)
                {
                    return Task.FromResult(new ToolResult
                    {
                        Success = false,
                        Error = "search_document_format 失败: " + ex.Message
                    });
                }
            };
        }
    }
}
