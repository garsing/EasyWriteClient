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
                    var wordApp = wordApplication as Word.Application;
                    if (wordApp?.ActiveDocument == null)
                    {
                        return Task.FromResult(new ToolResult { Success = false, Error = "没有活动的 Word 文档" });
                    }

                    DocumentState.BindAndActivate(wordApp.ActiveDocument);
                    DocumentState.EnsureCurrDocUuid(wordApp.ActiveDocument);

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
