using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    public static class F_GetDataProvenanceTool
    {
        public static void Register(
            Dictionary<string, Func<Dictionary<string, object>, Task<ToolResult>>> toolRegistry,
            object wordApplication)
        {
            toolRegistry["F_get_data_provenance"] = async (args) =>
            {
                try
                {
                    if (!ChannelDocument.TryResolve(args, wordApplication, out Word.Document doc, out ToolResult resolveError))
                    {
                        return resolveError;
                    }

                    object data = DataProvenanceReadHelper.ReadAll(doc);
                    await Task.CompletedTask;

                    return new ToolResult
                    {
                        Success = true,
                        Data = data,
                    };
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"[GetDataProvenance] 执行失败: {ex.Message}");
                    return new ToolResult { Success = false, Error = $"执行失败: {ex.Message}" };
                }
            };
        }
    }
}
