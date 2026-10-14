using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using WordAddIn1.DocumentHost;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    /// <summary>清除文档全部数据溯源。经 <see cref="DocumentHostAdapter"/>。</summary>
    public static class F_ClearDataProvenanceTool
    {
        public static void Register(
            Dictionary<string, Func<Dictionary<string, object>, Task<ToolResult>>> toolRegistry,
            object wordApplication)
        {
            toolRegistry["F_clear_data_provenance"] = async (args) =>
            {
                try
                {
                    if (!DocumentHostAdapter.TryResolveInteropDocument(
                            args,
                            wordApplication,
                            out InteropDocumentHandle docHandle,
                            out ToolResult resolveError))
                    {
                        return resolveError ?? Fail("无法解析文档渠道");
                    }

                    Word.Document doc = docHandle.Document;
                    if (!DataProvenanceClearHelper.HasProvenance(doc))
                    {
                        return new ToolResult
                        {
                            Success = true,
                            Data = new { cleared = false },
                        };
                    }

                    DataProvenanceClearHelper.ClearAll(doc);
                    DocumentState.ClearProcessDocumentCache("clear_data_provenance");
                    await Task.CompletedTask;
                    return new ToolResult
                    {
                        Success = true,
                        Data = new { cleared = true },
                    };
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"[ClearDataProvenance] 执行失败: {ex.Message}");
                    return Fail("执行失败: " + ex.Message);
                }
            };
        }

        private static ToolResult Fail(string error) =>
            new ToolResult { Success = false, Error = error };
    }
}
