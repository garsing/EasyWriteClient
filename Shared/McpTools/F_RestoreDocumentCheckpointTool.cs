using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace WordAddIn1
{
    public static class F_RestoreDocumentCheckpointTool
    {
        public static void Register(
            Dictionary<string, Func<Dictionary<string, object>, Task<ToolResult>>> toolRegistry,
            object wordApplication)
        {
            toolRegistry["F_restore_document_checkpoint"] = (args) =>
            {
                try
                {
                    if (args == null || !args.TryGetValue("timestamp", out object tsObj))
                    {
                        return Task.FromResult(new ToolResult { Success = false, Error = "缺少必需参数：timestamp" });
                    }

                    string timestamp = (tsObj?.ToString() ?? "").Trim();
                    if (string.IsNullOrEmpty(timestamp))
                    {
                        return Task.FromResult(new ToolResult { Success = false, Error = "timestamp 不能为空" });
                    }

                    return Task.FromResult(DocumentCheckpointService.RestoreCheckpoint(timestamp));
                }
                catch (Exception ex)
                {
                    return Task.FromResult(new ToolResult { Success = false, Error = $"恢复副本失败: {ex.Message}" });
                }
            };
        }
    }
}
