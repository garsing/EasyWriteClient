using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using WordAddIn1.DocumentHost;

namespace WordAddIn1
{
    /// <summary>
    /// 恢复文档检查点。经 <see cref="DocumentHostAdapter"/> 解析渠道后再回滚。
    /// </summary>
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

                    if (!DocumentHostAdapter.TryResolveInteropDocument(
                            args,
                            wordApplication,
                            out InteropDocumentHandle docHandle,
                            out ToolResult resolveError))
                    {
                        System.Diagnostics.Debug.WriteLine(
                            $"[F_restore_document_checkpoint] resolve failed: {resolveError?.Error}");
                        return Task.FromResult(resolveError);
                    }

                    System.Diagnostics.Debug.WriteLine(
                        $"[F_restore_document_checkpoint] host={docHandle.HostName}, channel_id={docHandle.ChannelId}, ts={timestamp}");

                    return Task.FromResult(
                        DocumentCheckpointService.RestoreCheckpoint(timestamp, docHandle.Document));
                }
                catch (Exception ex)
                {
                    return Task.FromResult(new ToolResult { Success = false, Error = $"恢复副本失败: {ex.Message}" });
                }
            };
        }
    }
}
