using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace WordAddIn1
{
    /// <summary>
    /// 恢复文档检查点。按当前渠道灌回文字稿 / 工作簿 / 演示文稿。
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

                    if (!ChannelContext.TryResolveChannel(args, out IOperationChannel channel, out string resolveError)
                        || channel == null)
                    {
                        return Task.FromResult(new ToolResult
                        {
                            Success = false,
                            Error = resolveError ?? "无可用文档渠道"
                        });
                    }

                    if (!CheckpointKindAdapters.TryFromChannel(channel, out CheckpointLiveTarget live, out string kindError))
                    {
                        return Task.FromResult(new ToolResult { Success = false, Error = kindError });
                    }

                    System.Diagnostics.Debug.WriteLine(
                        $"[F_restore_document_checkpoint] channel_id={channel.ChannelId}, ts={timestamp}");

                    return Task.FromResult(DocumentCheckpointService.RestoreCheckpoint(timestamp, live));
                }
                catch (Exception ex)
                {
                    return Task.FromResult(new ToolResult { Success = false, Error = $"恢复副本失败: {ex.Message}" });
                }
            };
        }
    }
}
