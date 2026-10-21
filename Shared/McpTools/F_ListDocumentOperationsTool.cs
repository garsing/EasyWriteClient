using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace WordAddIn1
{
    /// <summary>
    /// 列出当前渠道文档的操作/检查点记录（只读）。
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

                    string uuid = live.DocUuid;
                    System.Diagnostics.Debug.WriteLine(
                        $"[F_list_document_operations] channel_id={channel.ChannelId} uuid={uuid}");

                    string text = DocumentCheckpointService.FormatListTextForDocUuid(uuid);
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
