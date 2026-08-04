using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    /// <summary>
    /// 切换默认操作渠道（B3 / D12）。
    /// </summary>
    public static class F_SwitchDefaultChannelTool
    {
        public static void Register(
            Dictionary<string, Func<Dictionary<string, object>, Task<ToolResult>>> toolRegistry,
            object wordApplication)
        {
            toolRegistry["F_switch_default_channel"] = async (args) =>
            {
                try
                {
                    string channelId = args != null && args.ContainsKey("channel_id")
                        ? args["channel_id"]?.ToString()?.Trim()
                        : null;
                    if (string.IsNullOrEmpty(channelId))
                    {
                        return new ToolResult { Success = false, Error = "必须提供 channel_id 参数" };
                    }

                    if (!ChannelRegistry.TryGetWord(channelId, out WordChannel channel))
                    {
                        return new ToolResult { Success = false, Error = "未知 channel_id: " + channelId };
                    }

                    if (!channel.TryGetLiveDocument(out Word.Document doc))
                    {
                        return new ToolResult
                        {
                            Success = false,
                            Error = "渠道对应的 Word 文档已关闭: " + channelId
                        };
                    }

                    if (!ChannelRegistry.SetDefault(channelId))
                    {
                        return new ToolResult { Success = false, Error = "设置默认渠道失败: " + channelId };
                    }

                    try
                    {
                        doc.Activate();
                    }
                    catch (Exception)
                    {
                        // 可选 Activate；失败不阻断
                    }

                    DocumentState.BindAndActivate(doc);

                    await Task.CompletedTask;

                    return new ToolResult
                    {
                        Success = true,
                        Data = new Dictionary<string, object>
                        {
                            ["channel_id"] = channel.ChannelId,
                            ["kind"] = "word",
                            ["doc_uuid"] = channel.DocUuid,
                            ["path"] = channel.FilePath ?? "",
                            ["name"] = channel.TryGetDisplayName() ?? "",
                        }
                    };
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine("[ERROR] switch_default_channel: " + ex.Message);
                    return new ToolResult
                    {
                        Success = false,
                        Error = "切换默认渠道失败: " + ex.Message
                    };
                }
            };
        }
    }
}
