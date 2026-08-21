using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using WordAddIn1.Terminal;

namespace WordAddIn1
{
    public static class F_CloseTerminalTool
    {
        public static void Register(
            Dictionary<string, Func<Dictionary<string, object>, Task<ToolResult>>> toolRegistry,
            object wordApplication)
        {
            toolRegistry["F_close_terminal"] = _ =>
            {
                try
                {
                    AgentRunCancellation.ThrowIfCancelled();
                    if (ChannelHost.Kind != ChannelHostKind.Desktop)
                    {
                        return Task.FromResult(new ToolResult
                        {
                            Success = false,
                            Error = "仅易写桌面应用支持"
                        });
                    }

                    bool existed = TerminalSessionHub.Dispose(ConversationContext.CurrentId);
                    return Task.FromResult(new ToolResult
                    {
                        Success = true,
                        Data = new { closed = true, already = !existed }
                    });
                }
                catch (OperationCanceledException)
                {
                    return Task.FromResult(new ToolResult { Success = false, Error = "cancelled by user" });
                }
                catch (Exception ex)
                {
                    return Task.FromResult(new ToolResult { Success = false, Error = "关闭终端失败: " + ex.Message });
                }
            };
        }
    }
}
