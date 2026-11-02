using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using WordAddIn1.BrowserHost;

namespace WordAddIn1
{
    /// <summary>
    /// 易写浏览页内交互：click / hover / type / scroll / press / select。
    /// 可一次传 actions 数组。
    /// </summary>
    public static class F_BrowserInteractTool
    {
        public static void Register(
            Dictionary<string, Func<Dictionary<string, object>, Task<ToolResult>>> toolRegistry,
            object wordApplication)
        {
            toolRegistry["F_browser_interact"] = async (args) =>
            {
                try
                {
                    AgentRunCancellation.ThrowIfCancelled();

                    if (!BrowserInteractBatch.TryCompile(args, out List<BrowserInteractStep> steps, out string compileError))
                    {
                        return Fail(compileError ?? "必须提供 action 或 actions");
                    }

                    string channelId = TryGetString(args, "channel_id");
                    BrowserChannel channel;
                    if (!string.IsNullOrWhiteSpace(channelId))
                    {
                        if (!ChannelRegistry.TryGetBrowser(channelId, out channel) || channel == null)
                        {
                            return Fail("channel_id 不是浏览器渠道: " + channelId);
                        }
                    }
                    else if (!ChannelRegistry.TryGetDefault(out IOperationChannel def) || def == null)
                    {
                        return Fail("无默认渠道；请先 F_browser_navigate / F_browser_snapshot，或传入 channel_id");
                    }
                    else if (!(def is BrowserChannel browser))
                    {
                        return Fail("默认渠道不是浏览器渠道；请传入浏览器 channel_id（如 b1）");
                    }
                    else
                    {
                        channel = browser;
                    }

                    if (!string.Equals(channel.Track, "agent", StringComparison.OrdinalIgnoreCase)
                        && !string.Equals(channel.Track, "attach", StringComparison.OrdinalIgnoreCase))
                    {
                        return Fail("不支持的浏览器 track: " + channel.Track);
                    }

                    if (!channel.IsLive())
                    {
                        return Fail("浏览器渠道对应的页面已关闭: " + channel.ChannelId);
                    }

                    return await BrowserInteractBatch.Run(channel, steps).ConfigureAwait(true);
                }
                catch (OperationCanceledException)
                {
                    return Fail("cancelled by user");
                }
                catch (Exception ex)
                {
                    return Fail("F_browser_interact 失败: " + ex.Message);
                }
            };
        }

        private static string TryGetString(Dictionary<string, object> args, string key)
        {
            if (args == null || !args.ContainsKey(key) || args[key] == null)
            {
                return null;
            }

            return Convert.ToString(args[key])?.Trim();
        }

        private static ToolResult Fail(string error)
        {
            return new ToolResult { Success = false, Error = error };
        }
    }
}
