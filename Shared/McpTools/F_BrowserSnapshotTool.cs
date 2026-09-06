using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using WordAddIn1.BrowserHost;

namespace WordAddIn1
{
    /// <summary>
    /// 读取易写浏览页无障碍树快照（大概 / 按 ref 展开）。
    /// </summary>
    public static class F_BrowserSnapshotTool
    {
        public static void Register(
            Dictionary<string, Func<Dictionary<string, object>, Task<ToolResult>>> toolRegistry,
            object wordApplication)
        {
            toolRegistry["F_browser_snapshot"] = async (args) =>
            {
                try
                {
                    AgentRunCancellation.ThrowIfCancelled();

                    string channelId = TryGetString(args, "channel_id");
                    string refId = TryGetString(args, "ref");
                    string domSupplement = TryGetString(args, "dom_supplement");

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
                        return Fail("无默认渠道；请先 F_browser_navigate 打开网页，或传入 channel_id");
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

                    BrowserSnapshotResult result = await BrowserHostAdapter
                        .SnapshotAsync(channel, refId, domSupplement)
                        .ConfigureAwait(true);

                    if (!result.Success)
                    {
                        return Fail(result.Error ?? "snapshot 失败");
                    }

                    var data = new Dictionary<string, object>
                    {
                        ["channel_id"] = ChannelRegistry.ToPublicId(result.Channel.ChannelId),
                        ["kind"] = "browser",
                        ["track"] = result.Channel.Track,
                        ["url"] = result.Url ?? "",
                        ["title"] = result.Title ?? "",
                        ["mode"] = result.Mode ?? "overview",
                        ["truncated"] = result.Truncated,
                        ["snapshot"] = result.Snapshot ?? "",
                    };

                    if (!string.IsNullOrEmpty(result.Ref))
                    {
                        data["ref"] = result.Ref;
                    }

                    if (!string.IsNullOrEmpty(result.DomSupplement))
                    {
                        data["dom_supplement"] = result.DomSupplement;
                    }

                    if (result.Truncated && !string.IsNullOrEmpty(result.TruncatedReason))
                    {
                        data["truncated_reason"] = result.TruncatedReason;
                    }

                    return new ToolResult { Success = true, Data = data };
                }
                catch (OperationCanceledException)
                {
                    return Fail("cancelled by user");
                }
                catch (Exception ex)
                {
                    return Fail("F_browser_snapshot 失败: " + ex.Message);
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
