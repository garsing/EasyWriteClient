using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using WordAddIn1.BrowserHost;

namespace WordAddIn1
{
    /// <summary>
    /// 易写浏览页内交互：click / type / scroll / press / select。
    /// </summary>
    public static class F_BrowserInteractTool
    {
        private static readonly HashSet<string> AllowedKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "Enter", "Return", "Tab", "Escape", "Esc", "Backspace", "Delete",
            "ArrowUp", "ArrowDown", "ArrowLeft", "ArrowRight",
            "Home", "End", "PageUp", "PageDown"
        };

        public static void Register(
            Dictionary<string, Func<Dictionary<string, object>, Task<ToolResult>>> toolRegistry,
            object wordApplication)
        {
            toolRegistry["F_browser_interact"] = async (args) =>
            {
                try
                {
                    AgentRunCancellation.ThrowIfCancelled();

                    string action = TryGetString(args, "action");
                    if (string.IsNullOrWhiteSpace(action))
                    {
                        return Fail("必须提供 action（click|type|scroll|press|select）");
                    }

                    string channelId = TryGetString(args, "channel_id");
                    string refId = TryGetString(args, "ref");
                    string text = TryGetStringRaw(args, "text");
                    string direction = TryGetString(args, "direction");
                    string key = TryGetString(args, "key");
                    string option = TryGetString(args, "option");

                    string act = action.Trim().ToLowerInvariant();
                    if (act == "press" && !string.IsNullOrWhiteSpace(key) && !AllowedKeys.Contains(key.Trim()))
                    {
                        return Fail("key 不在白名单: " + key);
                    }

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
                        return Fail("默认渠道不是浏览器渠道；请传入 browser:agent: channel_id");
                    }
                    else
                    {
                        channel = browser;
                    }

                    if (!string.Equals(channel.Track, "agent", StringComparison.OrdinalIgnoreCase))
                    {
                        return Fail("本批仅支持 browser:agent: 渠道");
                    }

                    if (!channel.IsLive())
                    {
                        return Fail("浏览器渠道对应的页面已关闭: " + channel.ChannelId);
                    }

                    BrowserInteractResult result = await BrowserHostAdapter
                        .InteractAgentAsync(channel, act, refId, text, direction, key, option)
                        .ConfigureAwait(true);

                    if (!result.Success)
                    {
                        return Fail(result.Error ?? "interact 失败");
                    }

                    var data = new Dictionary<string, object>
                    {
                        ["channel_id"] = result.Channel.ChannelId,
                        ["kind"] = "browser",
                        ["track"] = result.Channel.Track,
                        ["action"] = result.Action ?? act,
                        ["url"] = result.Url ?? "",
                        ["title"] = result.Title ?? "",
                        ["refs_invalidated"] = result.RefsInvalidated,
                        ["message"] = result.Message ?? "",
                    };

                    if (!string.IsNullOrEmpty(result.Ref))
                    {
                        data["ref"] = result.Ref;
                    }

                    return new ToolResult { Success = true, Data = data };
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

        /// <summary>text 保留原样（仅 trim 首尾空白），失败路径不把全文写入错误。</summary>
        private static string TryGetStringRaw(Dictionary<string, object> args, string key)
        {
            if (args == null || !args.ContainsKey(key) || args[key] == null)
            {
                return null;
            }

            return Convert.ToString(args[key]);
        }

        private static ToolResult Fail(string error)
        {
            return new ToolResult { Success = false, Error = error };
        }
    }
}
