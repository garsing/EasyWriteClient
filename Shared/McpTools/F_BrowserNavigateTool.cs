using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using WordAddIn1.BrowserHost;

namespace WordAddIn1
{
    /// <summary>
    /// 打开/跳转易写浏览窗（WebView2），建立 browser:agent: 渠道。
    /// </summary>
    public static class F_BrowserNavigateTool
    {
        public static void Register(
            Dictionary<string, Func<Dictionary<string, object>, Task<ToolResult>>> toolRegistry,
            object wordApplication)
        {
            toolRegistry["F_browser_navigate"] = async (args) =>
            {
                try
                {
                    AgentRunCancellation.ThrowIfCancelled();

                    string url = TryGetString(args, "url");
                    if (string.IsNullOrWhiteSpace(url))
                    {
                        return Fail("必须提供 url 参数");
                    }

                    bool visible = ParseBool(args, "visible", true);
                    string channelId = TryGetString(args, "channel_id");

                    string existingTabUuid = null;
                    if (!string.IsNullOrWhiteSpace(channelId))
                    {
                        if (!ChannelRegistry.TryGetBrowser(channelId, out BrowserChannel existing))
                        {
                            return Fail("channel_id 不是浏览器渠道: " + channelId);
                        }

                        if (!string.Equals(existing.Track, "agent", StringComparison.OrdinalIgnoreCase))
                        {
                            return Fail("本批仅支持 browser:agent: 渠道");
                        }

                        if (!existing.IsLive())
                        {
                            return Fail("浏览器渠道对应的页面已关闭: " + channelId);
                        }

                        existingTabUuid = existing.TabUuid;
                    }
                    else if (ChannelRegistry.TryGetDefault(out IOperationChannel def)
                        && def is BrowserChannel defBrowser
                        && string.Equals(defBrowser.Track, "agent", StringComparison.OrdinalIgnoreCase)
                        && defBrowser.IsLive())
                    {
                        existingTabUuid = defBrowser.TabUuid;
                    }

                    BrowserNavigateResult result = await BrowserHostAdapter
                        .NavigateAgentAsync(url.Trim(), visible, existingTabUuid)
                        .ConfigureAwait(true);

                    if (!result.Success || result.Channel == null)
                    {
                        return Fail(result.Error ?? "打开失败");
                    }

                    return new ToolResult
                    {
                        Success = true,
                        Data = new Dictionary<string, object>
                        {
                            ["channel_id"] = result.Channel.ChannelId,
                            ["kind"] = "browser",
                            ["track"] = result.Channel.Track,
                            ["tab_uuid"] = result.Channel.TabUuid,
                            ["url"] = result.Url ?? "",
                            ["title"] = result.Title ?? "",
                            ["visible"] = result.Visible,
                        }
                    };
                }
                catch (OperationCanceledException)
                {
                    return Fail("cancelled by user");
                }
                catch (Exception ex)
                {
                    return Fail("F_browser_navigate 失败: " + ex.Message);
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

        private static bool ParseBool(Dictionary<string, object> args, string key, bool defaultValue)
        {
            if (args == null || !args.ContainsKey(key) || args[key] == null)
            {
                return defaultValue;
            }

            object raw = args[key];
            if (raw is bool b)
            {
                return b;
            }

            string s = Convert.ToString(raw);
            if (string.IsNullOrWhiteSpace(s))
            {
                return defaultValue;
            }

            if (bool.TryParse(s, out bool parsed))
            {
                return parsed;
            }

            if (s == "1" || string.Equals(s, "yes", StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            if (s == "0" || string.Equals(s, "no", StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }

            return defaultValue;
        }

        private static ToolResult Fail(string error)
        {
            return new ToolResult { Success = false, Error = error };
        }
    }
}
