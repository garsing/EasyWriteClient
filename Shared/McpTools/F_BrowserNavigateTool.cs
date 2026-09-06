using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using WordAddIn1.BrowserHost;

namespace WordAddIn1
{
    /// <summary>
    /// 打开/跳转：agent=易写浏览窗；attach=同标签跳转（Chrome/Edge）。
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
                    string host = TryGetString(args, "host");
                    if (string.IsNullOrWhiteSpace(host))
                    {
                        host = TryGetString(args, "browser");
                    }

                    // host: yiwrite | chrome | edge（省略=按渠道/默认，否则易写窗）
                    string hostNorm = string.IsNullOrWhiteSpace(host)
                        ? null
                        : host.Trim().ToLowerInvariant();
                    if (hostNorm == "yiwrite" || hostNorm == "agent" || hostNorm == "webview")
                    {
                        hostNorm = "yiwrite";
                    }
                    else if (hostNorm == "chrome" || hostNorm == "edge")
                    {
                        /* ok */
                    }
                    else if (hostNorm != null)
                    {
                        return Fail("host 仅支持 yiwrite | chrome | edge");
                    }

                    if ((hostNorm == "chrome" || hostNorm == "edge") && !visible)
                    {
                        return Fail("visible=false 时只能用易写浏览窗（host=yiwrite），不能拉起 Chrome/Edge");
                    }

                    BrowserChannel attachTarget = null;
                    string existingAgentTabUuid = null;

                    if (!string.IsNullOrWhiteSpace(channelId))
                    {
                        if (!ChannelRegistry.TryGetBrowser(channelId, out BrowserChannel existing))
                        {
                            return Fail("channel_id 不是浏览器渠道: " + channelId);
                        }

                        if (string.Equals(existing.Track, "attach", StringComparison.OrdinalIgnoreCase))
                        {
                            if (!visible)
                            {
                                return Fail("visible=false 不可用于附着 Chrome/Edge；请用易写浏览窗");
                            }

                            if (!existing.IsLive())
                            {
                                return Fail("附着标签已断开: " + channelId);
                            }

                            attachTarget = existing;
                        }
                        else if (string.Equals(existing.Track, "agent", StringComparison.OrdinalIgnoreCase))
                        {
                            if (!existing.IsLive())
                            {
                                return Fail("浏览器渠道对应的页面已关闭: " + channelId);
                            }

                            existingAgentTabUuid = existing.TabUuid;
                        }
                        else
                        {
                            return Fail("不支持的浏览器 track: " + existing.Track);
                        }
                    }
                    else if (hostNorm == null
                        && ChannelRegistry.TryGetDefault(out IOperationChannel def)
                        && def is BrowserChannel defBrowser
                        && defBrowser.IsLive())
                    {
                        if (string.Equals(defBrowser.Track, "attach", StringComparison.OrdinalIgnoreCase))
                        {
                            if (!visible)
                            {
                                return Fail("visible=false 不可用于附着 Chrome/Edge；请用易写浏览窗");
                            }

                            attachTarget = defBrowser;
                        }
                        else if (string.Equals(defBrowser.Track, "agent", StringComparison.OrdinalIgnoreCase))
                        {
                            existingAgentTabUuid = defBrowser.TabUuid;
                        }
                    }

                    BrowserNavigateResult result;
                    if (attachTarget != null && hostNorm != "yiwrite")
                    {
                        // 已有 attach 渠且未强制易写窗：同标签跳转
                        result = await BrowserHostAdapter
                            .NavigateAttachAsync(attachTarget, url.Trim())
                            .ConfigureAwait(true);
                    }
                    else if (hostNorm == "chrome" || hostNorm == "edge")
                    {
                        result = await BrowserHostAdapter
                            .LaunchAttachNavigateAsync(hostNorm, url.Trim())
                            .ConfigureAwait(true);
                    }
                    else
                    {
                        result = await BrowserHostAdapter
                            .NavigateAgentAsync(url.Trim(), visible, existingAgentTabUuid)
                            .ConfigureAwait(true);
                    }

                    if (!result.Success || result.Channel == null)
                    {
                        return Fail(result.Error ?? "打开失败");
                    }

                    return new ToolResult
                    {
                        Success = true,
                        Data = new Dictionary<string, object>
                        {
                            ["channel_id"] = ChannelRegistry.ToPublicId(result.Channel.ChannelId),
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

        private static ToolResult Fail(string message)
        {
            return new ToolResult { Success = false, Error = message };
        }
    }
}
