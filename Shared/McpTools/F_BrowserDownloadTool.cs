using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using WordAddIn1.BrowserHost;

namespace WordAddIn1
{
    /// <summary>
    /// 将易写浏览页下载落到会话工作区（ref 真下载 或 url 直取）。
    /// </summary>
    public static class F_BrowserDownloadTool
    {
        public static void Register(
            Dictionary<string, Func<Dictionary<string, object>, Task<ToolResult>>> toolRegistry,
            object wordApplication)
        {
            toolRegistry["F_browser_download"] = async (args) =>
            {
                try
                {
                    AgentRunCancellation.ThrowIfCancelled();

                    string channelId = TryGetString(args, "channel_id");
                    string refId = TryGetString(args, "ref");
                    string url = TryGetString(args, "url");

                    bool hasRef = !string.IsNullOrWhiteSpace(refId);
                    bool hasUrl = !string.IsNullOrWhiteSpace(url);
                    if (hasRef == hasUrl)
                    {
                        return Fail("须且仅能提供 ref 或 url 之一");
                    }

                    if (!FilePathResolver.TryResolveFromArgs(args, out ResolvedFilePath dest, out string destError, "path"))
                    {
                        return Fail(destError ?? "必须提供 path");
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
                        return Fail("无默认渠道；请先 F_browser_navigate，或传入 channel_id");
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

                    BrowserDownloadResult result = await BrowserHostAdapter
                        .DownloadAsync(channel, refId, url)
                        .ConfigureAwait(true);

                    if (!result.Success)
                    {
                        return Fail(result.Error ?? "download 失败");
                    }

                    ResolvedFilePath downloaded = null;
                    if (!string.IsNullOrEmpty(result.Filename))
                    {
                        if (!FilePathResolver.TryResolve(result.Filename, out downloaded, out _)
                            || !File.Exists(downloaded.LocalPath))
                        {
                            FilePathResolver.TryResolve(
                                "browser_dl/" + Path.GetFileName(result.Filename),
                                out downloaded,
                                out _);
                        }
                    }

                    if (downloaded != null
                        && File.Exists(downloaded.LocalPath)
                        && !string.Equals(downloaded.LocalPath, dest.LocalPath, StringComparison.OrdinalIgnoreCase))
                    {
                        var copied = await FilePathResolver
                            .WriteBytesAsync(dest, File.ReadAllBytes(downloaded.LocalPath))
                            .ConfigureAwait(true);
                        if (!copied.Success)
                        {
                            return Fail(copied.Error);
                        }
                    }
                    else if (!File.Exists(dest.LocalPath))
                    {
                        return Fail("下载完成但无法落到指定 path");
                    }

                    var data = new Dictionary<string, object>
                    {
                        ["channel_id"] = ChannelRegistry.ToPublicId(result.Channel.ChannelId),
                        ["kind"] = "browser",
                        ["track"] = result.Channel.Track,
                        ["url"] = result.Url ?? "",
                        ["title"] = result.Title ?? "",
                        ["path"] = dest.Display,
                        ["filename"] = dest.Display,
                        ["bytes"] = result.Bytes,
                        ["refs_invalidated"] = result.RefsInvalidated,
                        ["message"] = result.Message ?? "",
                    };

                    if (!string.IsNullOrEmpty(result.ContentType))
                    {
                        data["content_type"] = result.ContentType;
                    }

                    if (!string.IsNullOrEmpty(result.SourceUrl))
                    {
                        data["source_url"] = result.SourceUrl;
                    }

                    if (!string.IsNullOrEmpty(result.Ref))
                    {
                        data["ref"] = result.Ref;
                    }

                    BrowserResnapshotAfterAction.WriteSnapshotFields(
                        data,
                        expandedFrame: false,
                        result.Resnapshot,
                        result.Snapshot,
                        result.Mode,
                        result.Truncated,
                        result.TruncatedReason);

                    return new ToolResult { Success = true, Data = data };
                }
                catch (OperationCanceledException)
                {
                    return Fail("cancelled by user");
                }
                catch (Exception ex)
                {
                    return Fail("F_browser_download 失败: " + ex.Message);
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
