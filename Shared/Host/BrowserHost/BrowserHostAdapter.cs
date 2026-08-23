using System;
using System.Threading.Tasks;

namespace WordAddIn1.BrowserHost
{
    /// <summary>
    /// 浏览器适配层入口：按渠道 track 分流 agent（WebView2）/ attach（扩展）。
    /// </summary>
    public static class BrowserHostAdapter
    {
        public static void EnsureAttachHostStarted()
        {
            ExtensionHost.EnsureStarted();
        }

        public static bool IsAgentPageLive(string tabUuid)
        {
            return YiWriteBrowserHost.IsPageLive(tabUuid);
        }

        public static bool IsAttachPageLive(string tabUuid)
        {
            return ExtensionHost.IsAttachTabLive(tabUuid);
        }

        public static bool TryShowAgentPage(string tabUuid)
        {
            return YiWriteBrowserHost.TryShowPage(tabUuid);
        }

        public static bool TryShowAttachPage(string tabUuid)
        {
            return ExtensionHost.TryActivateAttachTab(tabUuid);
        }

        public static Task<BrowserNavigateResult> NavigateAgentAsync(
            string url,
            bool visible,
            string existingTabUuid)
        {
            return YiWriteBrowserHost.NavigateAsync(url, visible, existingTabUuid);
        }

        public static Task<BrowserNavigateResult> NavigateAttachAsync(BrowserChannel channel, string url)
        {
            return ExtensionHost.NavigateAsync(channel, url);
        }

        public static Task<BrowserNavigateResult> LaunchAttachNavigateAsync(string host, string url)
        {
            return ExtensionHost.LaunchNavigateAsync(host, url);
        }

        internal static Task<BrowserCaptureResult> CaptureViewportAsync(BrowserChannel channel)
        {
            if (channel != null
                && string.Equals(channel.Track, "attach", StringComparison.OrdinalIgnoreCase))
            {
                return ExtensionHost.CaptureViewportAsync(channel);
            }

            return YiWriteBrowserHost.CaptureViewportAsync(channel);
        }

        public static Task<BrowserSnapshotResult> SnapshotAsync(
            BrowserChannel channel,
            string refId,
            string domSupplement = null)
        {
            if (channel != null
                && string.Equals(channel.Track, "attach", StringComparison.OrdinalIgnoreCase))
            {
                return ExtensionHost.SnapshotAsync(channel, refId, domSupplement);
            }

            return YiWriteBrowserHost.SnapshotAsync(channel, refId, domSupplement);
        }

        public static Task<BrowserInteractResult> InteractAsync(
            BrowserChannel channel,
            string action,
            string refId,
            string text,
            string direction,
            string key,
            string option)
        {
            if (channel != null
                && string.Equals(channel.Track, "attach", StringComparison.OrdinalIgnoreCase))
            {
                return ExtensionHost.InteractAsync(
                    channel, action, refId, text, direction, key, option);
            }

            return YiWriteBrowserHost.InteractAsync(
                channel, action, refId, text, direction, key, option);
        }

        public static Task<BrowserDownloadResult> DownloadAsync(
            BrowserChannel channel,
            string refId,
            string url)
        {
            if (channel != null
                && string.Equals(channel.Track, "attach", StringComparison.OrdinalIgnoreCase))
            {
                return ExtensionHost.DownloadAsync(channel, refId, url);
            }

            return YiWriteBrowserHost.DownloadAsync(channel, refId, url);
        }

        // 兼容旧调用名
        public static Task<BrowserSnapshotResult> SnapshotAgentAsync(
            BrowserChannel channel,
            string refId,
            string domSupplement = null) =>
            SnapshotAsync(channel, refId, domSupplement);

        public static Task<BrowserInteractResult> InteractAgentAsync(
            BrowserChannel channel,
            string action,
            string refId,
            string text,
            string direction,
            string key,
            string option) =>
            InteractAsync(channel, action, refId, text, direction, key, option);

        public static Task<BrowserDownloadResult> DownloadAgentAsync(
            BrowserChannel channel,
            string refId,
            string url) =>
            DownloadAsync(channel, refId, url);
    }
}
