using System;
using System.Threading.Tasks;

namespace WordAddIn1.BrowserHost
{
    /// <summary>
    /// 浏览器适配层入口：navigate 按 track（页面在哪）。
    /// snapshot / interact / download / 截图后端按 ClickMode；track 只选传输。
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
            string domSupplement = null,
            int? rangeStart = null,
            int? rangeEnd = null)
        {
            if (string.IsNullOrWhiteSpace(refId))
            {
                if (!BrowserDomInputSupplement.TryResolveSelector(
                    domSupplement,
                    out _,
                    out string normalized,
                    out string domErr))
                {
                    return Task.FromResult(
                        BrowserSnapshotResult.Fail(domErr ?? "dom_supplement 无效"));
                }

                domSupplement = normalized;
            }

            BrowserOpMode mode = BrowserOpMode.Resolve(channel);
            if (mode.IsPointer)
            {
                return mode.IsAttach
                    ? AttachPointerHost.SnapshotAsync(channel, refId, domSupplement, rangeStart, rangeEnd)
                    : YiWriteBrowserHost.SnapshotAsync(channel, refId, domSupplement, rangeStart, rangeEnd);
            }

            if (rangeStart.HasValue || rangeEnd.HasValue)
            {
                return Task.FromResult(
                    BrowserSnapshotResult.Fail("当前浏览器轨不支持序号范围"));
            }

            return mode.IsAttach
                ? ExtensionHost.SnapshotAsync(channel, refId, domSupplement)
                : AgentLegacyHost.SnapshotAsync(channel, refId, domSupplement);
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
            BrowserOpMode mode = BrowserOpMode.Resolve(channel);
            if (mode.IsPointer)
            {
                return mode.IsAttach
                    ? AttachPointerHost.InteractAsync(
                        channel, action, refId, text, direction, key, option)
                    : YiWriteBrowserHost.InteractAsync(
                        channel, action, refId, text, direction, key, option);
            }

            return mode.IsAttach
                ? ExtensionHost.InteractAsync(
                    channel, action, refId, text, direction, key, option)
                : AgentLegacyHost.InteractAsync(
                    channel, action, refId, text, direction, key, option);
        }

        public static Task<BrowserDownloadResult> DownloadAsync(
            BrowserChannel channel,
            string refId,
            string url)
        {
            BrowserOpMode mode = BrowserOpMode.Resolve(channel);
            if (mode.IsPointer)
            {
                return mode.IsAttach
                    ? AttachPointerHost.DownloadAsync(channel, refId, url)
                    : YiWriteBrowserHost.DownloadAsync(channel, refId, url);
            }

            return mode.IsAttach
                ? ExtensionHost.DownloadAsync(channel, refId, url)
                : AgentLegacyHost.DownloadAsync(channel, refId, url);
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
