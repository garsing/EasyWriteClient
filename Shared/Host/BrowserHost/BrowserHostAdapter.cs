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

        internal static async Task<BrowserCaptureResult> CaptureLocalAsync(
            BrowserChannel channel,
            string refId,
            int up)
        {
            if (channel != null
                && string.Equals(channel.Track, "attach", StringComparison.OrdinalIgnoreCase))
            {
                var cdp = new AttachCdpSession(channel.TabUuid, channel.Url, channel.Title);
                return await BrowserLocalCapture.CaptureAsync(cdp, channel, refId, up).ConfigureAwait(true);
            }

            if (!YiWriteBrowserHost.TryPrepareFormForTab(channel?.TabUuid, out YiWriteBrowserForm form, out string prepError))
            {
                return BrowserCaptureResult.Fail(prepError ?? "引擎不可用");
            }

            try
            {
                return await BrowserLocalCapture.CaptureAsync(form, channel, refId, up).ConfigureAwait(true);
            }
            finally
            {
                form.UnpinToolPage();
            }
        }

        public static async Task<BrowserSnapshotResult> SnapshotAsync(
            BrowserChannel channel,
            string refId,
            string domSupplement = null,
            int? rangeStart = null,
            int? rangeEnd = null,
            bool hiddenChildren = false,
            bool clear = false)
        {
            if (clear && !string.IsNullOrWhiteSpace(refId))
            {
                return await ClearThenExpandAsync(channel, refId).ConfigureAwait(true);
            }

            if (string.IsNullOrWhiteSpace(refId))
            {
                if (!BrowserDomInputSupplement.TryResolveSelector(
                    domSupplement,
                    out _,
                    out string normalized,
                    out string domErr))
                {
                    return BrowserSnapshotResult.Fail(domErr ?? "dom_supplement 无效");
                }

                domSupplement = normalized;
            }

            if (clear)
            {
                if (channel != null)
                {
                    BrowserFrameMemory.Forget(channel.ChannelId);
                }
            }
            else if (string.IsNullOrWhiteSpace(refId)
                && channel != null
                && BrowserFrameMemory.TryGet(channel.ChannelId, out BrowserResnapshotLayer remembered))
            {
                bool? alive = await BrowserFrameMemory
                    .StillThereAsync(channel, remembered)
                    .ConfigureAwait(true);
                if (alive == false)
                {
                    BrowserFrameMemory.Forget(channel.ChannelId);
                }
                else
                {
                    BrowserSnapshotResult layered = await BrowserResnapshotAfterAction
                        .CaptureAsync(channel, remembered)
                        .ConfigureAwait(true);
                    if (layered != null && layered.Success)
                    {
                        return layered;
                    }

                    return layered ?? BrowserSnapshotResult.Fail("取无障碍树失败");
                }
            }

            BrowserOpMode mode = BrowserOpMode.Resolve(channel);
            if (mode.IsPointer)
            {
                return mode.IsAttach
                    ? await AttachPointerHost.SnapshotAsync(channel, refId, domSupplement, rangeStart, rangeEnd, hiddenChildren).ConfigureAwait(true)
                    : await YiWriteBrowserHost.SnapshotAsync(channel, refId, domSupplement, rangeStart, rangeEnd, hiddenChildren).ConfigureAwait(true);
            }

            return mode.IsAttach
                ? await ExtensionHost.SnapshotAsync(channel, refId, domSupplement, rangeStart, rangeEnd, hiddenChildren).ConfigureAwait(true)
                : await AgentLegacyHost.SnapshotAsync(channel, refId, domSupplement, rangeStart, rangeEnd, hiddenChildren).ConfigureAwait(true);
        }

        /// <summary>先从当前 ref 表取出 iframe 行，再忘掉记忆，然后拍子文档。失败时记忆保持已清。</summary>
        private static async Task<BrowserSnapshotResult> ClearThenExpandAsync(
            BrowserChannel channel,
            string refId)
        {
            if (channel == null)
            {
                return BrowserSnapshotResult.Fail("无浏览器渠道");
            }

            string id = refId.Trim();
            if (!BrowserRefStore.TryGet(channel.ChannelId, id, out BrowserRefEntry prior) || prior == null)
            {
                BrowserFrameMemory.Forget(channel.ChannelId);
                return BrowserSnapshotResult.Fail("ref 无效或已过期，请重新 F_browser_snapshot");
            }

            if (!BrowserRiskGuard.IsFrameRole(prior))
            {
                BrowserFrameMemory.Forget(channel.ChannelId);
                return BrowserSnapshotResult.Fail("该 ref 不是 iframe 结构行");
            }

            BrowserResnapshotLayer layer = BrowserFrameMemory.LayerForIframeEntry(channel, prior);
            if (layer == null)
            {
                BrowserFrameMemory.Forget(channel.ChannelId);
                return BrowserSnapshotResult.Fail("无法读取 iframe（无子 frame）");
            }

            BrowserFrameMemory.Forget(channel.ChannelId);
            BrowserSnapshotResult expanded = await BrowserResnapshotAfterAction
                .CaptureAsync(channel, layer)
                .ConfigureAwait(true);
            if (expanded != null && expanded.Success)
            {
                BrowserFrameMemory.Remember(channel.ChannelId, layer);
                return expanded;
            }

            return expanded ?? BrowserSnapshotResult.Fail("无法读取 iframe 内容");
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
            return InteractCoreAsync(channel, action, refId, text, direction, key, option, true);
        }

        public static Task<BrowserInteractResult> InteractStepAsync(
            BrowserChannel channel,
            string action,
            string refId,
            string text,
            string direction,
            string key,
            string option)
        {
            return InteractCoreAsync(channel, action, refId, text, direction, key, option, false);
        }

        private static Task<BrowserInteractResult> InteractCoreAsync(
            BrowserChannel channel,
            string action,
            string refId,
            string text,
            string direction,
            string key,
            string option,
            bool completeAfter)
        {
            BrowserOpMode mode = BrowserOpMode.Resolve(channel);
            if (mode.IsPointer)
            {
                return mode.IsAttach
                    ? AttachPointerHost.InteractAsync(
                        channel, action, refId, text, direction, key, option, completeAfter)
                    : YiWriteBrowserHost.InteractAsync(
                        channel, action, refId, text, direction, key, option, completeAfter);
            }

            return mode.IsAttach
                ? ExtensionHost.InteractAsync(
                    channel, action, refId, text, direction, key, option, completeAfter)
                : AgentLegacyHost.InteractAsync(
                    channel, action, refId, text, direction, key, option, completeAfter);
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
