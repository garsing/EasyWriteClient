using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace WordAddIn1.BrowserHost
{
    /// <summary>
    /// 操作成功后按该层拍整棵 overview，Replace 新表后把树带回同一工具结果。
    /// </summary>
    internal static class BrowserResnapshotAfterAction
    {
        public const string UseNewRefsSuffix = "请使用本次返回的新 ref。";

        public const string MissedTreeHint =
            "操作已执行，但树没拍到（页面已跳转或 iframe 已卸载）。请空 F_browser_snapshot 看壳。";

        public const string DownloadMissedTreeHint =
            "已下载到会话工作区，但树没拍到（页面已跳转或 iframe 已卸载）。请空 F_browser_snapshot 看壳。";

        public static async Task<BrowserSnapshotResult> CaptureAsync(
            BrowserChannel channel,
            BrowserResnapshotLayer layer)
        {
            if (channel == null)
            {
                return BrowserSnapshotResult.Fail("通道无效。");
            }

            if (string.Equals(channel.Track, "attach", StringComparison.OrdinalIgnoreCase))
            {
                return await ExtensionHost.SnapshotLayerAsync(channel, layer).ConfigureAwait(true);
            }

            return await YiWriteBrowserHost.SnapshotLayerAsync(channel, layer).ConfigureAwait(true);
        }

        public static async Task<BrowserInteractResult> CompleteInteractAsync(
            BrowserChannel channel,
            string action,
            string refId,
            BrowserResnapshotLayer layer)
        {
            BrowserSnapshotResult snap = await TryCaptureAsync(channel, layer).ConfigureAwait(true);
            if (IsUsable(snap))
            {
                return BrowserInteractResult.OkResnapshot(
                    channel,
                    action,
                    refId,
                    SuccessMessage(action),
                    snap);
            }

            if (channel != null)
            {
                BrowserRefStore.Clear(channel.ChannelId);
            }

            return BrowserInteractResult.OkMissedTree(channel, action, refId, MissedTreeHint);
        }

        public static async Task<BrowserDownloadResult> CompleteDownloadAsync(
            BrowserChannel channel,
            string pageUrl,
            string title,
            string filename,
            long bytes,
            string contentType,
            string sourceUrl,
            string refId,
            BrowserResnapshotLayer layer)
        {
            BrowserSnapshotResult snap = await TryCaptureAsync(channel, layer).ConfigureAwait(true);
            if (IsUsable(snap))
            {
                return BrowserDownloadResult.OkResnapshot(
                    channel,
                    pageUrl,
                    title,
                    filename,
                    bytes,
                    contentType,
                    sourceUrl,
                    refId,
                    snap);
            }

            if (channel != null)
            {
                BrowserRefStore.Clear(channel.ChannelId);
            }

            return BrowserDownloadResult.OkMissedTree(
                channel,
                pageUrl,
                title,
                filename,
                bytes,
                contentType,
                sourceUrl,
                refId);
        }

        public static void WriteSnapshotFields(
            IDictionary<string, object> data,
            bool expandedFrame,
            bool resnapshot,
            string snapshot,
            string mode,
            bool truncated,
            string truncatedReason)
        {
            if (data == null || string.IsNullOrWhiteSpace(snapshot))
            {
                return;
            }

            data["snapshot"] = snapshot;
            data["mode"] = mode ?? "overview";
            data["truncated"] = truncated;
            if (!string.IsNullOrWhiteSpace(truncatedReason))
            {
                data["truncated_reason"] = truncatedReason;
            }

            if (expandedFrame)
            {
                data["expanded_frame"] = true;
            }

            if (resnapshot)
            {
                data["resnapshot"] = true;
            }
        }

        private static async Task<BrowserSnapshotResult> TryCaptureAsync(
            BrowserChannel channel,
            BrowserResnapshotLayer layer)
        {
            try
            {
                return await CaptureAsync(channel, layer ?? BrowserResnapshotLayer.Shell())
                    .ConfigureAwait(true);
            }
            catch (Exception ex)
            {
                return BrowserSnapshotResult.Fail(ex.Message);
            }
        }

        private static bool IsUsable(BrowserSnapshotResult snap)
        {
            return snap != null && snap.Success && !string.IsNullOrWhiteSpace(snap.Snapshot);
        }

        private static string SuccessMessage(string action)
        {
            string verb;
            switch ((action ?? "").Trim().ToLowerInvariant())
            {
                case "click":
                    verb = "已点击";
                    break;
                case "type":
                    verb = "已输入";
                    break;
                case "scroll":
                    verb = "已滚动";
                    break;
                case "press":
                    verb = "已按键";
                    break;
                case "select":
                    verb = "已选择";
                    break;
                default:
                    verb = "已操作";
                    break;
            }

            return verb + "。" + UseNewRefsSuffix;
        }
    }
}
