using System;
using System.Collections.Generic;

namespace WordAddIn1.BrowserHost
{
    /// <summary>
    /// 每渠道最近一次成功 snapshot 的 ref → 内部节点定位（不对 Agent 暴露）。
    /// </summary>
    internal static class BrowserRefStore
    {
        private static readonly object Gate = new object();
        private static readonly Dictionary<string, Dictionary<string, BrowserRefEntry>> ByChannel =
            new Dictionary<string, Dictionary<string, BrowserRefEntry>>(StringComparer.OrdinalIgnoreCase);

        public static void Replace(string channelId, IDictionary<string, BrowserRefEntry> map)
        {
            if (string.IsNullOrWhiteSpace(channelId))
            {
                return;
            }

            lock (Gate)
            {
                if (map == null || map.Count == 0)
                {
                    ByChannel.Remove(channelId);
                    return;
                }

                ByChannel[channelId] = new Dictionary<string, BrowserRefEntry>(map, StringComparer.OrdinalIgnoreCase);
            }
        }

        public static void Clear(string channelId)
        {
            if (string.IsNullOrWhiteSpace(channelId))
            {
                return;
            }

            lock (Gate)
            {
                ByChannel.Remove(channelId);
            }
        }

        public static bool TryGet(string channelId, string refId, out BrowserRefEntry entry)
        {
            entry = null;
            if (string.IsNullOrWhiteSpace(channelId) || string.IsNullOrWhiteSpace(refId))
            {
                return false;
            }

            lock (Gate)
            {
                if (!ByChannel.TryGetValue(channelId, out Dictionary<string, BrowserRefEntry> map)
                    || map == null)
                {
                    return false;
                }

                return map.TryGetValue(refId.Trim(), out entry) && entry != null;
            }
        }
    }

    internal sealed class BrowserRefEntry
    {
        public string AxNodeId { get; set; }

        public int? BackendDomNodeId { get; set; }

        public string Role { get; set; }

        public string Name { get; set; }
    }
}
