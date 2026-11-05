using System.Collections.Concurrent;
using System.Threading.Tasks;

namespace WordAddIn1.BrowserHost
{
    /// <summary>
    /// 每个渠道记住当前展开到的那一层文档。不放进 <see cref="BrowserRefStore"/>：
    /// 换 ref 表会删掉渠道条目，记忆要单独留下来。进程内，不落库，不把 frameId 交给模型。
    /// </summary>
    internal static class BrowserFrameMemory
    {
        private static readonly ConcurrentDictionary<string, BrowserResnapshotLayer> Layers =
            new ConcurrentDictionary<string, BrowserResnapshotLayer>();

        public static void Remember(string channelId, BrowserResnapshotLayer layer)
        {
            if (string.IsNullOrWhiteSpace(channelId) || layer == null)
            {
                return;
            }

            if (string.IsNullOrWhiteSpace(layer.FrameId)
                && string.IsNullOrWhiteSpace(layer.CdpSessionId)
                && !layer.AttachFrameId.HasValue)
            {
                return;
            }

            if (Layers.TryGetValue(channelId, out BrowserResnapshotLayer old) && Same(old, layer))
            {
                return;
            }

            Layers[channelId] = layer;
        }

        public static bool TryGet(string channelId, out BrowserResnapshotLayer layer)
        {
            layer = null;
            if (string.IsNullOrWhiteSpace(channelId))
            {
                return false;
            }

            return Layers.TryGetValue(channelId, out layer) && layer != null;
        }

        public static void Forget(string channelId)
        {
            if (string.IsNullOrWhiteSpace(channelId))
            {
                return;
            }

            BrowserResnapshotLayer removed;
            Layers.TryRemove(channelId, out removed);
        }

        /// <summary>
        /// 从还没被清掉的 ref 表取出 iframe 行的子文档。指针轨用 CDP frameId；
        /// 扩展附着用 AttachFrameId；legacy 没有子帧时退到当时的 session。
        /// </summary>
        public static BrowserResnapshotLayer LayerForIframeEntry(BrowserChannel channel, BrowserRefEntry prior)
        {
            if (channel == null || prior == null)
            {
                return null;
            }

            BrowserOpMode mode = BrowserOpMode.Resolve(channel);
            if (mode.IsPointer)
            {
                if (string.IsNullOrWhiteSpace(prior.ChildFrameId))
                {
                    return null;
                }

                return new BrowserResnapshotLayer
                {
                    FrameId = prior.ChildFrameId.Trim(),
                    CdpSessionId = prior.CdpSessionId
                };
            }

            if (mode.IsAttach)
            {
                if (!prior.AttachFrameId.HasValue)
                {
                    return null;
                }

                return new BrowserResnapshotLayer
                {
                    AttachFrameId = prior.AttachFrameId
                };
            }

            if (!string.IsNullOrWhiteSpace(prior.ChildFrameId))
            {
                return new BrowserResnapshotLayer
                {
                    FrameId = prior.ChildFrameId.Trim(),
                    CdpSessionId = prior.CdpSessionId
                };
            }

            if (!string.IsNullOrWhiteSpace(prior.CdpSessionId))
            {
                return new BrowserResnapshotLayer
                {
                    CdpSessionId = prior.CdpSessionId
                };
            }

            return null;
        }

        /// <summary>
        /// true 帧还在；false 帧树已读到且没有这一帧；null 读不到帧树（超时等），调用方不要据此清记忆。
        /// 扩展附着的 AttachFrameId 不是 CDP frameId，对不上帧树，返回 null。
        /// </summary>
        public static async Task<bool?> StillThereAsync(BrowserChannel channel, BrowserResnapshotLayer layer)
        {
            if (channel == null || layer == null || string.IsNullOrWhiteSpace(layer.FrameId))
            {
                return null;
            }

            BrowserOpMode mode = BrowserOpMode.Resolve(channel);
            if (mode.IsAttach && mode.IsPointer)
            {
                return await AttachPointerHost.FrameListedAsync(channel, layer.FrameId).ConfigureAwait(true);
            }

            if (!mode.IsAttach)
            {
                return await YiWriteBrowserHost.FrameListedAsync(channel, layer.FrameId).ConfigureAwait(true);
            }

            return null;
        }

        public static async Task ForgetIfAbsentAsync(BrowserChannel channel)
        {
            if (channel == null || !TryGet(channel.ChannelId, out BrowserResnapshotLayer layer))
            {
                return;
            }

            bool? alive = await StillThereAsync(channel, layer).ConfigureAwait(true);
            if (alive == false)
            {
                Forget(channel.ChannelId);
            }
        }

        private static bool Same(BrowserResnapshotLayer left, BrowserResnapshotLayer right)
        {
            if (left == null || right == null)
            {
                return false;
            }

            return string.Equals(left.FrameId ?? "", right.FrameId ?? "", System.StringComparison.Ordinal)
                && string.Equals(left.CdpSessionId ?? "", right.CdpSessionId ?? "", System.StringComparison.Ordinal)
                && left.AttachFrameId == right.AttachFrameId;
        }
    }
}
