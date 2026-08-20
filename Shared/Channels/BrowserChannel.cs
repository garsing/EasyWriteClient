using System;

namespace WordAddIn1
{
    /// <summary>
    /// 浏览器页渠道：一条逻辑页一条；track = agent（易写窗）| attach（Chrome/Edge 扩展）。
    /// </summary>
    public sealed class BrowserChannel : IOperationChannel
    {
        public const string AgentPrefix = "browser:agent:";
        public const string AttachPrefix = "browser:attach:";

        public BrowserChannel(string channelId, string tabUuid, string track)
        {
            if (string.IsNullOrWhiteSpace(channelId))
            {
                throw new ArgumentException("channel_id 不能为空。", nameof(channelId));
            }

            if (string.IsNullOrWhiteSpace(tabUuid))
            {
                throw new ArgumentException("tab_uuid 不能为空。", nameof(tabUuid));
            }

            ChannelId = channelId.Trim();
            TabUuid = tabUuid.Trim();
            Track = string.IsNullOrWhiteSpace(track) ? "agent" : track.Trim();
        }

        public string ChannelId { get; }

        public ChannelKind Kind => ChannelKind.Browser;

        public string TabUuid { get; }

        /// <summary>agent | attach</summary>
        public string Track { get; }

        public string Url { get; private set; } = "";

        public string Title { get; private set; } = "";

        public bool Visible { get; private set; }

        public void UpdatePage(string url, string title, bool visible)
        {
            Url = url ?? "";
            Title = title ?? "";
            Visible = visible;
        }

        public bool IsLive()
        {
            if (string.Equals(Track, "attach", StringComparison.OrdinalIgnoreCase))
            {
                return BrowserHost.BrowserHostAdapter.IsAttachPageLive(TabUuid);
            }

            return BrowserHost.BrowserHostAdapter.IsAgentPageLive(TabUuid);
        }

        public bool TryActivate()
        {
            if (string.Equals(Track, "attach", StringComparison.OrdinalIgnoreCase))
            {
                return BrowserHost.BrowserHostAdapter.TryShowAttachPage(TabUuid);
            }

            return BrowserHost.BrowserHostAdapter.TryShowAgentPage(TabUuid);
        }

        public string TryGetDisplayName()
        {
            if (!string.IsNullOrWhiteSpace(Title))
            {
                return Title;
            }

            return string.IsNullOrWhiteSpace(Url) ? null : Url;
        }
    }
}
