using System;
using System.Collections.Generic;

namespace WordAddIn1
{
    /// <summary>
    /// 进程级操作渠道注册表。与 conversation 不强绑定（D15）。
    /// B0：骨架；B2 起接入 WordChannel / Document。
    /// </summary>
    public static class ChannelRegistry
    {
        private static readonly object Gate = new object();
        private static readonly Dictionary<string, IOperationChannel> Channels =
            new Dictionary<string, IOperationChannel>(StringComparer.Ordinal);
        private static string _defaultChannelId;

        public static string DefaultChannelId
        {
            get
            {
                lock (Gate)
                {
                    return _defaultChannelId;
                }
            }
        }

        public static void Register(IOperationChannel channel, bool setAsDefault = false)
        {
            if (channel == null)
            {
                throw new ArgumentNullException(nameof(channel));
            }

            if (string.IsNullOrWhiteSpace(channel.ChannelId))
            {
                throw new ArgumentException("channel_id 不能为空。", nameof(channel));
            }

            lock (Gate)
            {
                Channels[channel.ChannelId] = channel;
                if (setAsDefault || string.IsNullOrEmpty(_defaultChannelId))
                {
                    _defaultChannelId = channel.ChannelId;
                }
            }
        }

        public static bool TryGet(string channelId, out IOperationChannel channel)
        {
            channel = null;
            if (string.IsNullOrWhiteSpace(channelId))
            {
                return false;
            }

            lock (Gate)
            {
                return Channels.TryGetValue(channelId, out channel);
            }
        }

        public static bool SetDefault(string channelId)
        {
            if (string.IsNullOrWhiteSpace(channelId))
            {
                return false;
            }

            lock (Gate)
            {
                if (!Channels.ContainsKey(channelId))
                {
                    return false;
                }

                _defaultChannelId = channelId;
                return true;
            }
        }

        public static bool TryGetDefault(out IOperationChannel channel)
        {
            channel = null;
            lock (Gate)
            {
                if (string.IsNullOrEmpty(_defaultChannelId))
                {
                    return false;
                }

                return Channels.TryGetValue(_defaultChannelId, out channel);
            }
        }

        public static bool Remove(string channelId)
        {
            if (string.IsNullOrWhiteSpace(channelId))
            {
                return false;
            }

            lock (Gate)
            {
                bool removed = Channels.Remove(channelId);
                if (removed && string.Equals(_defaultChannelId, channelId, StringComparison.Ordinal))
                {
                    _defaultChannelId = null;
                }

                return removed;
            }
        }

        /// <summary>测试/进程退出用。</summary>
        public static void ClearAll()
        {
            lock (Gate)
            {
                Channels.Clear();
                _defaultChannelId = null;
            }
        }
    }
}
