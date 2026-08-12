using System;
using System.Collections.Generic;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    /// <summary>
    /// 进程级操作渠道注册表。与 conversation 不强绑定（D15）。
    /// </summary>
    public static class ChannelRegistry
    {
        private static readonly object Gate = new object();
        private static readonly Dictionary<string, IOperationChannel> Channels =
            new Dictionary<string, IOperationChannel>(StringComparer.Ordinal);
        private static readonly Dictionary<string, string> DocUuidToChannelId =
            new Dictionary<string, string>(StringComparer.Ordinal);
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

        /// <summary>
        /// 为 Word 文档查找或创建渠道；channel_id 形如 <c>word:{doc_uuid}</c>。
        /// </summary>
        /// <param name="claimDefaultIfEmpty">
        /// 默认渠道为空时是否自动占用（默认 true，保持 Plugin / 旧工具行为）。
        /// 「打开文件」探测路径须传 false，避免抢默认渠道。
        /// </param>
        public static WordChannel CreateOrGetWord(
            Word.Document doc,
            string filePath = null,
            bool claimDefaultIfEmpty = true)
        {
            if (doc == null)
            {
                throw new ArgumentNullException(nameof(doc));
            }

            string uuid = DocumentIdentity.EnsureUuid(doc);
            DocumentIdentity.EnsureCloseHandler(doc);

            bool createdNew = false;
            WordChannel result;
            lock (Gate)
            {
                if (DocUuidToChannelId.TryGetValue(uuid, out string existingId)
                    && Channels.TryGetValue(existingId, out IOperationChannel existing)
                    && existing is WordChannel wordChannel)
                {
                    wordChannel.UpdateDocument(doc, filePath);
                    result = wordChannel;
                }
                else
                {
                    string channelId = "word:" + uuid;
                    var created = new WordChannel(channelId, uuid, doc, filePath);
                    Channels[channelId] = created;
                    DocUuidToChannelId[uuid] = channelId;

                    if (claimDefaultIfEmpty && string.IsNullOrEmpty(_defaultChannelId))
                    {
                        _defaultChannelId = channelId;
                    }

                    createdNew = true;
                    result = created;
                }
            }

            // 锁外触发：Desktop 新建渠道 → 自动缩小（幂等）
            if (createdNew)
            {
                HostCallbacks.RaiseRequestCompact();
            }

            return result;
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
                if (channel is WordChannel wc && !string.IsNullOrEmpty(wc.DocUuid))
                {
                    DocUuidToChannelId[wc.DocUuid] = channel.ChannelId;
                }

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

        public static bool TryGetWord(string channelId, out WordChannel channel)
        {
            channel = null;
            if (!TryGet(channelId, out IOperationChannel ch) || !(ch is WordChannel wc))
            {
                return false;
            }

            channel = wc;
            return true;
        }

        public static bool TryGetByDocUuid(string docUuid, out WordChannel channel)
        {
            channel = null;
            if (string.IsNullOrWhiteSpace(docUuid))
            {
                return false;
            }

            lock (Gate)
            {
                if (!DocUuidToChannelId.TryGetValue(docUuid, out string channelId))
                {
                    return false;
                }

                if (!Channels.TryGetValue(channelId, out IOperationChannel ch) || !(ch is WordChannel wc))
                {
                    return false;
                }

                channel = wc;
                return true;
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

        public static IOperationChannel GetDefaultOrNull()
        {
            return TryGetDefault(out IOperationChannel channel) ? channel : null;
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

        public static bool TryGetDefaultWord(out WordChannel channel)
        {
            channel = null;
            if (!TryGetDefault(out IOperationChannel ch) || !(ch is WordChannel wc))
            {
                return false;
            }

            channel = wc;
            return true;
        }

        /// <summary>
        /// Plugin：活动文档变化时查找/创建 Word 渠道并设为默认（§4.3）。
        /// </summary>
        public static WordChannel SyncDefaultFromActiveDocument(Word.Document document)
        {
            if (document == null)
            {
                return null;
            }

            WordChannel channel = CreateOrGetWord(document);
            SetDefault(channel.ChannelId);
            return channel;
        }

        public static bool Remove(string channelId)
        {
            if (string.IsNullOrWhiteSpace(channelId))
            {
                return false;
            }

            lock (Gate)
            {
                if (!Channels.TryGetValue(channelId, out IOperationChannel ch))
                {
                    return false;
                }

                Channels.Remove(channelId);
                if (ch is WordChannel wc && !string.IsNullOrEmpty(wc.DocUuid))
                {
                    DocUuidToChannelId.Remove(wc.DocUuid);
                }

                if (string.Equals(_defaultChannelId, channelId, StringComparison.Ordinal))
                {
                    _defaultChannelId = null;
                }

                return true;
            }
        }

        public static bool RemoveByDocUuid(string docUuid)
        {
            if (string.IsNullOrWhiteSpace(docUuid))
            {
                return false;
            }

            string channelId;
            lock (Gate)
            {
                if (!DocUuidToChannelId.TryGetValue(docUuid, out channelId))
                {
                    return false;
                }
            }

            return Remove(channelId);
        }

        /// <summary>测试/进程退出用。</summary>
        public static void ClearAll()
        {
            lock (Gate)
            {
                Channels.Clear();
                DocUuidToChannelId.Clear();
                _defaultChannelId = null;
            }
        }
    }
}
