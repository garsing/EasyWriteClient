using System;
using WordAddIn1.OpenFiles;

namespace WordAddIn1
{
    public sealed class WppChannel : IOperationChannel
    {
        private readonly object _gate = new object();
        private object _presentation;
        private string _filePath;

        public WppChannel(string channelId, string docUuid, object presentation, string filePath)
        {
            if (string.IsNullOrWhiteSpace(channelId))
            {
                throw new ArgumentException("channel_id 不能为空。", nameof(channelId));
            }

            if (string.IsNullOrWhiteSpace(docUuid))
            {
                throw new ArgumentException("doc_uuid 不能为空。", nameof(docUuid));
            }

            ChannelId = channelId.Trim();
            DocUuid = docUuid.Trim();
            _presentation = presentation ?? throw new ArgumentNullException(nameof(presentation));
            _filePath = WordChannel.NormalizePath(filePath) ?? WppCom.TryReadFullName(presentation) ?? "";
        }

        public string ChannelId { get; }

        public ChannelKind Kind => ChannelKind.Wpp;

        public string DocUuid { get; }

        public string FilePath
        {
            get { return _filePath ?? ""; }
        }

        public void UpdatePresentation(object presentation, string filePath = null)
        {
            if (presentation == null)
            {
                throw new ArgumentNullException(nameof(presentation));
            }

            object old = null;
            lock (_gate)
            {
                if (!ReferenceEquals(_presentation, presentation))
                {
                    old = _presentation;
                    _presentation = presentation;
                }

                string next = WordChannel.NormalizePath(filePath) ?? WppCom.TryReadFullName(presentation);
                if (!string.IsNullOrEmpty(next))
                {
                    _filePath = next;
                }
            }

            // 探测器换 RCW 时只减一次引用；FinalRelease 容易拆掉仍被打开路径持有的同 COM
            ComRelease.ReleaseOnce(old);
        }

        public void ReleaseCom()
        {
            object presentation;
            lock (_gate)
            {
                presentation = _presentation;
                _presentation = null;
            }

            if (presentation != null)
            {
                try
                {
                    WppPresentationIdentity.Unregister(presentation);
                }
                catch (Exception)
                {
                }

                ComRelease.Safe(presentation);
            }
        }

        public bool TryGetLivePresentation(out object presentation)
        {
            presentation = null;
            object current;
            lock (_gate)
            {
                current = _presentation;
            }

            if (current == null)
            {
                return false;
            }

            try
            {
                var _ = WppCom.GetProperty(current, "Name");
                presentation = current;
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        public string TryGetDisplayName()
        {
            if (!TryGetLivePresentation(out object presentation))
            {
                return null;
            }

            return WppCom.TryReadName(presentation);
        }

        public bool TryActivate()
        {
            if (!TryGetLivePresentation(out object presentation))
            {
                return false;
            }

            try
            {
                object windows = WppCom.GetProperty(presentation, "Windows");
                object window = WppCom.GetIndexed(windows, 1);
                WppCom.Invoke(window, "Activate");
                return true;
            }
            catch (Exception)
            {
                try
                {
                    object app = WppCom.GetProperty(presentation, "Application");
                    WppCom.Invoke(app, "Activate");
                    return true;
                }
                catch (Exception)
                {
                    return false;
                }
            }
        }
    }
}
