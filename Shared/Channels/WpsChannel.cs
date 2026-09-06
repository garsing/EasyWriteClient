using System;
using WordAddIn1.OpenFiles;

namespace WordAddIn1
{
    /// <summary>
    /// WPS 文字操作渠道：绑定晚绑定 Document RCW；channel_id 形如 <c>w1</c>。
    /// </summary>
    public sealed class WpsChannel : IOperationChannel
    {
        private object _document;
        private string _filePath;

        public WpsChannel(string channelId, string docUuid, object document, string filePath)
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
            _document = document ?? throw new ArgumentNullException(nameof(document));
            _filePath = WordChannel.NormalizePath(filePath) ?? WpsCom.TryReadFullName(document) ?? "";
        }

        public string ChannelId { get; }

        public ChannelKind Kind => ChannelKind.Wps;

        public string DocUuid { get; }

        public string FilePath
        {
            get { return _filePath ?? ""; }
        }

        public object Document
        {
            get { return _document; }
        }

        public void UpdateDocument(object document, string filePath = null)
        {
            if (document == null)
            {
                throw new ArgumentNullException(nameof(document));
            }

            _document = document;
            string next = WordChannel.NormalizePath(filePath) ?? WpsCom.TryReadFullName(document);
            if (!string.IsNullOrEmpty(next))
            {
                _filePath = next;
            }
        }

        public bool TryGetLiveDocument(out object document)
        {
            document = null;
            if (_document == null)
            {
                return false;
            }

            try
            {
                // 触碰 COM 属性以检测已关闭文档
                var _ = WpsCom.GetProperty(_document, "Name");
                document = _document;
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        public string TryGetDisplayName()
        {
            if (!TryGetLiveDocument(out object doc))
            {
                return null;
            }

            return WpsCom.TryReadName(doc);
        }
    }
}
