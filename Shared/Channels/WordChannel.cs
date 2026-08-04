using System;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    /// <summary>
    /// Word 文档操作渠道：绑定一份 <see cref="Word.Document"/>（D11）。
    /// </summary>
    public sealed class WordChannel : IOperationChannel
    {
        private Word.Document _document;
        private string _filePath;

        public WordChannel(string channelId, string docUuid, Word.Document document, string filePath)
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
            _filePath = NormalizePath(filePath) ?? TryReadFullName(document) ?? "";
        }

        public string ChannelId { get; }

        public ChannelKind Kind => ChannelKind.Word;

        public string DocUuid { get; }

        public string FilePath
        {
            get { return _filePath ?? ""; }
        }

        /// <summary>原始 Document RCW；关闭后可能失效，优先用 <see cref="TryGetLiveDocument"/>。</summary>
        public Word.Document Document
        {
            get { return _document; }
        }

        public void UpdateDocument(Word.Document document, string filePath = null)
        {
            if (document == null)
            {
                throw new ArgumentNullException(nameof(document));
            }

            _document = document;
            string next = NormalizePath(filePath) ?? TryReadFullName(document);
            if (!string.IsNullOrEmpty(next))
            {
                _filePath = next;
            }
        }

        public bool TryGetLiveDocument(out Word.Document document)
        {
            document = null;
            if (_document == null)
            {
                return false;
            }

            try
            {
                // 触碰 COM 属性以检测已关闭文档
                var _ = _document.Name;
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
            if (!TryGetLiveDocument(out Word.Document doc))
            {
                return null;
            }

            try
            {
                return doc.Name;
            }
            catch (Exception)
            {
                return null;
            }
        }

        internal static string NormalizePath(string path)
        {
            if (string.IsNullOrWhiteSpace(path))
            {
                return null;
            }

            return path.Trim();
        }

        internal static string TryReadFullName(Word.Document document)
        {
            if (document == null)
            {
                return null;
            }

            try
            {
                string fullName = document.FullName;
                if (string.IsNullOrWhiteSpace(fullName))
                {
                    return null;
                }

                if (fullName.StartsWith("Unsaved", StringComparison.OrdinalIgnoreCase))
                {
                    return null;
                }

                return fullName;
            }
            catch (Exception)
            {
                return null;
            }
        }
    }
}
