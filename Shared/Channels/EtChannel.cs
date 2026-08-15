using System;
using WordAddIn1.OpenFiles;

namespace WordAddIn1
{
    public sealed class EtChannel : IOperationChannel
    {
        private object _workbook;
        private string _filePath;

        public EtChannel(string channelId, string docUuid, object workbook, string filePath)
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
            _workbook = workbook ?? throw new ArgumentNullException(nameof(workbook));
            _filePath = WordChannel.NormalizePath(filePath) ?? EtCom.TryReadFullName(workbook) ?? "";
        }

        public string ChannelId { get; }

        public ChannelKind Kind => ChannelKind.Et;

        public string DocUuid { get; }

        public string FilePath
        {
            get { return _filePath ?? ""; }
        }

        public void UpdateWorkbook(object workbook, string filePath = null)
        {
            _workbook = workbook ?? throw new ArgumentNullException(nameof(workbook));
            string next = WordChannel.NormalizePath(filePath) ?? EtCom.TryReadFullName(workbook);
            if (!string.IsNullOrEmpty(next))
            {
                _filePath = next;
            }
        }

        public bool TryGetLiveWorkbook(out object workbook)
        {
            workbook = null;
            if (_workbook == null)
            {
                return false;
            }

            try
            {
                var _ = EtCom.GetProperty(_workbook, "Name");
                workbook = _workbook;
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        public string TryGetDisplayName()
        {
            if (!TryGetLiveWorkbook(out object book))
            {
                return null;
            }

            return EtCom.TryReadName(book);
        }
    }
}
