using System;
using WordAddIn1.OpenFiles;

namespace WordAddIn1
{
    public sealed class EtChannel : IOperationChannel
    {
        private readonly object _gate = new object();
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
            if (workbook == null)
            {
                throw new ArgumentNullException(nameof(workbook));
            }

            object old = null;
            lock (_gate)
            {
                if (!ReferenceEquals(_workbook, workbook))
                {
                    old = _workbook;
                    _workbook = workbook;
                }

                string next = WordChannel.NormalizePath(filePath) ?? EtCom.TryReadFullName(workbook);
                if (!string.IsNullOrEmpty(next))
                {
                    _filePath = next;
                }
            }

            ComRelease.Safe(old);
        }

        /// <summary>
        /// 渠道移除时释放 Workbook RCW，避免拖住 et 进程。
        /// </summary>
        public void ReleaseCom()
        {
            object wb;
            lock (_gate)
            {
                wb = _workbook;
                _workbook = null;
            }

            if (wb != null)
            {
                try
                {
                    EtWorkbookIdentity.Unregister(wb);
                }
                catch (Exception)
                {
                }

                ComRelease.Safe(wb);
            }
        }

        public bool TryGetLiveWorkbook(out object workbook)
        {
            workbook = null;
            object wb;
            lock (_gate)
            {
                wb = _workbook;
            }

            if (wb == null)
            {
                return false;
            }

            try
            {
                var _ = EtCom.GetProperty(wb, "Name");
                workbook = wb;
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
