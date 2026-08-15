using System;
using Excel = Microsoft.Office.Interop.Excel;

namespace WordAddIn1
{
    public sealed class ExcelChannel : IOperationChannel
    {
        private Excel.Workbook _workbook;
        private string _filePath;

        public ExcelChannel(string channelId, string docUuid, Excel.Workbook workbook, string filePath)
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
            _filePath = WordChannel.NormalizePath(filePath) ?? TryReadFullName(workbook) ?? "";
        }

        public string ChannelId { get; }

        public ChannelKind Kind => ChannelKind.Excel;

        public string DocUuid { get; }

        public string FilePath
        {
            get { return _filePath ?? ""; }
        }

        public void UpdateWorkbook(Excel.Workbook workbook, string filePath = null)
        {
            _workbook = workbook ?? throw new ArgumentNullException(nameof(workbook));
            string next = WordChannel.NormalizePath(filePath) ?? TryReadFullName(workbook);
            if (!string.IsNullOrEmpty(next))
            {
                _filePath = next;
            }
        }

        public bool TryGetLiveWorkbook(out Excel.Workbook workbook)
        {
            workbook = null;
            if (_workbook == null)
            {
                return false;
            }

            try
            {
                var _ = _workbook.Name;
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
            if (!TryGetLiveWorkbook(out Excel.Workbook wb))
            {
                return null;
            }

            try
            {
                return wb.Name;
            }
            catch (Exception)
            {
                return null;
            }
        }

        internal static string TryReadFullName(Excel.Workbook workbook)
        {
            if (workbook == null)
            {
                return null;
            }

            try
            {
                string fullName = workbook.FullName;
                return string.IsNullOrWhiteSpace(fullName) ? null : fullName;
            }
            catch (Exception)
            {
                return null;
            }
        }
    }
}
