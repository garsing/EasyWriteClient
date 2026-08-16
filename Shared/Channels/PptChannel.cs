using System;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace WordAddIn1
{
    public sealed class PptChannel : IOperationChannel
    {
        private PowerPoint.Presentation _presentation;
        private string _filePath;

        public PptChannel(
            string channelId,
            string docUuid,
            PowerPoint.Presentation presentation,
            string filePath)
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
            _filePath = WordChannel.NormalizePath(filePath) ?? TryReadFullName(presentation) ?? "";
        }

        public string ChannelId { get; }

        public ChannelKind Kind => ChannelKind.Ppt;

        public string DocUuid { get; }

        public string FilePath
        {
            get { return _filePath ?? ""; }
        }

        public void UpdatePresentation(PowerPoint.Presentation presentation, string filePath = null)
        {
            _presentation = presentation ?? throw new ArgumentNullException(nameof(presentation));
            string next = WordChannel.NormalizePath(filePath) ?? TryReadFullName(presentation);
            if (!string.IsNullOrEmpty(next))
            {
                _filePath = next;
            }
        }

        public bool TryGetLivePresentation(out PowerPoint.Presentation presentation)
        {
            presentation = null;
            if (_presentation == null)
            {
                return false;
            }

            try
            {
                var _ = _presentation.Name;
                presentation = _presentation;
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        public string TryGetDisplayName()
        {
            if (!TryGetLivePresentation(out PowerPoint.Presentation presentation))
            {
                return null;
            }

            try
            {
                return presentation.Name;
            }
            catch (Exception)
            {
                return null;
            }
        }

        public bool TryActivate()
        {
            if (!TryGetLivePresentation(out PowerPoint.Presentation presentation))
            {
                return false;
            }

            try
            {
                presentation.Windows[1].Activate();
                return true;
            }
            catch (Exception)
            {
                try
                {
                    presentation.Application.Activate();
                    return true;
                }
                catch (Exception)
                {
                    return false;
                }
            }
        }

        internal static string TryReadFullName(PowerPoint.Presentation presentation)
        {
            if (presentation == null)
            {
                return null;
            }

            try
            {
                string fullName = presentation.FullName;
                return string.IsNullOrWhiteSpace(fullName) ? null : fullName;
            }
            catch (Exception)
            {
                return null;
            }
        }
    }
}
