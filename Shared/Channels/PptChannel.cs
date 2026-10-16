using System;
using WordAddIn1.OpenFiles;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace WordAddIn1
{
    public sealed class PptChannel : IOperationChannel
    {
        private readonly object _gate = new object();
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
            if (presentation == null)
            {
                throw new ArgumentNullException(nameof(presentation));
            }

            PowerPoint.Presentation old = null;
            lock (_gate)
            {
                if (!ReferenceEquals(_presentation, presentation))
                {
                    old = _presentation;
                    _presentation = presentation;
                }

                string next = WordChannel.NormalizePath(filePath) ?? TryReadFullName(presentation);
                if (!string.IsNullOrEmpty(next))
                {
                    _filePath = next;
                }
            }

            // 探测器换 RCW 时只减一次引用；FinalRelease 容易拆掉仍被打开路径持有的同 COM
            ComRelease.ReleaseOnce(old);
        }

        /// <summary>渠道摘掉时释放 Presentation RCW（对齐 WppChannel），避免钉住无窗空壳。</summary>
        public void ReleaseCom()
        {
            PowerPoint.Presentation presentation;
            lock (_gate)
            {
                presentation = _presentation;
                _presentation = null;
            }

            if (presentation != null)
            {
                ComRelease.Safe(presentation);
            }
        }

        public bool TryGetLivePresentation(out PowerPoint.Presentation presentation)
        {
            presentation = null;
            PowerPoint.Presentation current;
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
                var _ = current.Name;
                presentation = current;
                return true;
            }
            catch (Exception)
            {
                if (!TryRebindFromRunningApp(out PowerPoint.Presentation fresh))
                {
                    return false;
                }

                UpdatePresentation(fresh);
                presentation = fresh;
                return true;
            }
        }

        /// <summary>
        /// 手中的 Presentation 已断、PowerPoint 进程还在时，按文件路径从 Presentations 再取一次。不新开 PowerPoint。
        /// </summary>
        private bool TryRebindFromRunningApp(out PowerPoint.Presentation presentation)
        {
            presentation = null;
            string want = WordChannel.NormalizePath(_filePath);
            if (string.IsNullOrEmpty(want))
            {
                return false;
            }

            if (!PowerPointApplicationResolver.TryResolveTyped(
                    out PowerPoint.Application app,
                    out _,
                    createIfMissing: false,
                    makeVisible: false)
                || app == null)
            {
                return false;
            }

            try
            {
                foreach (PowerPoint.Presentation one in app.Presentations)
                {
                    string full = null;
                    try
                    {
                        full = WordChannel.NormalizePath(one.FullName);
                    }
                    catch (Exception)
                    {
                    }

                    if (!string.IsNullOrEmpty(full)
                        && string.Equals(full, want, StringComparison.OrdinalIgnoreCase))
                    {
                        presentation = one;
                        return true;
                    }
                }
            }
            catch (Exception)
            {
            }

            return false;
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
                PowerPointApplicationResolver.EnsureVisible(presentation.Application);
                PowerPointApplicationResolver.EnsurePresentationHasWindow(presentation);
                if (presentation.Windows.Count >= 1)
                {
                    presentation.Windows[1].Activate();
                    return true;
                }
            }
            catch (Exception)
            {
            }

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
