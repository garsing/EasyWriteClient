using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.CompilerServices;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace WordAddIn1.OpenFiles
{
    /// <summary>
    /// PowerPoint「打开文件」探测器：只附着已运行实例；建渠不抢默认。
    /// </summary>
    internal sealed class PptOpenFilesDetector : IOpenFilesAppDetector
    {
        public const string TypeKey = "ppt";

        private readonly object _gate = new object();
        private PowerPoint.Application _app;
        private bool _subscribed;
        private bool _disposed;
        private readonly Dictionary<int, string> _rcwToId = new Dictionary<int, string>();

        public string AppType => TypeKey;

        public bool EventsSubscribed
        {
            get
            {
                lock (_gate)
                {
                    return _subscribed;
                }
            }
        }

        public bool IsAttached
        {
            get
            {
                lock (_gate)
                {
                    return _app != null && IsAppAliveUnlocked();
                }
            }
        }

        public event Action<OpenFileItem> DocumentOpened;
        public event Action<string> DocumentClosed;
        public event Action Detached;

        public bool TryAttach()
        {
            if (_disposed)
            {
                return false;
            }

            lock (_gate)
            {
                if (_app != null && IsAppAliveUnlocked())
                {
                    if (!_subscribed)
                    {
                        TrySubscribeUnlocked();
                    }

                    return true;
                }

                TearDownUnlocked(raiseDetached: false);
                if (!PowerPointApplicationResolver.TryResolveTyped(
                        out PowerPoint.Application app,
                        out _,
                        createIfMissing: false,
                        makeVisible: false))
                {
                    return false;
                }

                _app = app;
                TrySubscribeUnlocked();
                EasyWriteDiagnostics.Log(DebugCategory.OpenFiles,
                    "[PptOpenFilesDetector] attached events=" + _subscribed);
                return true;
            }
        }

        public IReadOnlyList<OpenFileItem> Snapshot()
        {
            var result = new List<OpenFileItem>();
            PowerPoint.Application app;
            lock (_gate)
            {
                if (_app == null || !IsAppAliveUnlocked())
                {
                    return result;
                }

                app = _app;
                _rcwToId.Clear();
            }

            var usedIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            try
            {
                foreach (PowerPoint.Presentation presentation in app.Presentations)
                {
                    try
                    {
                        var item = MapPresentation(presentation, usedIds);
                        if (item == null)
                        {
                            continue;
                        }

                        EnsureChannel(presentation, item);
                        int key = RuntimeHelpers.GetHashCode(presentation);
                        lock (_gate)
                        {
                            _rcwToId[key] = item.Id;
                        }

                        result.Add(item);
                    }
                    catch (Exception ex)
                    {
                        EasyWriteDiagnostics.Log(DebugCategory.OpenFiles,
                            "[PptOpenFilesDetector] snapshot item skip: " + ex.Message);
                    }
                }
            }
            catch (Exception ex)
            {
                EasyWriteDiagnostics.Log(DebugCategory.OpenFiles,
                    "[PptOpenFilesDetector] snapshot failed: " + ex.Message);
                MarkDetached();
                return result;
            }

            if (result.Count == 0)
            {
                EasyWriteDiagnostics.Log(DebugCategory.OpenFiles,
                    "[PptOpenFilesDetector] no presentations — detach");
                MarkDetached();
            }

            return result;
        }

        public void Dispose()
        {
            if (_disposed)
            {
                return;
            }

            _disposed = true;
            lock (_gate)
            {
                TearDownUnlocked(raiseDetached: false);
            }
        }

        private void TrySubscribeUnlocked()
        {
            _subscribed = false;
            if (_app == null)
            {
                return;
            }

            try
            {
                var events = _app as PowerPoint.EApplication_Event;
                if (events == null)
                {
                    EasyWriteDiagnostics.Log(DebugCategory.OpenFiles,
                        "[PptOpenFilesDetector] no EApplication_Event — reconcile will be used");
                    return;
                }

                events.PresentationOpen += OnPresentationOpen;
                events.NewPresentation += OnNewPresentation;
                events.PresentationClose += OnPresentationClose;
                _subscribed = true;
            }
            catch (Exception ex)
            {
                _subscribed = false;
                EasyWriteDiagnostics.Log(DebugCategory.OpenFiles,
                    "[PptOpenFilesDetector] subscribe failed: " + ex.Message);
            }
        }

        private void OnPresentationOpen(PowerPoint.Presentation presentation)
        {
            HandleOpened(presentation);
        }

        private void OnNewPresentation(PowerPoint.Presentation presentation)
        {
            HandleOpened(presentation);
        }

        private void OnPresentationClose(PowerPoint.Presentation presentation)
        {
            if (presentation == null)
            {
                return;
            }

            try
            {
                string id = null;
                int key = RuntimeHelpers.GetHashCode(presentation);
                lock (_gate)
                {
                    if (_rcwToId.TryGetValue(key, out var mapped))
                    {
                        id = mapped;
                        _rcwToId.Remove(key);
                    }
                }

                if (string.IsNullOrEmpty(id))
                {
                    id = MapPresentation(presentation, new HashSet<string>(StringComparer.OrdinalIgnoreCase))?.Id;
                }

                if (!string.IsNullOrEmpty(id))
                {
                    DocumentClosed?.Invoke(id);
                }
            }
            catch (Exception ex)
            {
                EasyWriteDiagnostics.Log(DebugCategory.OpenFiles,
                    "[PptOpenFilesDetector] PresentationClose: " + ex.Message);
            }
        }

        private void HandleOpened(PowerPoint.Presentation presentation)
        {
            if (presentation == null)
            {
                return;
            }

            try
            {
                var used = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                lock (_gate)
                {
                    foreach (var existing in _rcwToId.Values)
                    {
                        used.Add(existing);
                    }
                }

                var item = MapPresentation(presentation, used);
                if (item == null)
                {
                    return;
                }

                EnsureChannel(presentation, item);
                int key = RuntimeHelpers.GetHashCode(presentation);
                lock (_gate)
                {
                    _rcwToId[key] = item.Id;
                }

                DocumentOpened?.Invoke(item);
            }
            catch (Exception ex)
            {
                EasyWriteDiagnostics.Log(DebugCategory.OpenFiles,
                    "[PptOpenFilesDetector] open event: " + ex.Message);
            }
        }

        private static void EnsureChannel(PowerPoint.Presentation presentation, OpenFileItem item)
        {
            if (presentation == null || item == null)
            {
                return;
            }

            try
            {
                PptChannel channel = ChannelRegistry.CreateOrGetPpt(
                    presentation,
                    item.FullPath,
                    claimDefaultIfEmpty: false);
                item.ChannelId = channel?.ChannelId;
            }
            catch (Exception ex)
            {
                EasyWriteDiagnostics.Log(DebugCategory.OpenFiles,
                    "[PptOpenFilesDetector] EnsureChannel: " + ex.Message);
                item.ChannelId = null;
            }
        }

        private static OpenFileItem MapPresentation(
            PowerPoint.Presentation presentation,
            HashSet<string> usedIds)
        {
            if (presentation == null)
            {
                return null;
            }

            string name = null;
            string fullName = null;
            try
            {
                name = presentation.Name;
                fullName = presentation.FullName;
            }
            catch (Exception)
            {
                return null;
            }

            bool saved = IsSavedPath(fullName);
            string display;
            string id;
            string fullPath = null;
            if (saved)
            {
                fullPath = NormalizePath(fullName);
                display = Path.GetFileName(fullPath);
                if (string.IsNullOrEmpty(display))
                {
                    display = fullPath;
                }

                id = "ppt:" + fullPath;
            }
            else
            {
                if (string.IsNullOrWhiteSpace(name))
                {
                    name = "演示文稿1";
                }

                display = name;
                id = "ppt:unsaved:" + name;
                int n = 2;
                while (usedIds != null && usedIds.Contains(id))
                {
                    id = "ppt:unsaved:" + name + "#" + n;
                    n++;
                }
            }

            usedIds?.Add(id);
            return new OpenFileItem
            {
                Id = id,
                AppType = TypeKey,
                DisplayName = display,
                FullPath = fullPath,
                IsSaved = saved,
                ChannelId = null
            };
        }

        private static bool IsSavedPath(string fullName)
        {
            if (string.IsNullOrWhiteSpace(fullName))
            {
                return false;
            }

            try
            {
                return Path.IsPathRooted(fullName) && !string.IsNullOrEmpty(Path.GetDirectoryName(fullName));
            }
            catch (Exception)
            {
                return false;
            }
        }

        private static string NormalizePath(string path)
        {
            try
            {
                return Path.GetFullPath(path);
            }
            catch (Exception)
            {
                return path;
            }
        }

        private bool IsAppAliveUnlocked()
        {
            try
            {
                var _ = _app?.Name;
                return _app != null;
            }
            catch (Exception)
            {
                return false;
            }
        }

        private void MarkDetached()
        {
            lock (_gate)
            {
                TearDownUnlocked(raiseDetached: true);
            }
        }

        private void TearDownUnlocked(bool raiseDetached)
        {
            if (_app != null && _subscribed)
            {
                try
                {
                    var events = _app as PowerPoint.EApplication_Event;
                    if (events != null)
                    {
                        events.PresentationOpen -= OnPresentationOpen;
                        events.NewPresentation -= OnNewPresentation;
                        events.PresentationClose -= OnPresentationClose;
                    }
                }
                catch (Exception)
                {
                }
            }

            _subscribed = false;
            _app = null;
            _rcwToId.Clear();

            if (raiseDetached)
            {
                try
                {
                    Detached?.Invoke();
                }
                catch (Exception)
                {
                }
            }
        }
    }
}
