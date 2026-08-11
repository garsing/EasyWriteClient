using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1.OpenFiles
{
    /// <summary>
    /// WPS 文字「打开文件」探测器：晚绑定附着、快照 Documents；可订则订事件，否则由 Monitor 对账。
    /// 禁止 CreateOrGetWord / 建渠道。
    /// </summary>
    internal sealed class WpsOpenFilesDetector : IOpenFilesAppDetector
    {
        public const string TypeKey = "wps";

        private readonly object _gate = new object();
        private object _app;
        private string _progId;
        private bool _subscribed;
        private bool _disposed;
        private readonly Dictionary<int, string> _rcwToId = new Dictionary<int, string>();

        public string AppType => TypeKey;

        /// <summary>事件 sink 是否挂成功；失败时 Monitor 启用低频对账。</summary>
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
                    return _app != null && WpsCom.IsAlive(_app);
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
                if (_app != null && WpsCom.IsAlive(_app))
                {
                    if (!_subscribed)
                    {
                        TrySubscribeUnlocked();
                    }

                    return true;
                }

                TearDownUnlocked(raiseDetached: false);

                string progId;
                object app = WpsCom.TryGetActiveApplication(out progId);
                if (app == null)
                {
                    return false;
                }

                _app = app;
                _progId = progId;
                TrySubscribeUnlocked();
                EasyWriteDiagnostics.Log(DebugCategory.OpenFiles,
                    "[WpsOpenFilesDetector] attached progId=" + progId
                    + " events=" + _subscribed);
                return true;
            }
        }

        public IReadOnlyList<OpenFileItem> Snapshot()
        {
            var result = new List<OpenFileItem>();
            object app;
            lock (_gate)
            {
                if (_app == null || !WpsCom.IsAlive(_app))
                {
                    return result;
                }

                app = _app;
                _rcwToId.Clear();
            }

            var usedIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            try
            {
                foreach (object doc in WpsCom.EnumerateDocuments(app))
                {
                    try
                    {
                        var item = MapDocument(doc, usedIds);
                        if (item == null)
                        {
                            continue;
                        }

                        // 禁止建渠道
                        item.ChannelId = null;

                        int key = RuntimeHelpers.GetHashCode(doc);
                        lock (_gate)
                        {
                            _rcwToId[key] = item.Id;
                        }

                        result.Add(item);
                    }
                    catch (Exception ex)
                    {
                        EasyWriteDiagnostics.Log(DebugCategory.OpenFiles,
                            "[WpsOpenFilesDetector] snapshot item skip: " + ex.Message);
                    }
                }
            }
            catch (Exception ex)
            {
                EasyWriteDiagnostics.Log(DebugCategory.OpenFiles,
                    "[WpsOpenFilesDetector] snapshot failed: " + ex.Message);
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

            // 部分 WPS 暴露与 Word 兼容的 ApplicationEvents4；仅用于事件，绝不 CreateOrGetWord
            try
            {
                var events = _app as Word.ApplicationEvents4_Event;
                if (events == null)
                {
                    EasyWriteDiagnostics.Log(DebugCategory.OpenFiles,
                        "[WpsOpenFilesDetector] no ApplicationEvents4_Event — reconcile will be used");
                    return;
                }

                events.DocumentOpen += OnDocumentOpen;
                events.NewDocument += OnNewDocument;
                events.DocumentBeforeClose += OnDocumentBeforeClose;
                _subscribed = true;
            }
            catch (Exception ex)
            {
                _subscribed = false;
                EasyWriteDiagnostics.Log(DebugCategory.OpenFiles,
                    "[WpsOpenFilesDetector] subscribe failed: " + ex.Message);
            }
        }

        private void OnDocumentOpen(Word.Document doc)
        {
            HandleOpened(doc);
        }

        private void OnNewDocument(Word.Document doc)
        {
            HandleOpened(doc);
        }

        private void OnDocumentBeforeClose(Word.Document doc, ref bool cancel)
        {
            if (doc == null)
            {
                return;
            }

            try
            {
                string id = null;
                int key = RuntimeHelpers.GetHashCode(doc);
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
                    var used = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                    id = MapDocument(doc, used)?.Id;
                }

                if (!string.IsNullOrEmpty(id))
                {
                    DocumentClosed?.Invoke(id);
                }
            }
            catch (Exception ex)
            {
                EasyWriteDiagnostics.Log(DebugCategory.OpenFiles,
                    "[WpsOpenFilesDetector] DocumentBeforeClose: " + ex.Message);
            }
        }

        private void HandleOpened(object doc)
        {
            if (doc == null)
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

                var item = MapDocument(doc, used);
                if (item == null)
                {
                    return;
                }

                item.ChannelId = null;

                int key = RuntimeHelpers.GetHashCode(doc);
                lock (_gate)
                {
                    _rcwToId[key] = item.Id;
                }

                DocumentOpened?.Invoke(item);
            }
            catch (Exception ex)
            {
                EasyWriteDiagnostics.Log(DebugCategory.OpenFiles,
                    "[WpsOpenFilesDetector] open event: " + ex.Message);
            }
        }

        private static OpenFileItem MapDocument(object doc, HashSet<string> usedIds)
        {
            if (doc == null)
            {
                return null;
            }

            string fullName = WpsCom.TryReadFullName(doc);
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

                id = "wps:" + fullPath;
            }
            else
            {
                string name = WpsCom.TryReadName(doc);
                if (string.IsNullOrWhiteSpace(name))
                {
                    name = "未命名文档";
                }

                display = name;
                id = "wps:unsaved:" + name;
                int n = 2;
                while (usedIds != null && usedIds.Contains(id))
                {
                    id = "wps:unsaved:" + name + "#" + n;
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

            return fullName.IndexOf(Path.DirectorySeparatorChar) >= 0
                || fullName.IndexOf(Path.AltDirectorySeparatorChar) >= 0;
        }

        private static string NormalizePath(string path)
        {
            try
            {
                return Path.GetFullPath(path).TrimEnd('\\', '/');
            }
            catch (Exception)
            {
                return path?.Trim();
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
                    var events = _app as Word.ApplicationEvents4_Event;
                    if (events != null)
                    {
                        events.DocumentOpen -= OnDocumentOpen;
                        events.NewDocument -= OnNewDocument;
                        events.DocumentBeforeClose -= OnDocumentBeforeClose;
                    }
                }
                catch (Exception)
                {
                }
            }

            _subscribed = false;
            _app = null;
            _progId = null;
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
