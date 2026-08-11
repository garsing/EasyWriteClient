using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1.OpenFiles
{
    /// <summary>
    /// Word「打开文件」探测器：只附着已运行 Word，快照 + DocumentOpen/NewDocument/DocumentBeforeClose。
    /// </summary>
    internal sealed class WordOpenFilesDetector : IOpenFilesAppDetector
    {
        private readonly Func<object> _resolveWordApp;
        private readonly object _gate = new object();
        private Word.Application _app;
        private bool _subscribed;
        private bool _disposed;
        private readonly Dictionary<int, string> _rcwToId = new Dictionary<int, string>();

        public WordOpenFilesDetector(Func<object> resolveWordApp)
        {
            _resolveWordApp = resolveWordApp ?? throw new ArgumentNullException(nameof(resolveWordApp));
        }

        public string AppType => "word";

        public bool IsAttached
        {
            get
            {
                lock (_gate)
                {
                    return _app != null && _subscribed;
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
                if (_app != null && _subscribed)
                {
                    if (IsAppAliveUnlocked())
                    {
                        return true;
                    }

                    TearDownUnlocked(raiseDetached: false);
                }

                Word.Application app = null;
                try
                {
                    app = _resolveWordApp?.Invoke() as Word.Application;
                }
                catch (Exception ex)
                {
                    EasyWriteDiagnostics.Log(DebugCategory.OpenFiles,
                        "[WordOpenFilesDetector] resolve failed: " + ex.Message);
                    app = null;
                }

                if (app == null)
                {
                    return false;
                }

                try
                {
                    var _ = app.Name;
                }
                catch (Exception ex)
                {
                    EasyWriteDiagnostics.Log(DebugCategory.OpenFiles,
                        "[WordOpenFilesDetector] app not alive: " + ex.Message);
                    return false;
                }

                _app = app;
                try
                {
                    var events = (Word.ApplicationEvents4_Event)_app;
                    events.DocumentOpen += OnDocumentOpen;
                    events.NewDocument += OnNewDocument;
                    events.DocumentBeforeClose += OnDocumentBeforeClose;
                    _subscribed = true;
                    EasyWriteDiagnostics.Log(DebugCategory.OpenFiles,
                        "[WordOpenFilesDetector] attached and subscribed");
                    return true;
                }
                catch (Exception ex)
                {
                    EasyWriteDiagnostics.Log(DebugCategory.OpenFiles,
                        "[WordOpenFilesDetector] subscribe failed: " + ex.Message);
                    TearDownUnlocked(raiseDetached: false);
                    return false;
                }
            }
        }

        public IReadOnlyList<OpenFileItem> Snapshot()
        {
            var result = new List<OpenFileItem>();
            Word.Application app;
            lock (_gate)
            {
                if (!_subscribed || _app == null || !IsAppAliveUnlocked())
                {
                    return result;
                }

                app = _app;
                _rcwToId.Clear();
            }

            var usedIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            try
            {
                foreach (Word.Document doc in app.Documents)
                {
                    try
                    {
                        var item = MapDocument(doc, usedIds);
                        if (item == null)
                        {
                            continue;
                        }

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
                            "[WordOpenFilesDetector] snapshot item skip: " + ex.Message);
                    }
                }
            }
            catch (Exception ex)
            {
                EasyWriteDiagnostics.Log(DebugCategory.OpenFiles,
                    "[WordOpenFilesDetector] snapshot failed: " + ex.Message);
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
                    var item = MapDocument(doc, used);
                    id = item?.Id;
                }

                if (!string.IsNullOrEmpty(id))
                {
                    DocumentClosed?.Invoke(id);
                }
            }
            catch (Exception ex)
            {
                EasyWriteDiagnostics.Log(DebugCategory.OpenFiles,
                    "[WordOpenFilesDetector] DocumentBeforeClose: " + ex.Message);
            }
        }

        private void HandleOpened(Word.Document doc)
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
                    "[WordOpenFilesDetector] open event: " + ex.Message);
                if (IsComDead(ex))
                {
                    MarkDetached();
                }
            }
        }

        private static OpenFileItem MapDocument(Word.Document doc, HashSet<string> usedIds)
        {
            if (doc == null)
            {
                return null;
            }

            string fullName = WordChannel.TryReadFullName(doc);
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

                id = fullPath;
            }
            else
            {
                string name = null;
                try
                {
                    name = doc.Name;
                }
                catch (Exception)
                {
                    name = null;
                }

                if (string.IsNullOrWhiteSpace(name))
                {
                    name = "未命名文档";
                }

                display = name;
                id = "word:unsaved:" + name;
                int n = 2;
                while (usedIds != null && usedIds.Contains(id))
                {
                    id = "word:unsaved:" + name + "#" + n;
                    n++;
                }
            }

            usedIds?.Add(id);

            return new OpenFileItem
            {
                Id = id,
                AppType = "word",
                DisplayName = display,
                FullPath = fullPath,
                IsSaved = saved
            };
        }

        private static bool IsSavedPath(string fullName)
        {
            if (string.IsNullOrWhiteSpace(fullName))
            {
                return false;
            }

            // 未保存文档 FullName 常仅为「文档1」无目录分隔符
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

        private bool IsAppAliveUnlocked()
        {
            if (_app == null)
            {
                return false;
            }

            try
            {
                var _ = _app.Name;
                return true;
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
                    var events = (Word.ApplicationEvents4_Event)_app;
                    events.DocumentOpen -= OnDocumentOpen;
                    events.NewDocument -= OnNewDocument;
                    events.DocumentBeforeClose -= OnDocumentBeforeClose;
                }
                catch (Exception)
                {
                    // Word 可能已退出
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

        private static bool IsComDead(Exception ex)
        {
            if (ex is COMException)
            {
                return true;
            }

            if (ex is InvalidComObjectException)
            {
                return true;
            }

            string msg = ex.Message ?? "";
            return msg.IndexOf("RPC", StringComparison.OrdinalIgnoreCase) >= 0
                || msg.IndexOf("COM", StringComparison.OrdinalIgnoreCase) >= 0;
        }
    }
}
