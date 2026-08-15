using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.CompilerServices;
using WordAddIn1.SpreadsheetHost;
using Excel = Microsoft.Office.Interop.Excel;

namespace WordAddIn1.OpenFiles
{
    internal sealed class EtOpenFilesDetector : IOpenFilesAppDetector
    {
        public const string TypeKey = "et";

        private readonly object _gate = new object();
        private object _app;
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
                    return _app != null && EtCom.IsAlive(_app);
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
                if (_app != null && EtCom.IsAlive(_app))
                {
                    if (!_subscribed)
                    {
                        TrySubscribeUnlocked();
                    }

                    return true;
                }

                TearDownUnlocked(raiseDetached: false);
                object app = EtCom.TryGetActiveApplication(out string progId);
                if (app == null)
                {
                    return false;
                }

                _app = app;
                TrySubscribeUnlocked();
                EasyWriteDiagnostics.Log(DebugCategory.OpenFiles,
                    "[EtOpenFilesDetector] attached progId=" + progId
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
                if (_app == null || !EtCom.IsAlive(_app))
                {
                    return result;
                }

                app = _app;
                _rcwToId.Clear();
            }

            var usedIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            try
            {
                foreach (object book in EtCom.EnumerateWorkbooks(app))
                {
                    try
                    {
                        var item = MapWorkbook(book, usedIds);
                        if (item == null)
                        {
                            // 跳过的簿仍会占住 RCW，必须释放
                            ComRelease.Safe(book);
                            continue;
                        }

                        EnsureChannel(book, item);
                        int key = RuntimeHelpers.GetHashCode(book);
                        lock (_gate)
                        {
                            _rcwToId[key] = item.Id;
                        }

                        result.Add(item);
                    }
                    catch (Exception ex)
                    {
                        ComRelease.Safe(book);
                        EasyWriteDiagnostics.Log(DebugCategory.OpenFiles,
                            "[EtOpenFilesDetector] snapshot item skip: " + ex.Message);
                    }
                }
            }
            catch (Exception ex)
            {
                EasyWriteDiagnostics.Log(DebugCategory.OpenFiles,
                    "[EtOpenFilesDetector] snapshot failed: " + ex.Message);
                MarkDetached();
                return result;
            }

            // 无打开簿时释放 Application，否则 GetActiveObject 持有会阻止用户关闭 et
            if (result.Count == 0)
            {
                EasyWriteDiagnostics.Log(DebugCategory.OpenFiles,
                    "[EtOpenFilesDetector] no workbooks — release Application RCW");
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
                var events = _app as Excel.AppEvents_Event;
                if (events == null)
                {
                    EasyWriteDiagnostics.Log(DebugCategory.OpenFiles,
                        "[EtOpenFilesDetector] no AppEvents_Event — reconcile will be used");
                    return;
                }

                events.WorkbookOpen += OnWorkbookOpen;
                events.NewWorkbook += OnNewWorkbook;
                events.WorkbookBeforeClose += OnWorkbookBeforeClose;
                _subscribed = true;
            }
            catch (Exception ex)
            {
                _subscribed = false;
                EasyWriteDiagnostics.Log(DebugCategory.OpenFiles,
                    "[EtOpenFilesDetector] subscribe failed: " + ex.Message);
            }
        }

        private void OnWorkbookOpen(Excel.Workbook book)
        {
            HandleOpened(book);
        }

        private void OnNewWorkbook(Excel.Workbook book)
        {
            HandleOpened(book);
        }

        private void OnWorkbookBeforeClose(Excel.Workbook book, ref bool cancel)
        {
            if (book == null)
            {
                return;
            }

            try
            {
                string id = null;
                int key = RuntimeHelpers.GetHashCode(book);
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
                    id = MapWorkbook(book, new HashSet<string>(StringComparer.OrdinalIgnoreCase))?.Id;
                }

                try
                {
                    string uuid = EtWorkbookIdentity.TryResolveUuid(book);
                    if (!string.IsNullOrEmpty(uuid))
                    {
                        ChannelRegistry.RemoveByDocUuid(uuid);
                    }
                }
                catch (Exception)
                {
                }

                if (!string.IsNullOrEmpty(id))
                {
                    DocumentClosed?.Invoke(id);
                }
            }
            catch (Exception ex)
            {
                EasyWriteDiagnostics.Log(DebugCategory.OpenFiles,
                    "[EtOpenFilesDetector] WorkbookBeforeClose: " + ex.Message);
            }
        }

        private void HandleOpened(object book)
        {
            if (book == null)
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

                var item = MapWorkbook(book, used);
                if (item == null)
                {
                    return;
                }

                EnsureChannel(book, item);
                int key = RuntimeHelpers.GetHashCode(book);
                lock (_gate)
                {
                    _rcwToId[key] = item.Id;
                }

                DocumentOpened?.Invoke(item);
            }
            catch (Exception ex)
            {
                EasyWriteDiagnostics.Log(DebugCategory.OpenFiles,
                    "[EtOpenFilesDetector] open event: " + ex.Message);
            }
        }

        private static void EnsureChannel(object book, OpenFileItem item)
        {
            if (book == null || item == null)
            {
                return;
            }

            try
            {
                EtChannel channel = ChannelRegistry.CreateOrGetEt(
                    book,
                    item.FullPath,
                    claimDefaultIfEmpty: true);
                item.ChannelId = channel?.ChannelId;
            }
            catch (Exception ex)
            {
                EasyWriteDiagnostics.Log(DebugCategory.OpenFiles,
                    "[EtOpenFilesDetector] EnsureChannel: " + ex.Message);
                item.ChannelId = null;
            }
        }

        private static OpenFileItem MapWorkbook(object book, HashSet<string> usedIds)
        {
            if (book == null)
            {
                return null;
            }

            string name = EtCom.TryReadName(book);
            string fullName = EtCom.TryReadFullName(book);
            bool isAddin = false;
            try
            {
                object addin = EtCom.GetProperty(book, "IsAddin");
                if (addin != null)
                {
                    isAddin = Convert.ToBoolean(addin);
                }
            }
            catch (Exception)
            {
            }

            if (SpreadsheetContentUtil.ShouldSkipWorkbook(name, fullName, isAddin))
            {
                return null;
            }

            bool saved = SpreadsheetContentUtil.IsSavedPath(fullName);
            string display;
            string id;
            string fullPath = null;
            if (saved)
            {
                fullPath = SpreadsheetContentUtil.NormalizePath(fullName);
                display = Path.GetFileName(fullPath);
                if (string.IsNullOrEmpty(display))
                {
                    display = fullPath;
                }

                id = "et:" + fullPath;
            }
            else
            {
                if (string.IsNullOrWhiteSpace(name))
                {
                    name = "未命名工作簿";
                }

                display = name;
                id = "et:unsaved:" + name;
                int n = 2;
                while (usedIds != null && usedIds.Contains(id))
                {
                    id = "et:unsaved:" + name + "#" + n;
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
                    var events = _app as Excel.AppEvents_Event;
                    if (events != null)
                    {
                        events.WorkbookOpen -= OnWorkbookOpen;
                        events.NewWorkbook -= OnNewWorkbook;
                        events.WorkbookBeforeClose -= OnWorkbookBeforeClose;
                    }
                }
                catch (Exception)
                {
                }
            }

            _subscribed = false;
            object app = _app;
            _app = null;
            _rcwToId.Clear();

            ComRelease.Safe(app);
            if (app != null)
            {
                ComRelease.CollectPending();
            }

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
