using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using WordAddIn1.SpreadsheetHost;
using Excel = Microsoft.Office.Interop.Excel;

namespace WordAddIn1.OpenFiles
{
    /// <summary>
    /// Excel「打开文件」探测器：对齐 Word——宿主注入 Application + 快照 + WorkbookOpen/New/BeforeClose。
    /// </summary>
    internal sealed class ExcelOpenFilesDetector : IOpenFilesAppDetector
    {
        public const string TypeKey = "excel";

        private readonly Func<object> _resolveExcelApp;
        private readonly object _gate = new object();
        private Excel.Application _app;
        private bool _subscribed;
        private bool _disposed;
        private readonly Dictionary<int, string> _rcwToId = new Dictionary<int, string>();

        public ExcelOpenFilesDetector(Func<object> resolveExcelApp)
        {
            _resolveExcelApp = resolveExcelApp ?? throw new ArgumentNullException(nameof(resolveExcelApp));
        }

        public string AppType => TypeKey;

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
                        int existingCount = -1;
                        try
                        {
                            existingCount = _app.Workbooks.Count;
                        }
                        catch (Exception)
                        {
                        }

                        // 粘在空壳实例上时强制重绑（多实例：GetActiveObject 常返回无簿的那个）
                        if (existingCount == 0)
                        {
                            EasyWriteDiagnostics.Log(DebugCategory.OpenFiles,
                                "[ExcelOpenFilesDetector] attached but workbooks=0 — rebind");
                            TearDownUnlocked(raiseDetached: false);
                        }
                        else
                        {
                            return true;
                        }
                    }
                    else
                    {
                        TearDownUnlocked(raiseDetached: false);
                    }
                }

                Excel.Application app = null;
                object raw = null;
                try
                {
                    raw = _resolveExcelApp?.Invoke();
                    app = raw as Excel.Application;
                    if (app == null && raw != null)
                    {
                        try
                        {
                            app = (Excel.Application)raw;
                        }
                        catch (Exception castEx)
                        {
                            EasyWriteDiagnostics.Log(DebugCategory.OpenFiles,
                                "[ExcelOpenFilesDetector] resolve cast failed type="
                                + raw.GetType().FullName + " err=" + castEx.Message);
                        }
                    }
                }
                catch (Exception ex)
                {
                    EasyWriteDiagnostics.Log(DebugCategory.OpenFiles,
                        "[ExcelOpenFilesDetector] resolve failed: " + ex.Message);
                    app = null;
                }

                if (app == null)
                {
                    EasyWriteDiagnostics.Log(DebugCategory.OpenFiles,
                        "[ExcelOpenFilesDetector] resolve returned null"
                        + (raw == null ? "" : " (raw type=" + raw.GetType().FullName + ")"));
                    return false;
                }

                try
                {
                    var _ = app.Name;
                }
                catch (Exception ex)
                {
                    EasyWriteDiagnostics.Log(DebugCategory.OpenFiles,
                        "[ExcelOpenFilesDetector] app not alive: " + ex.Message);
                    return false;
                }

                int wbCount = -1;
                try
                {
                    wbCount = app.Workbooks.Count;
                }
                catch (Exception)
                {
                }

                _app = app;
                try
                {
                    var events = (Excel.AppEvents_Event)_app;
                    events.WorkbookOpen += OnWorkbookOpen;
                    events.NewWorkbook += OnNewWorkbook;
                    events.WorkbookBeforeClose += OnWorkbookBeforeClose;
                    _subscribed = true;
                    EasyWriteDiagnostics.Log(DebugCategory.OpenFiles,
                        "[ExcelOpenFilesDetector] attached and subscribed name="
                        + (app.Name ?? "") + " workbooks=" + wbCount);
                    return true;
                }
                catch (Exception ex)
                {
                    EasyWriteDiagnostics.Log(DebugCategory.OpenFiles,
                        "[ExcelOpenFilesDetector] subscribe failed (will not attach): "
                        + ex.Message + " workbooks=" + wbCount);
                    TearDownUnlocked(raiseDetached: false);
                    return false;
                }
            }
        }

        public IReadOnlyList<OpenFileItem> Snapshot()
        {
            var result = new List<OpenFileItem>();
            Excel.Application app;
            lock (_gate)
            {
                if (!_subscribed || _app == null || !IsAppAliveUnlocked())
                {
                    EasyWriteDiagnostics.Log(DebugCategory.OpenFiles,
                        "[ExcelOpenFilesDetector] Snapshot skip subscribed="
                        + _subscribed + " appNull=" + (_app == null));
                    return result;
                }

                app = _app;
                _rcwToId.Clear();
            }

            var usedIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            int rawCount = 0;
            int mapped = 0;
            int skipped = 0;
            try
            {
                foreach (Excel.Workbook book in app.Workbooks)
                {
                    rawCount++;
                    try
                    {
                        var item = MapWorkbook(book, usedIds);
                        if (item == null)
                        {
                            skipped++;
                            string skipName = null;
                            try
                            {
                                skipName = book.Name;
                            }
                            catch (Exception)
                            {
                            }

                            EasyWriteDiagnostics.Log(DebugCategory.OpenFiles,
                                "[ExcelOpenFilesDetector] Snapshot skip book=" + (skipName ?? "?"));
                            continue;
                        }

                        EnsureChannel(book, item);
                        int key = RuntimeHelpers.GetHashCode(book);
                        lock (_gate)
                        {
                            _rcwToId[key] = item.Id;
                        }

                        result.Add(item);
                        mapped++;
                    }
                    catch (Exception ex)
                    {
                        skipped++;
                        EasyWriteDiagnostics.Log(DebugCategory.OpenFiles,
                            "[ExcelOpenFilesDetector] snapshot item skip: " + ex.Message);
                    }
                }
            }
            catch (Exception ex)
            {
                EasyWriteDiagnostics.Log(DebugCategory.OpenFiles,
                    "[ExcelOpenFilesDetector] snapshot failed: " + ex.Message);
                MarkDetached();
            }

            EasyWriteDiagnostics.Log(DebugCategory.OpenFiles,
                "[ExcelOpenFilesDetector] Snapshot raw=" + rawCount
                + " mapped=" + mapped + " skipped=" + skipped
                + " ids=[" + string.Join(",", result.ConvertAll(i => i.Id)) + "]");
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
                    string uuid = ExcelWorkbookIdentity.TryResolveUuid(book);
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
                    "[ExcelOpenFilesDetector] WorkbookBeforeClose: " + ex.Message);
            }
        }

        private void HandleOpened(Excel.Workbook book)
        {
            if (book == null)
            {
                return;
            }

            try
            {
                string openName = null;
                try
                {
                    openName = book.Name;
                }
                catch (Exception)
                {
                }

                EasyWriteDiagnostics.Log(DebugCategory.OpenFiles,
                    "[ExcelOpenFilesDetector] open event book=" + (openName ?? "?"));

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
                    EasyWriteDiagnostics.Log(DebugCategory.OpenFiles,
                        "[ExcelOpenFilesDetector] open event mapped null book=" + (openName ?? "?"));
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
                    "[ExcelOpenFilesDetector] open event: " + ex.Message);
                if (IsComDead(ex))
                {
                    MarkDetached();
                }
            }
        }

        private static void EnsureChannel(Excel.Workbook book, OpenFileItem item)
        {
            if (book == null || item == null)
            {
                return;
            }

            try
            {
                ExcelChannel channel = ChannelRegistry.CreateOrGetExcel(
                    book,
                    item.FullPath,
                    claimDefaultIfEmpty: true);
                item.ChannelId = channel?.ChannelId;
            }
            catch (Exception ex)
            {
                EasyWriteDiagnostics.Log(DebugCategory.OpenFiles,
                    "[ExcelOpenFilesDetector] EnsureChannel: " + ex.Message);
            }
        }

        private static OpenFileItem MapWorkbook(Excel.Workbook book, HashSet<string> usedIds)
        {
            if (book == null)
            {
                return null;
            }

            string name = null;
            string fullName = null;
            bool isAddin = false;
            try
            {
                name = book.Name;
                fullName = book.FullName;
                isAddin = book.IsAddin;
            }
            catch (Exception)
            {
                return null;
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

                id = "excel:" + fullPath;
            }
            else
            {
                if (string.IsNullOrWhiteSpace(name))
                {
                    name = "未命名工作簿";
                }

                display = name;
                id = "excel:unsaved:" + name;
                int n = 2;
                while (usedIds != null && usedIds.Contains(id))
                {
                    id = "excel:unsaved:" + name + "#" + n;
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
                IsSaved = saved
            };
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
                    var events = (Excel.AppEvents_Event)_app;
                    events.WorkbookOpen -= OnWorkbookOpen;
                    events.NewWorkbook -= OnNewWorkbook;
                    events.WorkbookBeforeClose -= OnWorkbookBeforeClose;
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

        private static bool IsComDead(Exception ex)
        {
            if (ex is COMException || ex is InvalidComObjectException)
            {
                return true;
            }

            string msg = ex.Message ?? "";
            return msg.IndexOf("RPC", StringComparison.OrdinalIgnoreCase) >= 0
                || msg.IndexOf("COM", StringComparison.OrdinalIgnoreCase) >= 0;
        }
    }
}
