using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading;

namespace WordAddIn1.OpenFiles
{
    /// <summary>
    /// 聚合「打开文件」探测器：Word / WPS / Excel / et 并行、进程晚启动重附着、Changed 回调。
    /// </summary>
    internal sealed class OpenFilesMonitor : IDisposable
    {
        private const int MaxOpenChannelsItems = 500;
        private static readonly TimeSpan ProcessPollInterval = TimeSpan.FromSeconds(7);
        private static readonly TimeSpan LateBindReconcileInterval = TimeSpan.FromSeconds(30);

        private readonly Func<object> _resolveWordApp;
        private readonly Func<object> _resolveExcelApp;
        private readonly SynchronizationContext _sync;
        private readonly object _gate = new object();
        private readonly Dictionary<string, OpenFileItem> _items =
            new Dictionary<string, OpenFileItem>(StringComparer.OrdinalIgnoreCase);

        private WordOpenFilesDetector _wordDetector;
        private WpsOpenFilesDetector _wpsDetector;
        private ExcelOpenFilesDetector _excelDetector;
        private EtOpenFilesDetector _etDetector;
        private Timer _processTimer;
        private Timer _lateBindReconcileTimer;
        private bool _started;
        private bool _disposed;
        private bool _wordEnabled;
        private bool _wpsEnabled;
        private bool _excelEnabled;
        private bool _etEnabled;
        private string[] _wordProcessNames = { "WINWORD" };
        private string[] _wpsProcessNames = { "wps" };
        private string[] _excelProcessNames = { "EXCEL" };
        private string[] _etProcessNames = { "et" };

        public OpenFilesMonitor(
            Func<object> resolveWordApp,
            Func<object> resolveExcelApp = null,
            SynchronizationContext syncContext = null)
        {
            _resolveWordApp = resolveWordApp ?? throw new ArgumentNullException(nameof(resolveWordApp));
            _resolveExcelApp = resolveExcelApp
                ?? (() =>
                {
                    if (ExcelApplicationResolver.TryResolve(
                            out object app,
                            out _,
                            createIfMissing: false,
                            makeVisible: false))
                    {
                        return app;
                    }

                    return null;
                });
            _sync = syncContext ?? SynchronizationContext.Current;
        }

        public event Action<IReadOnlyList<OpenFileItem>> Changed;

        public void Start()
        {
            if (_disposed)
            {
                return;
            }

            lock (_gate)
            {
                if (_started)
                {
                    return;
                }

                _started = true;
                ConfigureFromConfigUnlocked();
                CreateDetectorsUnlocked();
            }

            TryAttachAndSnapshot();
            StartProcessWatch();
            UpdateLateBindReconcileTimer();
        }

        public void Stop()
        {
            Dispose();
        }

        public IReadOnlyList<OpenFileItem> GetSnapshot()
        {
            lock (_gate)
            {
                return OrderedCopyUnlocked();
            }
        }

        public void TryAttachNow()
        {
            if (_disposed || !_started)
            {
                return;
            }

            TryAttachAndSnapshot();
            UpdateLateBindReconcileTimer();
        }

        public object BuildOpenChannelsPayload()
        {
            List<OpenFileItem> items;
            lock (_gate)
            {
                items = OrderedCopyUnlocked()
                    .Where(i => i != null && IsKnownAppType(i.AppType))
                    .ToList();
            }

            bool truncated = items.Count > MaxOpenChannelsItems;
            if (truncated)
            {
                items = items.Take(MaxOpenChannelsItems).ToList();
            }

            var payloadItems = items.Select(i => new Dictionary<string, object>
            {
                ["display_name"] = i.DisplayName ?? "",
                ["full_path"] = (object)i.FullPath ?? null,
                ["channel_id"] = i.ChannelId ?? "",
                ["app_type"] = i.AppType ?? "",
                ["is_saved"] = i.IsSaved
            }).ToList();

            return new Dictionary<string, object>
            {
                ["default_channel_id"] = (object)ChannelRegistry.DefaultChannelId ?? null,
                ["items"] = payloadItems,
                ["truncated"] = truncated
            };
        }

        public void Dispose()
        {
            if (_disposed)
            {
                return;
            }

            _disposed = true;
            Interlocked.Exchange(ref _processTimer, null)?.Dispose();
            Interlocked.Exchange(ref _lateBindReconcileTimer, null)?.Dispose();

            WordOpenFilesDetector word;
            WpsOpenFilesDetector wps;
            ExcelOpenFilesDetector excel;
            EtOpenFilesDetector et;
            lock (_gate)
            {
                _started = false;
                word = _wordDetector;
                wps = _wpsDetector;
                excel = _excelDetector;
                et = _etDetector;
                _wordDetector = null;
                _wpsDetector = null;
                _excelDetector = null;
                _etDetector = null;
                _items.Clear();
            }

            if (word != null)
            {
                UnhookWord(word);
                word.Dispose();
            }

            if (wps != null)
            {
                UnhookWps(wps);
                wps.Dispose();
            }

            if (excel != null)
            {
                UnhookExcel(excel);
                excel.Dispose();
            }

            if (et != null)
            {
                UnhookEt(et);
                et.Dispose();
            }
        }

        private static bool IsKnownAppType(string appType)
        {
            return string.Equals(appType, WordOpenFilesDetector.TypeKey, StringComparison.OrdinalIgnoreCase)
                || string.Equals(appType, WpsOpenFilesDetector.TypeKey, StringComparison.OrdinalIgnoreCase)
                || string.Equals(appType, ExcelOpenFilesDetector.TypeKey, StringComparison.OrdinalIgnoreCase)
                || string.Equals(appType, EtOpenFilesDetector.TypeKey, StringComparison.OrdinalIgnoreCase);
        }

        private void ConfigureFromConfigUnlocked()
        {
            _wordEnabled = false;
            _wpsEnabled = false;
            _excelEnabled = false;
            _etEnabled = false;
            _wordProcessNames = new[] { "WINWORD" };
            _wpsProcessNames = new[] { "wps" };
            _excelProcessNames = new[] { "EXCEL" };
            _etProcessNames = new[] { "et" };

            var apps = ConfigManager.Config?.OpenFiles?.Apps;
            if (apps == null || apps.Count == 0)
            {
                EasyWriteDiagnostics.Log(DebugCategory.OpenFiles,
                    "[OpenFilesMonitor] Apps empty — no detectors");
                return;
            }

            foreach (var entry in apps)
            {
                if (entry == null || string.IsNullOrWhiteSpace(entry.Type))
                {
                    continue;
                }

                string type = entry.Type.Trim().ToLowerInvariant();
                if (type == WordOpenFilesDetector.TypeKey)
                {
                    _wordEnabled = true;
                    _wordProcessNames = ReadProcessNames(entry.ProcessNames, "WINWORD");
                }
                else if (type == WpsOpenFilesDetector.TypeKey)
                {
                    _wpsEnabled = true;
                    _wpsProcessNames = ReadProcessNames(entry.ProcessNames, "wps");
                }
                else if (type == ExcelOpenFilesDetector.TypeKey)
                {
                    _excelEnabled = true;
                    _excelProcessNames = ReadProcessNames(entry.ProcessNames, "EXCEL");
                }
                else if (type == EtOpenFilesDetector.TypeKey)
                {
                    _etEnabled = true;
                    _etProcessNames = ReadProcessNames(entry.ProcessNames, "et");
                }
                else
                {
                    EasyWriteDiagnostics.Log(DebugCategory.OpenFiles,
                        "[OpenFilesMonitor] unknown App Type ignored: " + entry.Type);
                }
            }
        }

        private static string[] ReadProcessNames(List<string> configured, string fallback)
        {
            if (configured == null || configured.Count == 0)
            {
                return new[] { fallback };
            }

            var names = configured
                .Where(p => !string.IsNullOrWhiteSpace(p))
                .Select(p => p.Trim())
                .ToArray();
            return names.Length > 0 ? names : new[] { fallback };
        }

        private void CreateDetectorsUnlocked()
        {
            if (_wordEnabled)
            {
                _wordDetector = new WordOpenFilesDetector(_resolveWordApp);
                _wordDetector.DocumentOpened += OnDocumentOpened;
                _wordDetector.DocumentClosed += OnDocumentClosed;
                _wordDetector.Detached += OnWordDetached;
            }

            if (_wpsEnabled)
            {
                _wpsDetector = new WpsOpenFilesDetector();
                _wpsDetector.DocumentOpened += OnDocumentOpened;
                _wpsDetector.DocumentClosed += OnDocumentClosed;
                _wpsDetector.Detached += OnWpsDetached;
            }

            if (_excelEnabled)
            {
                _excelDetector = new ExcelOpenFilesDetector(_resolveExcelApp);
                _excelDetector.DocumentOpened += OnDocumentOpened;
                _excelDetector.DocumentClosed += OnDocumentClosed;
                _excelDetector.Detached += OnExcelDetached;
            }

            if (_etEnabled)
            {
                _etDetector = new EtOpenFilesDetector();
                _etDetector.DocumentOpened += OnDocumentOpened;
                _etDetector.DocumentClosed += OnDocumentClosed;
                _etDetector.Detached += OnEtDetached;
            }
        }

        private void UnhookWord(WordOpenFilesDetector word)
        {
            word.DocumentOpened -= OnDocumentOpened;
            word.DocumentClosed -= OnDocumentClosed;
            word.Detached -= OnWordDetached;
        }

        private void UnhookWps(WpsOpenFilesDetector wps)
        {
            wps.DocumentOpened -= OnDocumentOpened;
            wps.DocumentClosed -= OnDocumentClosed;
            wps.Detached -= OnWpsDetached;
        }

        private void UnhookExcel(ExcelOpenFilesDetector excel)
        {
            excel.DocumentOpened -= OnDocumentOpened;
            excel.DocumentClosed -= OnDocumentClosed;
            excel.Detached -= OnExcelDetached;
        }

        private void UnhookEt(EtOpenFilesDetector et)
        {
            et.DocumentOpened -= OnDocumentOpened;
            et.DocumentClosed -= OnDocumentClosed;
            et.Detached -= OnEtDetached;
        }

        private void TryAttachAndSnapshot()
        {
            bool any = false;
            any |= TryAttachOne(WordOpenFilesDetector.TypeKey);
            any |= TryAttachOne(WpsOpenFilesDetector.TypeKey);
            any |= TryAttachOne(ExcelOpenFilesDetector.TypeKey);
            any |= TryAttachOne(EtOpenFilesDetector.TypeKey);
            if (any || (!_wordEnabled && !_wpsEnabled && !_excelEnabled && !_etEnabled))
            {
                RaiseChanged();
            }
        }

        private bool TryAttachOne(string appType)
        {
            IOpenFilesAppDetector detector;
            lock (_gate)
            {
                if (string.Equals(appType, WordOpenFilesDetector.TypeKey, StringComparison.OrdinalIgnoreCase))
                {
                    detector = _wordDetector;
                }
                else if (string.Equals(appType, WpsOpenFilesDetector.TypeKey, StringComparison.OrdinalIgnoreCase))
                {
                    detector = _wpsDetector;
                }
                else if (string.Equals(appType, ExcelOpenFilesDetector.TypeKey, StringComparison.OrdinalIgnoreCase))
                {
                    detector = _excelDetector;
                }
                else if (string.Equals(appType, EtOpenFilesDetector.TypeKey, StringComparison.OrdinalIgnoreCase))
                {
                    detector = _etDetector;
                }
                else
                {
                    return false;
                }
            }

            if (detector == null)
            {
                return false;
            }

            try
            {
                if (!detector.TryAttach())
                {
                    return false;
                }

                var snapshot = detector.Snapshot();
                lock (_gate)
                {
                    RemoveByAppTypeUnlocked(appType, removeChannels: true);
                    foreach (var item in snapshot)
                    {
                        if (item != null && !string.IsNullOrEmpty(item.Id))
                        {
                            _items[item.Id] = item;
                        }
                    }
                }

                return true;
            }
            catch (Exception ex)
            {
                EasyWriteDiagnostics.Log(DebugCategory.OpenFiles,
                    "[OpenFilesMonitor] TryAttach " + appType + ": " + ex.Message);
                return false;
            }
        }

        private void StartProcessWatch()
        {
            if (!_wordEnabled && !_wpsEnabled && !_excelEnabled && !_etEnabled)
            {
                return;
            }

            _processTimer = new Timer(
                _ => OnProcessWatchTick(),
                null,
                ProcessPollInterval,
                ProcessPollInterval);
        }

        private void UpdateLateBindReconcileTimer()
        {
            WpsOpenFilesDetector wps;
            EtOpenFilesDetector et;
            lock (_gate)
            {
                wps = _wpsDetector;
                et = _etDetector;
            }

            bool needReconcile =
                (wps != null && wps.IsAttached && !wps.EventsSubscribed)
                || (et != null && et.IsAttached && !et.EventsSubscribed);
            if (needReconcile)
            {
                if (_lateBindReconcileTimer == null)
                {
                    EasyWriteDiagnostics.Log(DebugCategory.OpenFiles,
                        "[OpenFilesMonitor] start late-bind reconcile timer 30s");
                    _lateBindReconcileTimer = new Timer(
                        _ => OnLateBindReconcileTick(),
                        null,
                        LateBindReconcileInterval,
                        LateBindReconcileInterval);
                }
            }
            else
            {
                var t = Interlocked.Exchange(ref _lateBindReconcileTimer, null);
                if (t != null)
                {
                    EasyWriteDiagnostics.Log(DebugCategory.OpenFiles,
                        "[OpenFilesMonitor] stop late-bind reconcile timer");
                    t.Dispose();
                }
            }
        }

        private void OnLateBindReconcileTick()
        {
            if (_disposed || !_started)
            {
                return;
            }

            try
            {
                bool changed = false;
                WpsOpenFilesDetector wps;
                EtOpenFilesDetector et;
                lock (_gate)
                {
                    wps = _wpsDetector;
                    et = _etDetector;
                }

                if (wps != null && wps.IsAttached && !wps.EventsSubscribed)
                {
                    changed |= TryAttachOne(WpsOpenFilesDetector.TypeKey);
                }

                if (et != null && et.IsAttached && !et.EventsSubscribed)
                {
                    changed |= TryAttachOne(EtOpenFilesDetector.TypeKey);
                }

                if (changed)
                {
                    RaiseChanged();
                }

                UpdateLateBindReconcileTimer();
            }
            catch (Exception ex)
            {
                EasyWriteDiagnostics.Log(DebugCategory.OpenFiles,
                    "[OpenFilesMonitor] late-bind reconcile: " + ex.Message);
            }
        }

        private void OnProcessWatchTick()
        {
            if (_disposed || !_started)
            {
                return;
            }

            try
            {
                bool changed = false;
                changed |= WatchOne(
                    WordOpenFilesDetector.TypeKey,
                    _wordEnabled,
                    _wordProcessNames,
                    () =>
                    {
                        lock (_gate)
                        {
                            return _wordDetector;
                        }
                    },
                    RecreateWordDetector);

                changed |= WatchOne(
                    WpsOpenFilesDetector.TypeKey,
                    _wpsEnabled,
                    _wpsProcessNames,
                    () =>
                    {
                        lock (_gate)
                        {
                            return _wpsDetector;
                        }
                    },
                    RecreateWpsDetector);

                changed |= WatchOne(
                    ExcelOpenFilesDetector.TypeKey,
                    _excelEnabled,
                    _excelProcessNames,
                    () =>
                    {
                        lock (_gate)
                        {
                            return _excelDetector;
                        }
                    },
                    RecreateExcelDetector);

                changed |= WatchOne(
                    EtOpenFilesDetector.TypeKey,
                    _etEnabled,
                    _etProcessNames,
                    () =>
                    {
                        lock (_gate)
                        {
                            return _etDetector;
                        }
                    },
                    RecreateEtDetector);

                if (changed)
                {
                    RaiseChanged();
                }

                UpdateLateBindReconcileTimer();
            }
            catch (Exception ex)
            {
                EasyWriteDiagnostics.Log(DebugCategory.OpenFiles,
                    "[OpenFilesMonitor] process watch: " + ex.Message);
            }
        }

        private bool WatchOne(
            string appType,
            bool enabled,
            string[] processNames,
            Func<IOpenFilesAppDetector> getDetector,
            Action recreate)
        {
            if (!enabled)
            {
                return false;
            }

            bool running = IsAnyProcessRunning(processNames);
            var detector = getDetector();
            bool attached = detector != null && detector.IsAttached;

            if (!running)
            {
                if (attached || HasAppItems(appType))
                {
                    EasyWriteDiagnostics.Log(DebugCategory.OpenFiles,
                        "[OpenFilesMonitor] " + appType + " process gone — clear items");
                    recreate();
                    lock (_gate)
                    {
                        return RemoveByAppTypeUnlocked(appType, removeChannels: true);
                    }
                }

                return false;
            }

            if (!attached)
            {
                EasyWriteDiagnostics.Log(DebugCategory.OpenFiles,
                    "[OpenFilesMonitor] " + appType + " process present — try attach");
                return TryAttachOne(appType);
            }

            return false;
        }

        private void RecreateWordDetector()
        {
            lock (_gate)
            {
                var word = _wordDetector;
                if (word == null)
                {
                    return;
                }

                UnhookWord(word);
                word.Dispose();
                _wordDetector = new WordOpenFilesDetector(_resolveWordApp);
                _wordDetector.DocumentOpened += OnDocumentOpened;
                _wordDetector.DocumentClosed += OnDocumentClosed;
                _wordDetector.Detached += OnWordDetached;
            }
        }

        private void RecreateWpsDetector()
        {
            lock (_gate)
            {
                var wps = _wpsDetector;
                if (wps == null)
                {
                    return;
                }

                UnhookWps(wps);
                wps.Dispose();
                _wpsDetector = new WpsOpenFilesDetector();
                _wpsDetector.DocumentOpened += OnDocumentOpened;
                _wpsDetector.DocumentClosed += OnDocumentClosed;
                _wpsDetector.Detached += OnWpsDetached;
            }
        }

        private void RecreateExcelDetector()
        {
            lock (_gate)
            {
                var excel = _excelDetector;
                if (excel == null)
                {
                    return;
                }

                UnhookExcel(excel);
                excel.Dispose();
                _excelDetector = new ExcelOpenFilesDetector(_resolveExcelApp);
                _excelDetector.DocumentOpened += OnDocumentOpened;
                _excelDetector.DocumentClosed += OnDocumentClosed;
                _excelDetector.Detached += OnExcelDetached;
            }
        }

        private void RecreateEtDetector()
        {
            lock (_gate)
            {
                var et = _etDetector;
                if (et == null)
                {
                    return;
                }

                UnhookEt(et);
                et.Dispose();
                _etDetector = new EtOpenFilesDetector();
                _etDetector.DocumentOpened += OnDocumentOpened;
                _etDetector.DocumentClosed += OnDocumentClosed;
                _etDetector.Detached += OnEtDetached;
            }
        }

        private bool HasAppItems(string appType)
        {
            lock (_gate)
            {
                return _items.Values.Any(i =>
                    i != null && string.Equals(i.AppType, appType, StringComparison.OrdinalIgnoreCase));
            }
        }

        private static bool IsAnyProcessRunning(IEnumerable<string> processNames)
        {
            if (processNames == null)
            {
                return false;
            }

            foreach (string name in processNames)
            {
                if (string.IsNullOrWhiteSpace(name))
                {
                    continue;
                }

                string key = name.Trim();
                if (key.EndsWith(".exe", StringComparison.OrdinalIgnoreCase))
                {
                    key = key.Substring(0, key.Length - 4);
                }

                try
                {
                    var procs = Process.GetProcessesByName(key);
                    if (procs != null && procs.Length > 0)
                    {
                        foreach (var p in procs)
                        {
                            p.Dispose();
                        }

                        return true;
                    }
                }
                catch (Exception)
                {
                }
            }

            return false;
        }

        private void OnDocumentOpened(OpenFileItem item)
        {
            if (item == null || string.IsNullOrEmpty(item.Id))
            {
                return;
            }

            lock (_gate)
            {
                _items[item.Id] = item;
            }

            RaiseChanged();
        }

        private void OnDocumentClosed(string id)
        {
            if (string.IsNullOrEmpty(id))
            {
                return;
            }

            OpenFileItem removedItem = null;
            bool removed;
            lock (_gate)
            {
                if (_items.TryGetValue(id, out removedItem))
                {
                    _items.Remove(id);
                    removed = true;
                }
                else
                {
                    removed = false;
                }
            }

            TryRemoveChannel(removedItem);
            if (removed)
            {
                RaiseChanged();
            }
        }

        private void OnWordDetached()
        {
            RaiseDetached(WordOpenFilesDetector.TypeKey);
        }

        private void OnWpsDetached()
        {
            RaiseDetached(WpsOpenFilesDetector.TypeKey);
            UpdateLateBindReconcileTimer();
        }

        private void OnExcelDetached()
        {
            RaiseDetached(ExcelOpenFilesDetector.TypeKey);
        }

        private void OnEtDetached()
        {
            RaiseDetached(EtOpenFilesDetector.TypeKey);
            UpdateLateBindReconcileTimer();
        }

        private void RaiseDetached(string appType)
        {
            bool removed;
            lock (_gate)
            {
                removed = RemoveByAppTypeUnlocked(appType, removeChannels: true);
            }

            if (removed)
            {
                RaiseChanged();
            }
        }

        private bool RemoveByAppTypeUnlocked(string appType, bool removeChannels)
        {
            var toRemove = _items
                .Where(kv => kv.Value != null
                    && string.Equals(kv.Value.AppType, appType, StringComparison.OrdinalIgnoreCase))
                .ToList();

            foreach (var kv in toRemove)
            {
                if (removeChannels)
                {
                    TryRemoveChannel(kv.Value);
                }

                _items.Remove(kv.Key);
            }

            return toRemove.Count > 0;
        }

        private static void TryRemoveChannel(OpenFileItem item)
        {
            if (item == null || string.IsNullOrEmpty(item.ChannelId))
            {
                return;
            }

            try
            {
                if (ChannelRegistry.TryGet(item.ChannelId, out IOperationChannel ch)
                    && ch is WpsChannel wps)
                {
                    object doc = wps.Document;
                    if (doc != null)
                    {
                        WpsDocumentIdentity.ClearOnDocumentClose(doc);
                    }
                    else
                    {
                        DocumentState.ClearSessionByUuid(wps.DocUuid);
                    }

                    return;
                }

                ChannelRegistry.Remove(item.ChannelId);
            }
            catch (Exception ex)
            {
                EasyWriteDiagnostics.Log(DebugCategory.OpenFiles,
                    "[OpenFilesMonitor] Remove channel: " + ex.Message);
            }
        }

        private List<OpenFileItem> OrderedCopyUnlocked()
        {
            return _items.Values
                .OrderBy(x => x.DisplayName ?? "", StringComparer.CurrentCultureIgnoreCase)
                .ToList();
        }

        private void RaiseChanged()
        {
            List<OpenFileItem> list;
            lock (_gate)
            {
                list = OrderedCopyUnlocked();
            }

            void Invoke()
            {
                try
                {
                    Changed?.Invoke(list);
                }
                catch (Exception ex)
                {
                    EasyWriteDiagnostics.Log(DebugCategory.OpenFiles,
                        "[OpenFilesMonitor] Changed handler: " + ex.Message);
                }
            }

            if (_sync != null)
            {
                _sync.Post(_ => Invoke(), null);
            }
            else
            {
                Invoke();
            }
        }
    }
}
