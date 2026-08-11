using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading;

namespace WordAddIn1.OpenFiles
{
    /// <summary>
    /// 聚合「打开文件」探测器：Word + WPS 并行、进程晚启动重附着、Changed 回调。
    /// </summary>
    internal sealed class OpenFilesMonitor : IDisposable
    {
        private const int MaxOpenChannelsItems = 500;
        private static readonly TimeSpan ProcessPollInterval = TimeSpan.FromSeconds(7);
        private static readonly TimeSpan WpsReconcileInterval = TimeSpan.FromSeconds(30);

        private readonly Func<object> _resolveWordApp;
        private readonly SynchronizationContext _sync;
        private readonly object _gate = new object();
        private readonly Dictionary<string, OpenFileItem> _items =
            new Dictionary<string, OpenFileItem>(StringComparer.OrdinalIgnoreCase);

        private WordOpenFilesDetector _wordDetector;
        private WpsOpenFilesDetector _wpsDetector;
        private Timer _processTimer;
        private Timer _wpsReconcileTimer;
        private bool _started;
        private bool _disposed;
        private bool _wordEnabled;
        private bool _wpsEnabled;
        private string[] _wordProcessNames = { "WINWORD" };
        private string[] _wpsProcessNames = { "wps" };

        public OpenFilesMonitor(Func<object> resolveWordApp, SynchronizationContext syncContext = null)
        {
            _resolveWordApp = resolveWordApp ?? throw new ArgumentNullException(nameof(resolveWordApp));
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
            UpdateWpsReconcileTimer();
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

        /// <summary>Word 已由其它路径附着时尽快同步（非列表轮询）。</summary>
        public void TryAttachNow()
        {
            if (_disposed || !_started)
            {
                return;
            }

            TryAttachAndSnapshot();
            UpdateWpsReconcileTimer();
        }

        /// <summary>
        /// Desktop 聊天请求体 <c>open_channels</c>：仅 word 项（排除 wps）。
        /// </summary>
        public object BuildOpenChannelsPayload()
        {
            List<OpenFileItem> items;
            lock (_gate)
            {
                items = OrderedCopyUnlocked()
                    .Where(i => i != null
                        && string.Equals(i.AppType, WordOpenFilesDetector.TypeKey, StringComparison.OrdinalIgnoreCase))
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
            Interlocked.Exchange(ref _wpsReconcileTimer, null)?.Dispose();

            WordOpenFilesDetector word;
            WpsOpenFilesDetector wps;
            lock (_gate)
            {
                _started = false;
                word = _wordDetector;
                wps = _wpsDetector;
                _wordDetector = null;
                _wpsDetector = null;
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
        }

        private void ConfigureFromConfigUnlocked()
        {
            _wordEnabled = false;
            _wpsEnabled = false;
            _wordProcessNames = new[] { "WINWORD" };
            _wpsProcessNames = new[] { "wps" };

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

        private void TryAttachAndSnapshot()
        {
            bool any = false;
            any |= TryAttachOne(WordOpenFilesDetector.TypeKey);
            any |= TryAttachOne(WpsOpenFilesDetector.TypeKey);
            if (any || (!_wordEnabled && !_wpsEnabled))
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
            if (!_wordEnabled && !_wpsEnabled)
            {
                return;
            }

            _processTimer = new Timer(
                _ => OnProcessWatchTick(),
                null,
                ProcessPollInterval,
                ProcessPollInterval);
        }

        private void UpdateWpsReconcileTimer()
        {
            WpsOpenFilesDetector wps;
            lock (_gate)
            {
                wps = _wpsDetector;
            }

            // I6：仅当 WPS 已附着且事件未订成功时才对账
            bool needReconcile = wps != null && wps.IsAttached && !wps.EventsSubscribed;
            if (needReconcile)
            {
                if (_wpsReconcileTimer == null)
                {
                    EasyWriteDiagnostics.Log(DebugCategory.OpenFiles,
                        "[OpenFilesMonitor] start WPS reconcile timer 30s");
                    _wpsReconcileTimer = new Timer(
                        _ => OnWpsReconcileTick(),
                        null,
                        WpsReconcileInterval,
                        WpsReconcileInterval);
                }
            }
            else
            {
                var t = Interlocked.Exchange(ref _wpsReconcileTimer, null);
                if (t != null)
                {
                    EasyWriteDiagnostics.Log(DebugCategory.OpenFiles,
                        "[OpenFilesMonitor] stop WPS reconcile timer");
                    t.Dispose();
                }
            }
        }

        private void OnWpsReconcileTick()
        {
            if (_disposed || !_started)
            {
                return;
            }

            try
            {
                WpsOpenFilesDetector wps;
                lock (_gate)
                {
                    wps = _wpsDetector;
                }

                if (wps == null || !wps.IsAttached || wps.EventsSubscribed)
                {
                    UpdateWpsReconcileTimer();
                    return;
                }

                EasyWriteDiagnostics.Log(DebugCategory.OpenFiles,
                    "[OpenFilesMonitor] WPS reconcile snapshot");
                if (TryAttachOne(WpsOpenFilesDetector.TypeKey))
                {
                    RaiseChanged();
                }

                UpdateWpsReconcileTimer();
            }
            catch (Exception ex)
            {
                EasyWriteDiagnostics.Log(DebugCategory.OpenFiles,
                    "[OpenFilesMonitor] WPS reconcile: " + ex.Message);
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

                if (changed)
                {
                    RaiseChanged();
                }

                UpdateWpsReconcileTimer();
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

            // WPS 永不带渠道
            if (string.Equals(item.AppType, WpsOpenFilesDetector.TypeKey, StringComparison.OrdinalIgnoreCase))
            {
                item.ChannelId = null;
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
            bool removed;
            lock (_gate)
            {
                removed = RemoveByAppTypeUnlocked(WordOpenFilesDetector.TypeKey, removeChannels: true);
            }

            if (removed)
            {
                RaiseChanged();
            }
        }

        private void OnWpsDetached()
        {
            bool removed;
            lock (_gate)
            {
                removed = RemoveByAppTypeUnlocked(WpsOpenFilesDetector.TypeKey, removeChannels: false);
            }

            UpdateWpsReconcileTimer();

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
