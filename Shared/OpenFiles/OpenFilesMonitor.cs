using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading;

namespace WordAddIn1.OpenFiles
{
    /// <summary>
    /// 聚合「打开文件」探测器：内存列表、进程晚启动重附着、Changed 回调。
    /// </summary>
    internal sealed class OpenFilesMonitor : IDisposable
    {
        private const int MaxOpenChannelsItems = 500;
        private static readonly TimeSpan ProcessPollInterval = TimeSpan.FromSeconds(7);

        private readonly Func<object> _resolveWordApp;
        private readonly SynchronizationContext _sync;
        private readonly object _gate = new object();
        private readonly Dictionary<string, OpenFileItem> _items =
            new Dictionary<string, OpenFileItem>(StringComparer.OrdinalIgnoreCase);

        private WordOpenFilesDetector _wordDetector;
        private Timer _processTimer;
        private bool _started;
        private bool _disposed;
        private bool _wordEnabled;
        private string[] _wordProcessNames = { "WINWORD" };

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
        }

        /// <summary>
        /// Desktop 聊天请求体 <c>open_channels</c>（snake_case，与后端一致）。
        /// </summary>
        public object BuildOpenChannelsPayload()
        {
            List<OpenFileItem> items;
            lock (_gate)
            {
                items = OrderedCopyUnlocked();
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

            var timer = Interlocked.Exchange(ref _processTimer, null);
            timer?.Dispose();

            WordOpenFilesDetector word;
            lock (_gate)
            {
                _started = false;
                word = _wordDetector;
                _wordDetector = null;
                _items.Clear();
            }

            if (word != null)
            {
                Unhook(word);
                word.Dispose();
            }
        }

        private void ConfigureFromConfigUnlocked()
        {
            _wordEnabled = false;
            _wordProcessNames = new[] { "WINWORD" };

            var settings = ConfigManager.Config?.OpenFiles;
            var apps = settings?.Apps;
            if (apps == null)
            {
                return;
            }

            foreach (var entry in apps)
            {
                if (entry == null || string.IsNullOrWhiteSpace(entry.Type))
                {
                    continue;
                }

                string type = entry.Type.Trim().ToLowerInvariant();
                if (type == "word")
                {
                    _wordEnabled = true;
                    if (entry.ProcessNames != null && entry.ProcessNames.Count > 0)
                    {
                        _wordProcessNames = entry.ProcessNames
                            .Where(p => !string.IsNullOrWhiteSpace(p))
                            .Select(p => p.Trim())
                            .ToArray();
                        if (_wordProcessNames.Length == 0)
                        {
                            _wordProcessNames = new[] { "WINWORD" };
                        }
                    }
                }
                else
                {
                    EasyWriteDiagnostics.Log(DebugCategory.OpenFiles,
                        "[OpenFilesMonitor] unknown App Type ignored: " + entry.Type);
                }
            }
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
            else
            {
                EasyWriteDiagnostics.Log(DebugCategory.OpenFiles,
                    "[OpenFilesMonitor] word not in Apps whitelist");
            }
        }

        private void Unhook(WordOpenFilesDetector word)
        {
            word.DocumentOpened -= OnDocumentOpened;
            word.DocumentClosed -= OnDocumentClosed;
            word.Detached -= OnWordDetached;
        }

        private void TryAttachAndSnapshot()
        {
            WordOpenFilesDetector word;
            lock (_gate)
            {
                word = _wordDetector;
            }

            if (word == null)
            {
                RaiseChanged();
                return;
            }

            try
            {
                if (!word.TryAttach())
                {
                    return;
                }

                var snapshot = word.Snapshot();
                lock (_gate)
                {
                    RemoveByAppTypeUnlocked("word");
                    foreach (var item in snapshot)
                    {
                        if (item != null && !string.IsNullOrEmpty(item.Id))
                        {
                            _items[item.Id] = item;
                        }
                    }
                }

                RaiseChanged();
            }
            catch (Exception ex)
            {
                EasyWriteDiagnostics.Log(DebugCategory.OpenFiles,
                    "[OpenFilesMonitor] TryAttachAndSnapshot: " + ex.Message);
            }
        }

        private void StartProcessWatch()
        {
            if (!_wordEnabled)
            {
                return;
            }

            _processTimer = new Timer(
                _ => OnProcessWatchTick(),
                null,
                ProcessPollInterval,
                ProcessPollInterval);
        }

        private void OnProcessWatchTick()
        {
            if (_disposed || !_started)
            {
                return;
            }

            try
            {
                bool wordRunning = IsAnyProcessRunning(_wordProcessNames);
                WordOpenFilesDetector word;
                bool attached;
                lock (_gate)
                {
                    word = _wordDetector;
                    attached = word != null && word.IsAttached;
                }

                if (!wordRunning)
                {
                    if (attached || HasAppItems("word"))
                    {
                        EasyWriteDiagnostics.Log(DebugCategory.OpenFiles,
                            "[OpenFilesMonitor] WINWORD gone — clear word items");
                        if (word != null)
                        {
                            lock (_gate)
                            {
                                if (ReferenceEquals(_wordDetector, word))
                                {
                                    Unhook(word);
                                    word.Dispose();
                                    _wordDetector = new WordOpenFilesDetector(_resolveWordApp);
                                    _wordDetector.DocumentOpened += OnDocumentOpened;
                                    _wordDetector.DocumentClosed += OnDocumentClosed;
                                    _wordDetector.Detached += OnWordDetached;
                                }
                            }
                        }

                        bool removed;
                        lock (_gate)
                        {
                            removed = RemoveByAppTypeUnlocked("word");
                        }

                        if (removed)
                        {
                            RaiseChanged();
                        }
                    }

                    return;
                }

                if (!attached)
                {
                    EasyWriteDiagnostics.Log(DebugCategory.OpenFiles,
                        "[OpenFilesMonitor] WINWORD present — try attach");
                    TryAttachAndSnapshot();
                }
            }
            catch (Exception ex)
            {
                EasyWriteDiagnostics.Log(DebugCategory.OpenFiles,
                    "[OpenFilesMonitor] process watch: " + ex.Message);
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
            bool removed;
            lock (_gate)
            {
                removed = RemoveByAppTypeUnlocked("word");
            }

            if (removed)
            {
                RaiseChanged();
            }
        }

        private bool RemoveByAppTypeUnlocked(string appType)
        {
            var toRemove = _items
                .Where(kv => kv.Value != null
                    && string.Equals(kv.Value.AppType, appType, StringComparison.OrdinalIgnoreCase))
                .ToList();

            foreach (var kv in toRemove)
            {
                TryRemoveChannel(kv.Value);
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
