using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading;

namespace WordAddIn1.OpenFiles
{
    /// <summary>
    /// 聚合「打开文件」探测器：Word / WPS / Excel / et / ppt / wpp 并行、进程晚启动重附着、Changed 回调。
    /// </summary>
    internal sealed class OpenFilesMonitor : IDisposable
    {
        private const int MaxOpenChannelsItems = 500;
        private static readonly TimeSpan ProcessPollInterval = TimeSpan.FromSeconds(7);
        private static readonly TimeSpan LateBindReconcileInterval = TimeSpan.FromSeconds(30);
        /// <summary>演示文稿无可靠关闭事件，加快重扫（对齐用户对 wpp 关闭滞后的反馈）。</summary>
        private static readonly TimeSpan PresentationPollInterval = TimeSpan.FromSeconds(2);

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
        private PptOpenFilesDetector _pptDetector;
        private WppOpenFilesDetector _wppDetector;
        private Timer _processTimer;
        private Timer _lateBindReconcileTimer;
        private Timer _presentationPollTimer;
        private bool _started;
        private bool _disposed;
        private bool _wordEnabled;
        private bool _wpsEnabled;
        private bool _excelEnabled;
        private bool _etEnabled;
        private bool _pptEnabled;
        private bool _wppEnabled;
        private string[] _wordProcessNames = { "WINWORD" };
        private string[] _wpsProcessNames = { "wps" };
        private string[] _excelProcessNames = { "EXCEL" };
        private string[] _etProcessNames = { "et" };
        private string[] _pptProcessNames = { "POWERPNT" };
        private string[] _wppProcessNames = { "wpp" };

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

            InvokeDetectorOnSync(() =>
            {
                TryAttachAndSnapshot();
                return true;
            });
            StartProcessWatch();
            StartPresentationWatch();
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

        /// <summary>侧栏 / 芯片用：channelId 为助手短号（w1/x1/p1/b1），与 system open_channels 一致。</summary>
        public IReadOnlyList<object> GetFrontendSnapshot()
        {
            return ToFrontendItems(GetSnapshot());
        }

        public static IReadOnlyList<object> ToFrontendItems(IReadOnlyList<OpenFileItem> items)
        {
            if (items == null || items.Count == 0)
            {
                return Array.Empty<object>();
            }

            var list = new List<object>(items.Count);
            for (int i = 0; i < items.Count; i++)
            {
                OpenFileItem item = items[i];
                if (item == null)
                {
                    continue;
                }

                string raw = item.ChannelId;
                string publicId = string.IsNullOrWhiteSpace(raw)
                    ? ""
                    : (ChannelRegistry.ToPublicId(raw) ?? "");
                list.Add(new
                {
                    id = item.Id,
                    appType = item.AppType,
                    displayName = item.DisplayName,
                    fullPath = item.FullPath,
                    isSaved = item.IsSaved,
                    channelId = publicId
                });
            }

            return list;
        }

        public void TryAttachNow()
        {
            if (_disposed || !_started)
            {
                return;
            }

            InvokeDetectorOnSync(() =>
            {
                TryAttachAndSnapshot();
                return true;
            });
            UpdateLateBindReconcileTimer();
        }

        /// <summary>浏览页：按网页粒度写入侧栏（不经 COM 探测）。itemId 只给展示/分组，channelId 才是 b1。</summary>
        public void UpsertBrowserItem(string itemId, string channelId, string displayName, string url)
        {
            if (_disposed || string.IsNullOrWhiteSpace(itemId) || string.IsNullOrWhiteSpace(channelId))
            {
                return;
            }

            string id = itemId.Trim();
            string name = !string.IsNullOrWhiteSpace(displayName)
                ? displayName.Trim()
                : (!string.IsNullOrWhiteSpace(url) ? url.Trim() : "易写浏览器");

            var item = new OpenFileItem
            {
                Id = id,
                AppType = BrowserTypeKey,
                DisplayName = name,
                FullPath = string.IsNullOrWhiteSpace(url) ? null : url.Trim(),
                IsSaved = true,
                ChannelId = channelId.Trim()
            };

            OnDocumentOpened(item);
        }

        /// <summary>浏览页关闭：仅从侧栏移除；渠道由 Host 自行 Remove。参数为侧栏行键。</summary>
        public void RemoveBrowserItem(string itemId)
        {
            if (_disposed || string.IsNullOrWhiteSpace(itemId))
            {
                return;
            }

            string id = itemId.Trim();
            bool removed;
            lock (_gate)
            {
                removed = _items.Remove(id);
            }

            if (removed)
            {
                RaiseChanged();
            }
        }

        public const string BrowserTypeKey = "browser";

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
                ["channel_id"] = ChannelRegistry.ToPublicId(i.ChannelId) ?? "",
                ["app_type"] = i.AppType ?? "",
                ["is_saved"] = i.IsSaved
            }).ToList();

            return new Dictionary<string, object>
            {
                ["default_channel_id"] = (object)ChannelRegistry.PublicDefaultChannelId ?? null,
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
            Interlocked.Exchange(ref _presentationPollTimer, null)?.Dispose();

            WordOpenFilesDetector word;
            WpsOpenFilesDetector wps;
            ExcelOpenFilesDetector excel;
            EtOpenFilesDetector et;
            PptOpenFilesDetector ppt;
            WppOpenFilesDetector wpp;
            lock (_gate)
            {
                _started = false;
                word = _wordDetector;
                wps = _wpsDetector;
                excel = _excelDetector;
                et = _etDetector;
                ppt = _pptDetector;
                wpp = _wppDetector;
                _wordDetector = null;
                _wpsDetector = null;
                _excelDetector = null;
                _etDetector = null;
                _pptDetector = null;
                _wppDetector = null;
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

            if (ppt != null)
            {
                UnhookPpt(ppt);
                ppt.Dispose();
            }

            if (wpp != null)
            {
                UnhookWpp(wpp);
                wpp.Dispose();
            }
        }

        private static bool IsKnownAppType(string appType)
        {
            return string.Equals(appType, WordOpenFilesDetector.TypeKey, StringComparison.OrdinalIgnoreCase)
                || string.Equals(appType, WpsOpenFilesDetector.TypeKey, StringComparison.OrdinalIgnoreCase)
                || string.Equals(appType, ExcelOpenFilesDetector.TypeKey, StringComparison.OrdinalIgnoreCase)
                || string.Equals(appType, EtOpenFilesDetector.TypeKey, StringComparison.OrdinalIgnoreCase)
                || string.Equals(appType, PptOpenFilesDetector.TypeKey, StringComparison.OrdinalIgnoreCase)
                || string.Equals(appType, WppOpenFilesDetector.TypeKey, StringComparison.OrdinalIgnoreCase)
                || string.Equals(appType, BrowserTypeKey, StringComparison.OrdinalIgnoreCase);
        }

        private void ConfigureFromConfigUnlocked()
        {
            _wordEnabled = false;
            _wpsEnabled = false;
            _excelEnabled = false;
            _etEnabled = false;
            _pptEnabled = false;
            _wppEnabled = false;
            _wordProcessNames = new[] { "WINWORD" };
            _wpsProcessNames = new[] { "wps" };
            _excelProcessNames = new[] { "EXCEL" };
            _etProcessNames = new[] { "et" };
            _pptProcessNames = new[] { "POWERPNT" };
            _wppProcessNames = new[] { "wpp" };

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
                else if (type == PptOpenFilesDetector.TypeKey)
                {
                    _pptEnabled = true;
                    _pptProcessNames = ReadProcessNames(entry.ProcessNames, "POWERPNT");
                }
                else if (type == WppOpenFilesDetector.TypeKey)
                {
                    _wppEnabled = true;
                    _wppProcessNames = ReadProcessNames(entry.ProcessNames, "wpp");
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

            if (_pptEnabled)
            {
                _pptDetector = new PptOpenFilesDetector();
                _pptDetector.DocumentOpened += OnDocumentOpened;
                _pptDetector.DocumentClosed += OnDocumentClosed;
                _pptDetector.Detached += OnPptDetached;
            }

            if (_wppEnabled)
            {
                _wppDetector = new WppOpenFilesDetector();
                _wppDetector.DocumentOpened += OnDocumentOpened;
                _wppDetector.DocumentClosed += OnDocumentClosed;
                _wppDetector.Detached += OnWppDetached;
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

        private void UnhookPpt(PptOpenFilesDetector ppt)
        {
            ppt.DocumentOpened -= OnDocumentOpened;
            ppt.DocumentClosed -= OnDocumentClosed;
            ppt.Detached -= OnPptDetached;
        }

        private void UnhookWpp(WppOpenFilesDetector wpp)
        {
            wpp.DocumentOpened -= OnDocumentOpened;
            wpp.DocumentClosed -= OnDocumentClosed;
            wpp.Detached -= OnWppDetached;
        }

        private void TryAttachAndSnapshot()
        {
            bool any = false;
            any |= TryAttachOne(WordOpenFilesDetector.TypeKey);
            any |= TryAttachOne(WpsOpenFilesDetector.TypeKey);
            any |= TryAttachOne(ExcelOpenFilesDetector.TypeKey);
            any |= TryAttachOne(EtOpenFilesDetector.TypeKey);
            any |= TryAttachOne(PptOpenFilesDetector.TypeKey);
            any |= TryAttachOne(WppOpenFilesDetector.TypeKey);
            if (any || (!_wordEnabled && !_wpsEnabled && !_excelEnabled && !_etEnabled && !_pptEnabled && !_wppEnabled))
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
                else if (string.Equals(appType, PptOpenFilesDetector.TypeKey, StringComparison.OrdinalIgnoreCase))
                {
                    detector = _pptDetector;
                }
                else if (string.Equals(appType, WppOpenFilesDetector.TypeKey, StringComparison.OrdinalIgnoreCase))
                {
                    detector = _wppDetector;
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
            if (!_wordEnabled && !_wpsEnabled && !_excelEnabled && !_etEnabled && !_pptEnabled && !_wppEnabled)
            {
                return;
            }

            _processTimer = new Timer(
                _ => OnProcessWatchTick(),
                null,
                ProcessPollInterval,
                ProcessPollInterval);
        }

        /// <summary>ppt/wpp：进程仍在但演示文稿已关时，2s 重扫侧栏，避免干等到 7s 进程轮询。</summary>
        private void StartPresentationWatch()
        {
            if (!_pptEnabled && !_wppEnabled)
            {
                return;
            }

            _presentationPollTimer = new Timer(
                _ => OnPresentationWatchTick(),
                null,
                PresentationPollInterval,
                PresentationPollInterval);
        }

        private void OnPresentationWatchTick()
        {
            if (_disposed || !_started)
            {
                return;
            }

            try
            {
                bool changed = false;
                if (_pptEnabled)
                {
                    changed |= InvokeDetectorOnSync(() =>
                    {
                        PptOpenFilesDetector ppt;
                        lock (_gate)
                        {
                            ppt = _pptDetector;
                        }

                        if (ppt == null)
                        {
                            return false;
                        }

                        if (!ppt.IsAttached)
                        {
                            return TryAttachOne(PptOpenFilesDetector.TypeKey);
                        }

                        return TryResnapshotOne(PptOpenFilesDetector.TypeKey);
                    });
                }

                if (_wppEnabled)
                {
                    changed |= InvokeDetectorOnSync(() =>
                    {
                        WppOpenFilesDetector wpp;
                        lock (_gate)
                        {
                            wpp = _wppDetector;
                        }

                        if (wpp == null)
                        {
                            return false;
                        }

                        if (!wpp.IsAttached)
                        {
                            return TryAttachOne(WppOpenFilesDetector.TypeKey);
                        }

                        return TryResnapshotOne(WppOpenFilesDetector.TypeKey);
                    });
                }

                if (changed)
                {
                    RaiseChanged();
                }
            }
            catch (Exception ex)
            {
                EasyWriteDiagnostics.Log(DebugCategory.OpenFiles,
                    "[OpenFilesMonitor] presentation watch: " + ex.Message);
            }
        }

        private void UpdateLateBindReconcileTimer()
        {
            WpsOpenFilesDetector wps;
            EtOpenFilesDetector et;
            PptOpenFilesDetector ppt;
            WppOpenFilesDetector wpp;
            lock (_gate)
            {
                wps = _wpsDetector;
                et = _etDetector;
                ppt = _pptDetector;
                wpp = _wppDetector;
            }

            bool needReconcile =
                (wps != null && wps.IsAttached && !wps.EventsSubscribed)
                || (et != null && et.IsAttached && !et.EventsSubscribed)
                || (ppt != null && ppt.IsAttached && !ppt.EventsSubscribed)
                || (wpp != null && wpp.IsAttached && !wpp.EventsSubscribed);
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
                PptOpenFilesDetector ppt;
                WppOpenFilesDetector wpp;
                lock (_gate)
                {
                    wps = _wpsDetector;
                    et = _etDetector;
                    ppt = _pptDetector;
                    wpp = _wppDetector;
                }

                if (wps != null && wps.IsAttached && !wps.EventsSubscribed)
                {
                    changed |= TryAttachOne(WpsOpenFilesDetector.TypeKey);
                }

                if (et != null && et.IsAttached && !et.EventsSubscribed)
                {
                    changed |= TryAttachOne(EtOpenFilesDetector.TypeKey);
                }

                if (ppt != null && ppt.IsAttached && !ppt.EventsSubscribed)
                {
                    changed |= TryAttachOne(PptOpenFilesDetector.TypeKey);
                }

                if (wpp != null && wpp.IsAttached && !wpp.EventsSubscribed)
                {
                    changed |= TryAttachOne(WppOpenFilesDetector.TypeKey);
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

                changed |= WatchOne(
                    PptOpenFilesDetector.TypeKey,
                    _pptEnabled,
                    _pptProcessNames,
                    () =>
                    {
                        lock (_gate)
                        {
                            return _pptDetector;
                        }
                    },
                    RecreatePptDetector);

                changed |= WatchOne(
                    WppOpenFilesDetector.TypeKey,
                    _wppEnabled,
                    _wppProcessNames,
                    () =>
                    {
                        lock (_gate)
                        {
                            return _wppDetector;
                        }
                    },
                    RecreateWppDetector);

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
                return InvokeDetectorOnSync(() => TryAttachOne(appType));
            }

            // Excel/et/ppt/wpp：事件不可靠时定期重扫；无文稿时 Snapshot 会释放 Application。
            if (string.Equals(appType, ExcelOpenFilesDetector.TypeKey, StringComparison.OrdinalIgnoreCase)
                || string.Equals(appType, EtOpenFilesDetector.TypeKey, StringComparison.OrdinalIgnoreCase)
                || string.Equals(appType, PptOpenFilesDetector.TypeKey, StringComparison.OrdinalIgnoreCase)
                || string.Equals(appType, WppOpenFilesDetector.TypeKey, StringComparison.OrdinalIgnoreCase))
            {
                return InvokeDetectorOnSync(() => TryResnapshotOne(appType));
            }

            return false;
        }

        private bool InvokeDetectorOnSync(Func<bool> action)
        {
            if (action == null)
            {
                return false;
            }

            if (_sync == null)
            {
                return action();
            }

            bool result = false;
            Exception error = null;
            _sync.Send(
                _ =>
                {
                    try
                    {
                        result = action();
                    }
                    catch (Exception ex)
                    {
                        error = ex;
                    }
                },
                null);

            if (error != null)
            {
                EasyWriteDiagnostics.Log(DebugCategory.OpenFiles,
                    "[OpenFilesMonitor] sync invoke: " + error.Message);
                return false;
            }

            return result;
        }

        /// <summary>
        /// 已附着时重扫；仅当条目增减/变化时返回 true。关闭的簿才 RemoveChannel。
        /// </summary>
        private bool TryResnapshotOne(string appType)
        {
            IOpenFilesAppDetector detector;
            lock (_gate)
            {
                if (string.Equals(appType, ExcelOpenFilesDetector.TypeKey, StringComparison.OrdinalIgnoreCase))
                {
                    detector = _excelDetector;
                }
                else if (string.Equals(appType, EtOpenFilesDetector.TypeKey, StringComparison.OrdinalIgnoreCase))
                {
                    detector = _etDetector;
                }
                else if (string.Equals(appType, PptOpenFilesDetector.TypeKey, StringComparison.OrdinalIgnoreCase))
                {
                    detector = _pptDetector;
                }
                else if (string.Equals(appType, WppOpenFilesDetector.TypeKey, StringComparison.OrdinalIgnoreCase))
                {
                    detector = _wppDetector;
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

                var snapshot = detector.Snapshot() ?? Array.Empty<OpenFileItem>();
                var newById = new Dictionary<string, OpenFileItem>(StringComparer.OrdinalIgnoreCase);
                foreach (var item in snapshot)
                {
                    if (item != null && !string.IsNullOrEmpty(item.Id))
                    {
                        newById[item.Id] = item;
                    }
                }

                lock (_gate)
                {
                    var oldItems = _items
                        .Where(kv => kv.Value != null
                            && string.Equals(kv.Value.AppType, appType, StringComparison.OrdinalIgnoreCase))
                        .ToList();

                    bool changed = false;
                    foreach (var kv in oldItems)
                    {
                        if (!newById.ContainsKey(kv.Key))
                        {
                            if (!ChannelStillHeld(kv.Value, newById.Values))
                            {
                                TryRemoveChannel(kv.Value);
                            }

                            _items.Remove(kv.Key);
                            changed = true;
                        }
                    }

                    foreach (var kv in newById)
                    {
                        if (!_items.TryGetValue(kv.Key, out OpenFileItem existing)
                            || existing == null
                            || !OpenFileItemEquals(existing, kv.Value))
                        {
                            _items[kv.Key] = kv.Value;
                            changed = true;
                        }
                        else if (string.IsNullOrEmpty(existing.ChannelId)
                                 && !string.IsNullOrEmpty(kv.Value.ChannelId))
                        {
                            existing.ChannelId = kv.Value.ChannelId;
                            changed = true;
                        }
                    }

                    return changed;
                }
            }
            catch (Exception ex)
            {
                EasyWriteDiagnostics.Log(DebugCategory.OpenFiles,
                    "[OpenFilesMonitor] resnapshot " + appType + ": " + ex.Message);
                return false;
            }
        }

        private static bool OpenFileItemEquals(OpenFileItem a, OpenFileItem b)
        {
            if (ReferenceEquals(a, b))
            {
                return true;
            }

            if (a == null || b == null)
            {
                return false;
            }

            return string.Equals(a.Id, b.Id, StringComparison.OrdinalIgnoreCase)
                && string.Equals(a.DisplayName, b.DisplayName, StringComparison.Ordinal)
                && string.Equals(a.FullPath, b.FullPath, StringComparison.OrdinalIgnoreCase)
                && string.Equals(a.ChannelId, b.ChannelId, StringComparison.OrdinalIgnoreCase)
                && a.IsSaved == b.IsSaved;
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

        private void RecreatePptDetector()
        {
            lock (_gate)
            {
                var ppt = _pptDetector;
                if (ppt == null)
                {
                    return;
                }

                UnhookPpt(ppt);
                ppt.Dispose();
                _pptDetector = new PptOpenFilesDetector();
                _pptDetector.DocumentOpened += OnDocumentOpened;
                _pptDetector.DocumentClosed += OnDocumentClosed;
                _pptDetector.Detached += OnPptDetached;
            }
        }

        private void RecreateWppDetector()
        {
            lock (_gate)
            {
                var wpp = _wppDetector;
                if (wpp == null)
                {
                    return;
                }

                UnhookWpp(wpp);
                wpp.Dispose();
                _wppDetector = new WppOpenFilesDetector();
                _wppDetector.DocumentOpened += OnDocumentOpened;
                _wppDetector.DocumentClosed += OnDocumentClosed;
                _wppDetector.Detached += OnWppDetached;
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
            bool stillHeld = false;
            lock (_gate)
            {
                if (_items.TryGetValue(id, out removedItem))
                {
                    _items.Remove(id);
                    removed = true;
                    stillHeld = ChannelStillHeld(removedItem, _items.Values);
                }
                else
                {
                    removed = false;
                }
            }

            if (removedItem != null && !stillHeld)
            {
                TryRemoveChannel(removedItem);
            }

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

        private void OnPptDetached()
        {
            RaiseDetached(PptOpenFilesDetector.TypeKey);
            UpdateLateBindReconcileTimer();
        }

        private void OnWppDetached()
        {
            RaiseDetached(WppOpenFilesDetector.TypeKey);
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

        private static bool ChannelStillHeld(OpenFileItem leaving, IEnumerable<OpenFileItem> remaining)
        {
            if (leaving == null || string.IsNullOrEmpty(leaving.ChannelId) || remaining == null)
            {
                return false;
            }

            foreach (OpenFileItem item in remaining)
            {
                if (item == null || string.IsNullOrEmpty(item.ChannelId))
                {
                    continue;
                }

                if (string.Equals(item.ChannelId, leaving.ChannelId, StringComparison.OrdinalIgnoreCase)
                    && !string.Equals(item.Id, leaving.Id, StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
            }

            return false;
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
