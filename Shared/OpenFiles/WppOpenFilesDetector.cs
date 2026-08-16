using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.CompilerServices;

namespace WordAddIn1.OpenFiles
{
    /// <summary>
    /// WPS 演示「打开文件」探测器：只附着已运行实例；建渠不抢默认。
    /// </summary>
    internal sealed class WppOpenFilesDetector : IOpenFilesAppDetector
    {
        public const string TypeKey = "wpp";

        private readonly object _gate = new object();
        private object _app;
        private bool _disposed;
        private readonly Dictionary<int, string> _rcwToId = new Dictionary<int, string>();

        public string AppType => TypeKey;

        public bool EventsSubscribed => false;

        public bool IsAttached
        {
            get
            {
                lock (_gate)
                {
                    return _app != null && WppCom.IsAlive(_app);
                }
            }
        }

        public event Action<OpenFileItem> DocumentOpened;
        public event Action<string> DocumentClosed;
        public event Action Detached;

        // 本期靠 Monitor 定时 Snapshot 合并条目；事件留给后续 ProgId 事件订阅。
        private void RaiseOpenedForReconcile(OpenFileItem item)
        {
            if (item != null)
            {
                DocumentOpened?.Invoke(item);
            }
        }

        private void RaiseClosedForReconcile(string id)
        {
            if (!string.IsNullOrEmpty(id))
            {
                DocumentClosed?.Invoke(id);
            }
        }

        public bool TryAttach()
        {
            if (_disposed)
            {
                return false;
            }

            lock (_gate)
            {
                if (_app != null && WppCom.IsAlive(_app))
                {
                    return true;
                }

                TearDownUnlocked(raiseDetached: false);
                object app = WppCom.TryGetActiveApplication(out string progId);
                if (app == null)
                {
                    return false;
                }

                _app = app;
                EasyWriteDiagnostics.Log(DebugCategory.OpenFiles,
                    "[WppOpenFilesDetector] attached progId=" + progId);
                return true;
            }
        }

        public IReadOnlyList<OpenFileItem> Snapshot()
        {
            var result = new List<OpenFileItem>();
            object app;
            lock (_gate)
            {
                if (_app == null || !WppCom.IsAlive(_app))
                {
                    return result;
                }

                app = _app;
                _rcwToId.Clear();
            }

            var usedIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            try
            {
                foreach (object presentation in WppCom.EnumeratePresentations(app))
                {
                    try
                    {
                        var item = MapPresentation(presentation, usedIds);
                        if (item == null)
                        {
                            // 跳过项勿 FinalRelease：与打开路径/渠道可能共享同一 COM
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
                            "[WppOpenFilesDetector] snapshot item skip: " + ex.Message);
                    }
                }
            }
            catch (Exception ex)
            {
                EasyWriteDiagnostics.Log(DebugCategory.OpenFiles,
                    "[WppOpenFilesDetector] snapshot failed: " + ex.Message);
                MarkDetached();
                return result;
            }

            if (result.Count == 0)
            {
                EasyWriteDiagnostics.Log(DebugCategory.OpenFiles,
                    "[WppOpenFilesDetector] no presentations — release Application RCW");
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

        private static void EnsureChannel(object presentation, OpenFileItem item)
        {
            if (presentation == null || item == null)
            {
                return;
            }

            try
            {
                WppChannel channel = ChannelRegistry.CreateOrGetWpp(
                    presentation,
                    item.FullPath,
                    claimDefaultIfEmpty: false);
                item.ChannelId = channel?.ChannelId;
            }
            catch (Exception ex)
            {
                EasyWriteDiagnostics.Log(DebugCategory.OpenFiles,
                    "[WppOpenFilesDetector] EnsureChannel: " + ex.Message);
                item.ChannelId = null;
            }
        }

        private static OpenFileItem MapPresentation(object presentation, HashSet<string> usedIds)
        {
            if (presentation == null)
            {
                return null;
            }

            string name = WppCom.TryReadName(presentation);
            string fullName = WppCom.TryReadFullName(presentation);
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

                id = "wpp:" + fullPath;
            }
            else
            {
                if (string.IsNullOrWhiteSpace(name))
                {
                    name = "演示文稿1";
                }

                display = name;
                id = "wpp:unsaved:" + name;
                int n = 2;
                while (usedIds != null && usedIds.Contains(id))
                {
                    id = "wpp:unsaved:" + name + "#" + n;
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

        private void MarkDetached()
        {
            lock (_gate)
            {
                TearDownUnlocked(raiseDetached: true);
            }
        }

        private void TearDownUnlocked(bool raiseDetached)
        {
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
