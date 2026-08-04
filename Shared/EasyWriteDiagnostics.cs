using System;
using System.Collections.Generic;
using System.Diagnostics;

namespace WordAddIn1
{
    /// <summary>
    /// 分类 Debug 输出：由 config.json → App.DebugCategories 控制（含 all）。
    /// </summary>
    internal static class EasyWriteDiagnostics
    {
        private static readonly object Sync = new object();
        private static HashSet<string> _enabled = new HashSet<string>(StringComparer.Ordinal);
        private static bool _allEnabled;
        private static bool _initialized;

        /// <summary>任一类别启用（含 all）时为 true，供 EasyWriteLog 决定是否镜像文件。</summary>
        public static bool IsAnyEnabled()
        {
            EnsureInitialized();
            lock (Sync)
            {
                return _allEnabled || _enabled.Count > 0;
            }
        }

        public static bool IsEnabled(string category)
        {
            EnsureInitialized();
            string c = NormalizeLogCategory(category);
            lock (Sync)
            {
                if (_allEnabled)
                {
                    return true;
                }

                return _enabled.Contains(c);
            }
        }

        public static void Log(string category, string message)
        {
            if (message == null)
            {
                return;
            }

            string c = NormalizeLogCategory(category);
            if (!IsEnabled(c))
            {
                return;
            }

            Debug.WriteLine($"[debug:{c}] {message}");
        }

        public static void Log(string category, string format, params object[] args)
        {
            if (format == null)
            {
                return;
            }

            string c = NormalizeLogCategory(category);
            if (!IsEnabled(c))
            {
                return;
            }

            string message = (args == null || args.Length == 0) ? format : string.Format(format, args);
            Debug.WriteLine($"[debug:{c}] {message}");
        }

        /// <summary>文档提取详细日志（原 DocumentExtractVerbose）。</summary>
        public static void LogDocumentExtractVerbose(string message)
        {
            Log(DebugCategory.DocumentExtract, message);
        }

        /// <summary>表格行值读写诊断日志。</summary>
        public static void LogTableRowValues(string message, bool verboseOnly = false)
        {
            if (verboseOnly && !IsEnabled(DebugCategory.Table))
            {
                return;
            }

            Log(DebugCategory.Table, message);
        }

        /// <summary>
        /// 步骤耗时测评：using 包住代码块，结束时输出 [debug:timing] step=… elapsed_ms=…
        /// 未启用 timing 类别时为零开销（不启动 Stopwatch）。
        /// </summary>
        public static IDisposable Time(string step)
        {
            if (string.IsNullOrWhiteSpace(step) || !IsEnabled(DebugCategory.Timing))
            {
                return NullTimingScope.Instance;
            }

            return new TimingScope(step.Trim());
        }

        /// <summary>直接记录已测得的耗时（毫秒）。</summary>
        public static void LogTiming(string step, long elapsedMs, string detail = null)
        {
            if (!IsEnabled(DebugCategory.Timing))
            {
                return;
            }

            if (string.IsNullOrEmpty(detail))
            {
                Log(DebugCategory.Timing, $"step={step} elapsed_ms={elapsedMs}");
            }
            else
            {
                Log(DebugCategory.Timing, $"step={step} elapsed_ms={elapsedMs} {detail}");
            }
        }

        private sealed class NullTimingScope : IDisposable
        {
            public static readonly NullTimingScope Instance = new NullTimingScope();

            public void Dispose()
            {
            }
        }

        private sealed class TimingScope : IDisposable
        {
            private readonly string _step;
            private readonly Stopwatch _sw;
            private bool _disposed;

            public TimingScope(string step)
            {
                _step = step;
                _sw = Stopwatch.StartNew();
            }

            public void Dispose()
            {
                if (_disposed)
                {
                    return;
                }

                _disposed = true;
                _sw.Stop();
                LogTiming(_step, _sw.ElapsedMilliseconds);
            }
        }

        /// <summary>从 ConfigManager 刷新启用集合；LoadConfig / ReloadConfig 后调用。</summary>
        public static void RefreshFromConfig()
        {
            IList<string> raw = null;
            try
            {
                raw = ConfigManager.Config?.App?.DebugCategories;
            }
            catch
            {
                raw = null;
            }

            lock (Sync)
            {
                ApplyRawCategoriesUnlocked(raw);
                _initialized = true;
            }
        }

        private static void EnsureInitialized()
        {
            if (_initialized)
            {
                return;
            }

            // 先触发 Config 加载（内部会 RefreshFromConfig）；勿在持有 Sync 时访问 Config，避免死锁
            try
            {
                var _ = ConfigManager.Config;
            }
            catch
            {
                // ignore
            }

            if (_initialized)
            {
                return;
            }

            IList<string> raw = null;
            try
            {
                raw = ConfigManager.Config?.App?.DebugCategories;
            }
            catch
            {
                raw = null;
            }

            lock (Sync)
            {
                if (_initialized)
                {
                    return;
                }

                ApplyRawCategoriesUnlocked(raw);
                _initialized = true;
            }
        }

        private static void ApplyRawCategoriesUnlocked(IList<string> raw)
        {
            var next = new HashSet<string>(StringComparer.Ordinal);
            bool all = false;

            if (raw != null)
            {
                foreach (string item in raw)
                {
                    if (string.IsNullOrWhiteSpace(item))
                    {
                        continue;
                    }

                    string c = item.Trim().ToLowerInvariant();
                    if (c == "all")
                    {
                        all = true;
                        continue;
                    }

                    if (DebugCategory.Known.Contains(c))
                    {
                        next.Add(c);
                    }
                    else
                    {
                        Debug.WriteLine($"[debug:{DebugCategory.Config}] unknown category in config: {c}");
                    }
                }
            }

            _enabled = next;
            _allEnabled = all;
        }

        private static string NormalizeLogCategory(string category)
        {
            if (string.IsNullOrWhiteSpace(category))
            {
                return DebugCategory.Misc;
            }

            string c = category.Trim().ToLowerInvariant();
            if (c == "all" || !DebugCategory.Known.Contains(c))
            {
                return DebugCategory.Misc;
            }

            return c;
        }
    }
}
