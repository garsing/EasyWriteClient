using System;
using System.Diagnostics;
using System.IO;
using System.Text;

namespace WordAddIn1
{
    /// <summary>
    /// 统一文件日志：写入 WordAddIn2/logs/ 下带时间戳的会话文件，并镜像 Debug 输出。
    /// </summary>
    public static class EasyWriteLog
    {
        private static readonly object Sync = new object();
        private static StreamWriter _writer;
        private static EasyWriteTraceListener _listener;
        private static string _logDirectory;

        /// <summary>日志目录（WordAddIn2/logs）。</summary>
        public static string LogDirectory
        {
            get
            {
                if (_logDirectory == null)
                {
                    _logDirectory = ResolveLogDirectory();
                }

                return _logDirectory;
            }
        }

        /// <summary>当前会话日志文件的完整路径。</summary>
        public static string CurrentLogPath { get; private set; }

        /// <summary>
        /// 插件启动时调用：创建 logs 目录、首个会话文件，并监听所有 Debug 输出。
        /// </summary>
        public static void Initialize(string sessionPrefix = "EasyWrite")
        {
            lock (Sync)
            {
                if (_listener != null)
                {
                    return;
                }

                Directory.CreateDirectory(LogDirectory);

                // 任一 DebugCategories 启用（或 all）时才镜像 Debug 到文件
                if (!EasyWriteDiagnostics.IsAnyEnabled())
                {
                    return;
                }

                BeginSessionInternal(sessionPrefix, autoFlush: true);

                _listener = new EasyWriteTraceListener();
                Debug.Listeners.Add(_listener);
            }
        }

        /// <summary>
        /// 开始新的测试会话（新时间戳日志文件）。返回新日志文件路径。
        /// </summary>
        public static string BeginSession(string sessionPrefix = "EasyWrite")
        {
            lock (Sync)
            {
                Directory.CreateDirectory(LogDirectory);
                return BeginSessionInternal(sessionPrefix);
            }
        }

        /// <summary>
        /// 直接写入当前会话日志（不经过 Debug，避免重复）。
        /// </summary>
        public static void WriteLine(string message)
        {
            if (string.IsNullOrEmpty(message))
            {
                return;
            }

            lock (Sync)
            {
                EnsureWriter();
                AppendLine(message);
            }
        }

        /// <summary>插件关闭时调用：移除监听并关闭文件。</summary>
        public static void Shutdown()
        {
            lock (Sync)
            {
                if (_listener != null)
                {
                    Debug.Listeners.Remove(_listener);
                    _listener.Dispose();
                    _listener = null;
                }

                CloseWriter();
            }
        }

        private static void EnsureWriter()
        {
            if (_writer == null)
            {
                BeginSessionInternal("EasyWrite");
            }
        }

        private static string BeginSessionInternal(string sessionPrefix, bool autoFlush = false)
        {
            CloseWriter();

            string safePrefix = SanitizeFileName(sessionPrefix);
            string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
            string fileName = $"{safePrefix}_{timestamp}.log";
            CurrentLogPath = Path.Combine(LogDirectory, fileName);

            _writer = new StreamWriter(CurrentLogPath, append: false, encoding: Encoding.UTF8)
            {
                AutoFlush = autoFlush
            };

            AppendLine("===== EasyWrite 日志会话 =====");
            AppendLine($"开始时间: {DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}");
            AppendLine($"文件: {CurrentLogPath}");
            AppendLine("================================");

            Debug.WriteLine($"[EasyWriteLog] 日志文件: {CurrentLogPath}");
            return CurrentLogPath;
        }

        private static void AppendLine(string message)
        {
            if (_writer == null)
            {
                return;
            }

            _writer.WriteLine($"{DateTime.Now:HH:mm:ss.fff} {message}");
        }

        private static void CloseWriter()
        {
            if (_writer == null)
            {
                return;
            }

            try
            {
                _writer.Flush();
                _writer.Dispose();
            }
            catch
            {
                // 关闭日志不应影响插件退出
            }
            finally
            {
                _writer = null;
            }
        }

        private static string ResolveLogDirectory()
        {
            // VSTO 影子拷贝下 Assembly.Location 指向 AppData，BaseDirectory 仍指向 bin\Debug
            string projectRoot = McpToolsHelpers.GetProjectRootDirectory();
            return Path.GetFullPath(Path.Combine(projectRoot, "logs"));
        }

        private static string SanitizeFileName(string name)
        {
            if (string.IsNullOrWhiteSpace(name))
            {
                return "EasyWrite";
            }

            name = name.Trim();
            foreach (char c in Path.GetInvalidFileNameChars())
            {
                name = name.Replace(c, '_');
            }

            return name;
        }

        private sealed class EasyWriteTraceListener : TraceListener
        {
            public override void Write(string message)
            {
                if (string.IsNullOrEmpty(message))
                {
                    return;
                }

                lock (Sync)
                {
                    if (_writer == null)
                    {
                        return;
                    }

                    _writer.Write($"{DateTime.Now:HH:mm:ss.fff} {message}");
                }
            }

            public override void WriteLine(string message)
            {
                if (string.IsNullOrEmpty(message))
                {
                    return;
                }

                lock (Sync)
                {
                    if (_writer == null)
                    {
                        return;
                    }

                    _writer.WriteLine($"{DateTime.Now:HH:mm:ss.fff} {message}");
                }
            }
        }
    }
}
