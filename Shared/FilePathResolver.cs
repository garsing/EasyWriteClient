using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Threading.Tasks;

namespace WordAddIn1
{
    public enum FilePathKind
    {
        Absolute,
        Workspace
    }

    public sealed class ResolvedFilePath
    {
        public FilePathKind Kind { get; set; }
        public string LocalPath { get; set; }
        public string Relative { get; set; }
        public string Display { get; set; }
    }

    public sealed class FilePathIoResult
    {
        public bool Success { get; set; }
        public string Error { get; set; }
        public string Text { get; set; }
        public byte[] Bytes { get; set; }
        public ResolvedFilePath Path { get; set; }
    }

    /// <summary>
    /// 磁盘路径统一入口（D1 / D13 / D18 / D22）。清单内 F_* 只调本类型。
    /// </summary>
    public static class FilePathResolver
    {
        private const int MaxRelativeDepth = 8;
        private const string WorkspacePrefix = "workspace:";

        private static readonly HashSet<string> WindowsReservedNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "CON", "PRN", "AUX", "NUL",
            "COM1", "COM2", "COM3", "COM4", "COM5", "COM6", "COM7", "COM8", "COM9",
            "LPT1", "LPT2", "LPT3", "LPT4", "LPT5", "LPT6", "LPT7", "LPT8", "LPT9"
        };

        public static string TryGetArg(Dictionary<string, object> args, string primary = "path", params string[] aliases)
        {
            if (args == null)
            {
                return null;
            }

            string value = ReadArg(args, primary);
            if (!string.IsNullOrEmpty(value))
            {
                return value;
            }

            if (aliases == null)
            {
                return null;
            }

            foreach (string alias in aliases)
            {
                value = ReadArg(args, alias);
                if (!string.IsNullOrEmpty(value))
                {
                    return value;
                }
            }

            return null;
        }

        public static bool TryResolveFromArgs(
            Dictionary<string, object> args,
            out ResolvedFilePath resolved,
            out string error,
            string primary = "path",
            params string[] aliases)
        {
            return TryResolve(TryGetArg(args, primary, aliases), out resolved, out error);
        }

        public static bool TryResolve(string raw, out ResolvedFilePath resolved, out string error, string conversationId = null)
        {
            resolved = null;
            error = null;

            string input = StripWorkspacePrefix(raw?.Trim() ?? "");
            if (string.IsNullOrEmpty(input))
            {
                error = "必须提供 path 参数";
                return false;
            }

            if (IsAbsolutePath(input))
            {
                return TryResolveAbsolute(input, out resolved, out error);
            }

            return TryResolveWorkspace(input, conversationId, out resolved, out error);
        }

        public static async Task<FilePathIoResult> ReadAsync(ResolvedFilePath resolved)
        {
            var ensured = await EnsureLocalFileAsync(resolved).ConfigureAwait(false);
            if (!ensured.Success)
            {
                return ensured;
            }

            try
            {
                using (var reader = new StreamReader(resolved.LocalPath, true))
                {
                    string text = reader.ReadToEnd();
                    return new FilePathIoResult
                    {
                        Success = true,
                        Text = text,
                        Path = resolved
                    };
                }
            }
            catch (Exception ex)
            {
                return FailIo($"读取文件失败: {ex.Message}", resolved);
            }
        }

        public static async Task<FilePathIoResult> ReadBytesAsync(ResolvedFilePath resolved)
        {
            var ensured = await EnsureLocalFileAsync(resolved).ConfigureAwait(false);
            if (!ensured.Success)
            {
                return ensured;
            }

            try
            {
                byte[] bytes = File.ReadAllBytes(resolved.LocalPath);
                return new FilePathIoResult
                {
                    Success = true,
                    Bytes = bytes,
                    Path = resolved
                };
            }
            catch (Exception ex)
            {
                return FailIo($"读取文件失败: {ex.Message}", resolved);
            }
        }

        public static async Task<FilePathIoResult> WriteAsync(ResolvedFilePath resolved, string text, Encoding encoding = null)
        {
            if (resolved == null)
            {
                return FailIo("路径未解析");
            }

            encoding = encoding ?? Encoding.UTF8;
            return await WriteBytesAsync(resolved, encoding.GetBytes(text ?? "")).ConfigureAwait(false);
        }

        public static async Task<FilePathIoResult> WriteBytesAsync(ResolvedFilePath resolved, byte[] bytes)
        {
            if (resolved == null || string.IsNullOrEmpty(resolved.LocalPath))
            {
                return FailIo("路径未解析");
            }

            try
            {
                string parent = Path.GetDirectoryName(resolved.LocalPath);
                if (!string.IsNullOrEmpty(parent))
                {
                    Directory.CreateDirectory(parent);
                }

                File.WriteAllBytes(resolved.LocalPath, bytes ?? Array.Empty<byte>());
            }
            catch (Exception ex)
            {
                return FailIo($"写入文件失败: {ex.Message}", resolved);
            }

            if (resolved.Kind == FilePathKind.Workspace)
            {
                if (string.IsNullOrEmpty(resolved.Relative))
                {
                    return FailIo("工作区相对路径无效", resolved);
                }

                bool uploaded = await McpToolsHelpers
                    .UploadWorkspaceFileAsync(resolved.LocalPath, resolved.Relative)
                    .ConfigureAwait(false);
                if (!uploaded)
                {
                    return FailIo("已写入本机，但同步到后端会话目录失败", resolved);
                }
            }

            return new FilePathIoResult
            {
                Success = true,
                Path = resolved
            };
        }

        private static async Task<FilePathIoResult> EnsureLocalFileAsync(ResolvedFilePath resolved)
        {
            if (resolved == null || string.IsNullOrEmpty(resolved.LocalPath))
            {
                return FailIo("路径未解析");
            }

            if (resolved.Kind == FilePathKind.Absolute)
            {
                if (!File.Exists(resolved.LocalPath))
                {
                    return FailIo($"文件不存在: {resolved.Display}", resolved);
                }

                return new FilePathIoResult { Success = true, Path = resolved };
            }

            var ensure = await McpToolsHelpers.EnsureWorkspaceFileAsync(resolved.Relative).ConfigureAwait(false);
            if (!ensure.success || string.IsNullOrEmpty(ensure.localPath) || !File.Exists(ensure.localPath))
            {
                return FailIo(string.IsNullOrEmpty(ensure.error) ? $"文件不存在: {resolved.Display}" : ensure.error, resolved);
            }

            resolved.LocalPath = ensure.localPath;
            return new FilePathIoResult { Success = true, Path = resolved };
        }

        private static bool TryResolveAbsolute(string input, out ResolvedFilePath resolved, out string error)
        {
            resolved = null;
            error = null;
            string full;
            try
            {
                full = Path.GetFullPath(input);
            }
            catch (Exception ex)
            {
                error = "路径非法: " + ex.Message;
                return false;
            }

            if (WorkspacePathResolver.IsBlockedSystemPath(full))
            {
                error = "不能读写系统目录（Windows / Program Files）";
                return false;
            }

            resolved = new ResolvedFilePath
            {
                Kind = FilePathKind.Absolute,
                LocalPath = full,
                Relative = null,
                Display = full
            };
            return true;
        }

        private static bool TryResolveWorkspace(string input, string conversationId, out ResolvedFilePath resolved, out string error)
        {
            resolved = null;
            error = null;

            string sessionDir;
            try
            {
                sessionDir = WorkspacePathResolver.GetSessionDirectory(conversationId);
            }
            catch (Exception ex)
            {
                error = ex.Message;
                return false;
            }

            string combined = CombineUnderSession(sessionDir, input);
            if (combined == null)
            {
                error = "非法工作区路径";
                return false;
            }

            string full;
            string sessionFull;
            try
            {
                full = Path.GetFullPath(combined);
                sessionFull = Path.GetFullPath(sessionDir);
            }
            catch (Exception ex)
            {
                error = "路径非法: " + ex.Message;
                return false;
            }

            if (!IsUnderSession(full, sessionFull))
            {
                error = "不得逃出会话目录；出区请使用绝对路径";
                return false;
            }

            string relative = ToRelative(full, sessionFull);
            if (string.IsNullOrEmpty(relative))
            {
                error = "工作区 path 必须指向文件，不能是会话根目录";
                return false;
            }

            string relativeError;
            if (!IsSafeRelative(relative, out relativeError))
            {
                error = relativeError;
                return false;
            }

            resolved = new ResolvedFilePath
            {
                Kind = FilePathKind.Workspace,
                LocalPath = full,
                Relative = relative,
                Display = relative
            };
            return true;
        }

        private static bool IsAbsolutePath(string path)
        {
            if (string.IsNullOrEmpty(path))
            {
                return false;
            }

            if (path.StartsWith(@"\\", StringComparison.Ordinal))
            {
                return true;
            }

            return path.Length >= 3
                && char.IsLetter(path[0])
                && path[1] == ':'
                && (path[2] == '\\' || path[2] == '/');
        }

        private static string StripWorkspacePrefix(string path)
        {
            if (string.IsNullOrEmpty(path))
            {
                return path;
            }

            if (path.StartsWith(WorkspacePrefix, StringComparison.OrdinalIgnoreCase))
            {
                return path.Substring(WorkspacePrefix.Length).Trim();
            }

            return path;
        }

        private static string CombineUnderSession(string sessionDir, string relativeRaw)
        {
            string norm = relativeRaw.Replace('/', Path.DirectorySeparatorChar).Trim();
            while (norm.StartsWith("\\", StringComparison.Ordinal) || norm.StartsWith("/", StringComparison.Ordinal))
            {
                norm = norm.Substring(1);
            }

            if (string.IsNullOrEmpty(norm))
            {
                return null;
            }

            return Path.Combine(sessionDir, norm);
        }

        private static bool IsUnderSession(string fullPath, string sessionDir)
        {
            string sessionPrefix = sessionDir.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar)
                + Path.DirectorySeparatorChar;
            string full = fullPath.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar)
                + Path.DirectorySeparatorChar;
            return full.StartsWith(sessionPrefix, StringComparison.OrdinalIgnoreCase);
        }

        private static string ToRelative(string fullPath, string sessionDir)
        {
            string sessionPrefix = sessionDir.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar)
                + Path.DirectorySeparatorChar;
            string full = Path.GetFullPath(fullPath);
            if (!full.StartsWith(sessionPrefix, StringComparison.OrdinalIgnoreCase))
            {
                if (string.Equals(
                    full.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar),
                    sessionDir.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar),
                    StringComparison.OrdinalIgnoreCase))
                {
                    return "";
                }

                return null;
            }

            return full.Substring(sessionPrefix.Length).Replace('\\', '/');
        }

        private static bool IsSafeRelative(string relative, out string error)
        {
            error = null;
            var parts = relative.Split(new[] { '/' }, StringSplitOptions.RemoveEmptyEntries);
            if (parts.Length == 0)
            {
                error = "工作区 path 必须指向文件";
                return false;
            }

            if (parts.Length > MaxRelativeDepth)
            {
                error = "工作区相对路径过深";
                return false;
            }

            foreach (string part in parts)
            {
                if (part == "." || part == "..")
                {
                    error = "非法工作区路径";
                    return false;
                }

                var stem = part.Contains(".") ? part.Substring(0, part.LastIndexOf('.')) : part;
                if (WindowsReservedNames.Contains(stem))
                {
                    error = "非法文件名: " + part;
                    return false;
                }
            }

            return true;
        }

        private static string ReadArg(Dictionary<string, object> args, string key)
        {
            if (string.IsNullOrEmpty(key) || !args.ContainsKey(key) || args[key] == null)
            {
                return null;
            }

            string value = args[key].ToString()?.Trim();
            return string.IsNullOrEmpty(value) ? null : value;
        }

        private static FilePathIoResult FailIo(string error, ResolvedFilePath path = null)
        {
            return new FilePathIoResult
            {
                Success = false,
                Error = error,
                Path = path
            };
        }
    }
}
