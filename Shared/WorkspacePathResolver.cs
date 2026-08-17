using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;

namespace WordAddIn1
{
    public class NoWritableWorkspaceDriveException : Exception
    {
        public NoWritableWorkspaceDriveException()
            : base("本机无可用工作区盘符")
        {
        }
    }

    /// <summary>
    /// 前端工作区路径解析（D33 remap、sessions 目录、读写搜索）。
    /// 本地路径：{用户根}/sessions/{id}/（用户根默认 D:/YiWrite/WorkSpace）。
    /// </summary>
    public static class WorkspacePathResolver
    {
        public const string DefaultRoot = @"D:/YiWrite/WorkSpace";

        private static readonly HashSet<string> WindowsReservedNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "CON", "PRN", "AUX", "NUL",
            "COM1", "COM2", "COM3", "COM4", "COM5", "COM6", "COM7", "COM8", "COM9",
            "LPT1", "LPT2", "LPT3", "LPT4", "LPT5", "LPT6", "LPT7", "LPT8", "LPT9"
        };

        public static string ResolveWorkspaceRoot(string mysqlPath)
        {
            var path = string.IsNullOrWhiteSpace(mysqlPath) ? DefaultRoot : mysqlPath.Trim();
            path = path.Replace('\\', '/');

            if (DriveExistsAndWritable(path))
            {
                return NormalizePathSlashes(path);
            }

            var drive = ParseDriveLetter(path);
            var suffix = ExtractPathSuffix(path);

            if (DriveExists(drive))
            {
                throw new InvalidOperationException($"工作目录不可写：{path}");
            }

            if (!string.Equals(drive, "C:", StringComparison.OrdinalIgnoreCase) && DriveExists("C:"))
            {
                return "C:" + suffix;
            }

            throw new NoWritableWorkspaceDriveException();
        }

        public static string EnsureEffectiveRoot(string mysqlPath)
        {
            var effective = ResolveWorkspaceRoot(mysqlPath);
            try
            {
                Directory.CreateDirectory(effective);
                if (!IsWritableDirectory(effective))
                {
                    throw new InvalidOperationException($"工作目录不可写：{effective}");
                }
            }
            catch (Exception ex) when (!(ex is InvalidOperationException) && !(ex is NoWritableWorkspaceDriveException))
            {
                throw new InvalidOperationException($"无法创建工作目录：{effective}", ex);
            }

            return NormalizePathSlashes(effective);
        }

        public static bool IsRemapped(string configured, string effective)
        {
            return !string.Equals(
                NormalizePathSlashes(configured ?? ""),
                NormalizePathSlashes(effective ?? ""),
                StringComparison.OrdinalIgnoreCase);
        }

        public static string GetSessionDirectory(string conversationId = null)
        {
            var root = UserService.Instance.WorkspaceRootEffective;
            if (string.IsNullOrWhiteSpace(root))
            {
                throw new InvalidOperationException("工作区根路径未初始化，请先登录并配置工作目录");
            }

            var conv = NormalizeConversationId(conversationId ?? ConversationContext.CurrentId);
            var dir = Path.Combine(root, "sessions", conv);
            Directory.CreateDirectory(dir);
            return dir;
        }

        public static string ResolveWritePath(string filename, string conversationId = null)
        {
            var relative = SanitizeWorkspaceRelativePath(filename);
            if (relative == null)
            {
                throw new ArgumentException($"非法文件名: {filename}", nameof(filename));
            }

            var sessionDir = GetSessionDirectory(conversationId);
            string full = sessionDir;
            foreach (var part in relative.Split('/'))
            {
                full = Path.Combine(full, part);
            }

            string parent = Path.GetDirectoryName(full);
            if (!string.IsNullOrEmpty(parent))
            {
                Directory.CreateDirectory(parent);
            }

            return full;
        }

        public static string ResolveReadPath(string filename, string conversationId = null)
        {
            var relative = SanitizeWorkspaceRelativePath(filename);
            if (relative == null)
            {
                return null;
            }

            var conv = NormalizeConversationId(conversationId ?? ConversationContext.CurrentId);
            var root = UserService.Instance.WorkspaceRootEffective;
            if (string.IsNullOrWhiteSpace(root))
            {
                return null;
            }

            var mode = ConfigManager.GetWorkspaceMode();

            var candidates = new List<string>();
            if (mode == "sessions")
            {
                candidates.Add(CombineUnder(Path.Combine(root, "sessions", conv), relative));
                candidates.Add(CombineUnder(Path.Combine(root, "sessions", "_pending"), relative));
            }
            else
            {
                candidates.Add(CombineUnder(Path.Combine(root, "sessions", conv), relative));
                foreach (var sessDir in SortedOtherSessionDirs(root, conv))
                {
                    candidates.Add(CombineUnder(sessDir, relative));
                }

                candidates.Add(CombineUnder(root, relative));
                candidates.Add(CombineUnder(Path.Combine(root, "sessions", "_pending"), relative));
            }

            string primary = candidates.Count > 0 ? candidates[0] : null;
            foreach (var path in candidates)
            {
                if (!string.IsNullOrEmpty(path) && File.Exists(path))
                {
#if DEBUG
                    if (mode == "history" && primary != null
                        && !string.Equals(Path.GetFullPath(path), Path.GetFullPath(primary), StringComparison.OrdinalIgnoreCase))
                    {
                        Debug.WriteLine($"resolved_from={path}");
                    }
#endif
                    return path;
                }
            }

            return null;
        }

        public static string BuildRelativePath(string conversationId, string filename)
        {
            var relative = SanitizeWorkspaceRelativePath(filename);
            if (relative == null)
            {
                throw new ArgumentException($"非法文件名: {filename}", nameof(filename));
            }

            return $"sessions/{NormalizeConversationId(conversationId ?? ConversationContext.CurrentId)}/{relative}";
        }

        /// <summary>
        /// 工作区内相对路径：裸文件名，或少量子目录（如 ppt_images/sid256-s5.png）。
        /// 禁止 ..、盘符、绝对路径。
        /// </summary>
        public static string SanitizeWorkspaceRelativePath(string relativePath)
        {
            if (string.IsNullOrWhiteSpace(relativePath))
            {
                return null;
            }

            var norm = relativePath.Replace('\\', '/').Trim();
            while (norm.StartsWith("/", StringComparison.Ordinal))
            {
                norm = norm.Substring(1);
            }

            if (string.IsNullOrEmpty(norm)
                || norm.IndexOf("..", StringComparison.Ordinal) >= 0
                || norm.IndexOf(':') >= 0
                || Path.IsPathRooted(relativePath.Trim()))
            {
                return null;
            }

            var parts = norm.Split(new[] { '/' }, StringSplitOptions.RemoveEmptyEntries);
            if (parts.Length == 0 || parts.Length > 4)
            {
                return null;
            }

            var cleaned = new List<string>(parts.Length);
            foreach (var part in parts)
            {
                var safe = SanitizeFilename(part);
                if (safe == null)
                {
                    return null;
                }

                cleaned.Add(safe);
            }

            return string.Join("/", cleaned);
        }

        private static string CombineUnder(string baseDir, string relativeWithSlashes)
        {
            if (string.IsNullOrEmpty(baseDir) || string.IsNullOrEmpty(relativeWithSlashes))
            {
                return null;
            }

            string full = baseDir;
            foreach (var part in relativeWithSlashes.Split('/'))
            {
                full = Path.Combine(full, part);
            }

            return full;
        }

        public static string SanitizeFilename(string filename)
        {
            if (string.IsNullOrWhiteSpace(filename))
            {
                return null;
            }

            var name = Path.GetFileName(filename.Replace('\\', '/').Trim());
            if (string.IsNullOrEmpty(name) || name == "." || name == "..")
            {
                return null;
            }

            if (name.Contains("..") || name.Contains("/") || name.Contains("\\") || name.Contains(":"))
            {
                return null;
            }

            var stem = name.Contains(".") ? name.Substring(0, name.LastIndexOf('.')) : name;
            if (WindowsReservedNames.Contains(stem.ToUpperInvariant()))
            {
                return null;
            }

            return name;
        }

        public static bool IsBlockedSystemPath(string path)
        {
            if (string.IsNullOrWhiteSpace(path))
            {
                return true;
            }

            string full;
            try
            {
                full = Path.GetFullPath(path).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar)
                    + Path.DirectorySeparatorChar;
            }
            catch
            {
                return true;
            }

            string[] blocked =
            {
                Environment.GetFolderPath(Environment.SpecialFolder.Windows) + Path.DirectorySeparatorChar,
                Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles) + Path.DirectorySeparatorChar,
                Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86) + Path.DirectorySeparatorChar
            };

            foreach (var prefix in blocked)
            {
                if (!string.IsNullOrEmpty(prefix)
                    && full.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
            }

            return false;
        }

        public static bool ValidateWorkspacePathForSave(string path, out string errorMessage)
        {
            errorMessage = null;
            if (string.IsNullOrWhiteSpace(path))
            {
                errorMessage = "工作目录不能为空";
                return false;
            }

            if (IsBlockedSystemPath(path))
            {
                errorMessage = "不能选择系统目录，请重新选择";
                return false;
            }

            try
            {
                Directory.CreateDirectory(path);
                if (!IsWritableDirectory(path))
                {
                    errorMessage = "目录不可写，请重新选择";
                    return false;
                }
            }
            catch (Exception ex)
            {
                errorMessage = $"目录无效：{ex.Message}";
                return false;
            }

            return true;
        }

        private static string NormalizeConversationId(string id)
        {
            return string.IsNullOrEmpty(id) || id == "-1" ? "_pending" : id;
        }

        private static List<string> SortedOtherSessionDirs(string agentRoot, string excludeConv)
        {
            var sessionsRoot = Path.Combine(agentRoot, "sessions");
            if (!Directory.Exists(sessionsRoot))
            {
                return new List<string>();
            }

            var entries = new List<Tuple<DateTime, string>>();
            foreach (var dir in Directory.GetDirectories(sessionsRoot))
            {
                var name = Path.GetFileName(dir);
                if (name == "_pending" || string.Equals(name, excludeConv, StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                DateTime mtime;
                try
                {
                    mtime = Directory.GetLastWriteTimeUtc(dir);
                }
                catch
                {
                    mtime = DateTime.MinValue;
                }

                entries.Add(Tuple.Create(mtime, dir));
            }

            entries.Sort((a, b) => b.Item1.CompareTo(a.Item1));
            return entries.Select(e => e.Item2).ToList();
        }

        private static string ParseDriveLetter(string path)
        {
            if (path.Length >= 2 && path[1] == ':')
            {
                return path.Substring(0, 2).ToUpperInvariant();
            }

            return "";
        }

        private static string ExtractPathSuffix(string path)
        {
            if (path.Length >= 2 && path[1] == ':')
            {
                var suffix = path.Substring(2);
                if (string.IsNullOrEmpty(suffix))
                {
                    return "/";
                }

                return suffix.StartsWith("/") ? suffix : "/" + suffix;
            }

            return "/" + path.TrimStart('/');
        }

        private static bool DriveExists(string drive)
        {
            if (string.IsNullOrEmpty(drive) || drive.Length < 2 || drive[1] != ':')
            {
                return false;
            }

            try
            {
                var letter = drive[0];
                return DriveInfo.GetDrives().Any(d =>
                    char.ToUpperInvariant(d.Name[0]) == char.ToUpperInvariant(letter) && d.IsReady);
            }
            catch
            {
                return false;
            }
        }

        private static bool DriveExistsAndWritable(string path)
        {
            var drive = ParseDriveLetter(path.Replace('\\', '/'));
            if (!DriveExists(drive))
            {
                return false;
            }

            try
            {
                var dir = path.Replace('/', Path.DirectorySeparatorChar);
                if (!Directory.Exists(dir))
                {
                    Directory.CreateDirectory(dir);
                }

                return IsWritableDirectory(dir);
            }
            catch
            {
                return false;
            }
        }

        private static bool IsWritableDirectory(string path)
        {
            try
            {
                Directory.CreateDirectory(path);
                var testFile = Path.Combine(path, ".ew_write_test_" + Guid.NewGuid().ToString("N"));
                File.WriteAllText(testFile, "ok");
                File.Delete(testFile);
                return true;
            }
            catch
            {
                return false;
            }
        }

        private static string NormalizePathSlashes(string path)
        {
            return path?.Replace('\\', '/') ?? "";
        }
    }
}
