using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;

namespace WordAddIn1
{
    /// <summary>
    /// 会话工作区对账（打开对话 Pull；写工具 Around；B_* WS Pull）。
    /// </summary>
    public static class WorkspaceReconcile
    {
        public const double MtimeSkewSeconds = 1.0;
        public const int OpenConversationTtlSec = 150;
        public const int DefaultClientTtlSec = 930;
        public const int LockPollMs = 200;

        public static bool IsTempName(string name)
        {
            if (string.IsNullOrEmpty(name))
            {
                return true;
            }

            if (name.StartsWith(".", StringComparison.Ordinal)
                || name.StartsWith("~", StringComparison.Ordinal))
            {
                return true;
            }

            return name.EndsWith(".tmp", StringComparison.OrdinalIgnoreCase);
        }

        public static string NormalizeConversationId(string conversationId)
        {
            if (string.IsNullOrEmpty(conversationId) || conversationId == "-1")
            {
                return "_pending";
            }

            return conversationId.Trim();
        }

        public static async Task<(bool ok, string error)> PullForOpenConversationAsync()
        {
            if (!UserService.Instance.CheckLoginStatus()
                || string.IsNullOrWhiteSpace(UserService.Instance.WorkspaceRootEffective))
            {
                ConversationContext.WorkspaceReady = false;
                return (false, "用户未登录或工作区未初始化");
            }

            string conv = NormalizeConversationId(ConversationContext.CurrentId);
            var acquired = await AcquireAsync(OpenConversationTtlSec, OpenConversationTtlSec, "open-conversation")
                .ConfigureAwait(false);
            if (!acquired.ok)
            {
                ConversationContext.WorkspaceReady = false;
                return (false, acquired.error);
            }

            try
            {
                var pulled = await PullAsync(conv).ConfigureAwait(false);
                ConversationContext.WorkspaceReady = pulled.ok;
                return pulled;
            }
            finally
            {
                await BackendApiClient.ReleaseWorkspaceLockAsync(acquired.lockId).ConfigureAwait(false);
            }
        }

        public static async Task<ToolResult> PullAssumingLockedAsync(string conversationId)
        {
            string conv = NormalizeConversationId(
                string.IsNullOrEmpty(conversationId) ? ConversationContext.CurrentId : conversationId);
            var pulled = await PullAsync(conv).ConfigureAwait(false);
            ConversationContext.WorkspaceReady = pulled.ok;
            if (!pulled.ok)
            {
                return new ToolResult { Success = false, Error = "会话工作区对齐失败: " + pulled.error };
            }

            return new ToolResult { Success = true };
        }

        public static async Task<ToolResult> AroundFrontendWriteAsync(
            int ttlSec,
            Func<Task<ToolResult>> handler)
        {
            string conv = NormalizeConversationId(ConversationContext.CurrentId);
            int waitSec = Math.Max(1, ttlSec);
            var acquired = await AcquireAsync(ttlSec, waitSec, "frontend-write").ConfigureAwait(false);
            if (!acquired.ok)
            {
                return new ToolResult
                {
                    Success = false,
                    Error = string.IsNullOrEmpty(acquired.error)
                        ? "用户工作区正被占用，等待写锁超时"
                        : acquired.error
                };
            }

            try
            {
                var pulled = await PullAsync(conv).ConfigureAwait(false);
                if (!pulled.ok)
                {
                    return new ToolResult { Success = false, Error = "会话工作区对齐失败: " + pulled.error };
                }

                ConversationContext.WorkspaceReady = true;
                ToolResult result;
                try
                {
                    result = await handler().ConfigureAwait(false);
                }
                catch (Exception ex)
                {
                    result = new ToolResult { Success = false, Error = ex.Message };
                }

                var pushed = await PushAsync(conv).ConfigureAwait(false);
                if (!pushed.ok)
                {
                    return new ToolResult { Success = false, Error = "会话工作区对齐失败: " + pushed.error };
                }

                var equal = await TreesEqualAsync(conv).ConfigureAwait(false);
                if (!equal.ok)
                {
                    return new ToolResult { Success = false, Error = "会话工作区对齐失败: " + equal.error };
                }

                return result;
            }
            finally
            {
                await BackendApiClient.ReleaseWorkspaceLockAsync(acquired.lockId).ConfigureAwait(false);
            }
        }

        public static bool ShouldWrapFrontendWrite(string toolName, Dictionary<string, object> args)
        {
            if (string.IsNullOrEmpty(toolName) || !WrapTools.Contains(toolName))
            {
                return false;
            }

            if (string.Equals(toolName, "F_run_terminal", StringComparison.Ordinal))
            {
                return true;
            }

            if (string.Equals(toolName, "F_copy_workspace_file", StringComparison.Ordinal))
            {
                return HasWorkspacePath(args, "dest");
            }

            if (string.Equals(toolName, "F_read_excel_range", StringComparison.Ordinal))
            {
                return HasWorkspacePath(args, "export_csv") || HasWorkspacePath(args, "path");
            }

            if (string.Equals(toolName, "F_capture_document_page_image", StringComparison.Ordinal))
            {
                return HasWorkspacePath(args, "path");
            }

            if (string.Equals(toolName, "F_read_ppt_html", StringComparison.Ordinal))
            {
                return HasWorkspacePath(args, "path");
            }

            if (ExtractTools.Contains(toolName))
            {
                string raw = FilePathResolver.TryGetArg(args, "path");
                if (string.IsNullOrEmpty(raw))
                {
                    return true;
                }

                return HasWorkspacePath(args, "path");
            }

            if (HasWorkspacePath(args, "path") || HasWorkspacePath(args, "output_path"))
            {
                return true;
            }

            string path = FilePathResolver.TryGetArg(args, "path");
            string output = FilePathResolver.TryGetArg(args, "output_path");
            if (!string.IsNullOrEmpty(path) || !string.IsNullOrEmpty(output))
            {
                return false;
            }

            return true;
        }

        public static int TtlSecForTool(string toolName, Dictionary<string, object> args)
        {
            if (string.Equals(toolName, "F_run_terminal", StringComparison.Ordinal))
            {
                int timeout = 120;
                if (args != null && args.ContainsKey("timeout_sec") && args["timeout_sec"] != null)
                {
                    int.TryParse(args["timeout_sec"].ToString(), out timeout);
                    if (timeout < 1)
                    {
                        timeout = 120;
                    }
                }

                return timeout + 30;
            }

            return DefaultClientTtlSec;
        }

        private static readonly HashSet<string> WrapTools = new HashSet<string>(StringComparer.Ordinal)
        {
            "F_write_file",
            "F_copy_workspace_file",
            "F_modify_yaml_file",
            "F_csv_to_xml",
            "F_run_terminal",
            "F_browser_download",
            "F_read_excel_range",
            "F_capture_document_page_image",
            "F_read_ppt_html",
            "F_extract_table_format",
            "F_extract_table_data",
            "F_extract_chart_xml"
        };

        private static readonly HashSet<string> ExtractTools = new HashSet<string>(StringComparer.Ordinal)
        {
            "F_extract_table_format",
            "F_extract_table_data",
            "F_extract_chart_xml"
        };

        private static bool HasWorkspacePath(Dictionary<string, object> args, string key)
        {
            string raw = FilePathResolver.TryGetArg(args, key);
            if (string.IsNullOrEmpty(raw))
            {
                return false;
            }

            if (!FilePathResolver.TryResolve(raw, out ResolvedFilePath resolved, out _))
            {
                return true;
            }

            return resolved.Kind == FilePathKind.Workspace;
        }

        private static async Task<(bool ok, string lockId, string error)> AcquireAsync(
            int ttlSec,
            int waitSec,
            string reason)
        {
            DateTime deadline = DateTime.UtcNow.AddSeconds(Math.Max(1, waitSec));
            string conv = NormalizeConversationId(ConversationContext.CurrentId);
            while (true)
            {
                var attempt = await BackendApiClient
                    .TryAcquireWorkspaceLockAsync(ttlSec, conv, reason)
                    .ConfigureAwait(false);
                if (attempt.Acquired && !string.IsNullOrEmpty(attempt.LockId))
                {
                    return (true, attempt.LockId, null);
                }

                if (!attempt.HeldByOther)
                {
                    return (false, null, string.IsNullOrEmpty(attempt.Error) ? "申请写锁失败" : attempt.Error);
                }

                if (DateTime.UtcNow >= deadline)
                {
                    return (false, null, "用户工作区正被占用，等待写锁超时");
                }

                await Task.Delay(LockPollMs).ConfigureAwait(false);
            }
        }

        private static async Task<(bool ok, string error)> PullAsync(string conversationId)
        {
            var listed = await BackendApiClient.ListWorkspaceTreeAsync(conversationId).ConfigureAwait(false);
            if (!listed.ok)
            {
                return (false, listed.error);
            }

            Dictionary<string, LocalEntry> local = WalkLocal(conversationId);
            var backend = new Dictionary<string, BackendApiClient.WorkspaceFileStat>(StringComparer.OrdinalIgnoreCase);
            if (listed.files != null)
            {
                foreach (var file in listed.files)
                {
                    if (file == null || string.IsNullOrEmpty(file.RelativePath) || IsTempRelative(file.RelativePath))
                    {
                        continue;
                    }

                    backend[file.RelativePath.Replace('\\', '/')] = file;
                }
            }

            foreach (var pair in backend)
            {
                local.TryGetValue(pair.Key, out LocalEntry localEntry);
                if (SameFile(localEntry, pair.Value))
                {
                    continue;
                }

                string localPath = ToLocalPath(conversationId, pair.Key);
                if (string.IsNullOrEmpty(localPath))
                {
                    return (false, "无法解析本地路径: " + pair.Key);
                }

                if (IsHeldOpenOffice(localPath)
                    || (localEntry != null && IsHeldOpenOffice(localEntry.LocalPath)))
                {
                    continue;
                }

                bool downloaded = await BackendApiClient
                    .DownloadFileFromUserDirectoryAsync(pair.Key, localPath)
                    .ConfigureAwait(false);
                if (!downloaded || !File.Exists(localPath))
                {
                    return (false, "下载失败: " + pair.Key);
                }

                AlignMtime(localPath, pair.Value.MtimeUtc);
            }

            foreach (var pair in local)
            {
                if (!backend.ContainsKey(pair.Key) && File.Exists(pair.Value.LocalPath))
                {
                    if (IsHeldOpenOffice(pair.Value.LocalPath))
                    {
                        continue;
                    }

                    try
                    {
                        File.Delete(pair.Value.LocalPath);
                    }
                    catch (IOException)
                    {
                        if (IsHeldOpenOffice(pair.Value.LocalPath) || HasOfficeLockFile(pair.Value.LocalPath))
                        {
                            continue;
                        }

                        return (false, "删除本地多余文件失败: 文件被占用 (" + pair.Key + ")");
                    }
                    catch (Exception ex)
                    {
                        return (false, "删除本地多余文件失败: " + ex.Message);
                    }
                }
            }

            return (true, null);
        }

        private static async Task<(bool ok, string error)> PushAsync(string conversationId)
        {
            var listed = await BackendApiClient.ListWorkspaceTreeAsync(conversationId).ConfigureAwait(false);
            if (!listed.ok)
            {
                return (false, listed.error);
            }

            Dictionary<string, LocalEntry> local = WalkLocal(conversationId);
            var backend = new Dictionary<string, BackendApiClient.WorkspaceFileStat>(StringComparer.OrdinalIgnoreCase);
            if (listed.files != null)
            {
                foreach (var file in listed.files)
                {
                    if (file == null || string.IsNullOrEmpty(file.RelativePath) || IsTempRelative(file.RelativePath))
                    {
                        continue;
                    }

                    backend[file.RelativePath.Replace('\\', '/')] = file;
                }
            }

            foreach (var pair in local)
            {
                if (IsHeldOpenOffice(pair.Value.LocalPath))
                {
                    continue;
                }

                backend.TryGetValue(pair.Key, out BackendApiClient.WorkspaceFileStat remote);
                if (SameFile(pair.Value, remote))
                {
                    continue;
                }

                bool uploaded = await BackendApiClient
                    .UploadUserFileAsync(pair.Value.LocalPath, pair.Key)
                    .ConfigureAwait(false);
                if (!uploaded)
                {
                    return (false, "上传失败: " + pair.Key);
                }
            }

            foreach (var pair in backend)
            {
                if (!local.ContainsKey(pair.Key))
                {
                    bool deleted = await BackendApiClient.DeleteWorkspaceFileAsync(pair.Key).ConfigureAwait(false);
                    if (!deleted)
                    {
                        return (false, "删除云端文件失败: " + pair.Key);
                    }
                }
            }

            return (true, null);
        }

        private static async Task<(bool ok, string error)> TreesEqualAsync(string conversationId)
        {
            var listed = await BackendApiClient.ListWorkspaceTreeAsync(conversationId).ConfigureAwait(false);
            if (!listed.ok)
            {
                return (false, listed.error);
            }

            Dictionary<string, LocalEntry> local = WalkLocal(conversationId);
            var backendKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            if (listed.files != null)
            {
                foreach (var file in listed.files)
                {
                    if (file == null || string.IsNullOrEmpty(file.RelativePath) || IsTempRelative(file.RelativePath))
                    {
                        continue;
                    }

                    string key = file.RelativePath.Replace('\\', '/');
                    backendKeys.Add(key);
                    local.TryGetValue(key, out LocalEntry localEntry);
                    if (localEntry != null && IsHeldOpenOffice(localEntry.LocalPath))
                    {
                        continue;
                    }

                    if (!SameFile(localEntry, file))
                    {
                        return (false, "内容不一致: " + key);
                    }
                }
            }

            foreach (var pair in local)
            {
                if (backendKeys.Contains(pair.Key) || IsHeldOpenOffice(pair.Value.LocalPath))
                {
                    continue;
                }

                return (false, "前端多出文件: " + pair.Key);
            }

            return (true, null);
        }

        private static bool IsTempRelative(string relativePath)
        {
            string name = Path.GetFileName(relativePath.Replace('/', Path.DirectorySeparatorChar));
            return IsTempName(name);
        }

        /// <summary>
        /// 渠道里正打开的办公文件，或同目录存在 Office <c>~$</c> 锁文件。对账跳过（不传、不覆盖、不删）。
        /// </summary>
        public static bool IsHeldOpenOffice(string localPath)
        {
            if (string.IsNullOrEmpty(localPath))
            {
                return false;
            }

            string full;
            try
            {
                full = Path.GetFullPath(localPath);
            }
            catch
            {
                return false;
            }

            foreach (IOperationChannel channel in ChannelRegistry.Snapshot())
            {
                string open = TryChannelFilePath(channel);
                if (string.IsNullOrEmpty(open))
                {
                    continue;
                }

                string openFull;
                try
                {
                    openFull = Path.GetFullPath(open);
                }
                catch
                {
                    continue;
                }

                if (string.Equals(full, openFull, StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
            }

            return HasOfficeLockFile(full);
        }

        private static string TryChannelFilePath(IOperationChannel channel)
        {
            if (channel is WordChannel word)
            {
                return word.FilePath;
            }

            if (channel is WpsChannel wps)
            {
                return wps.FilePath;
            }

            if (channel is ExcelChannel excel)
            {
                return excel.FilePath;
            }

            if (channel is EtChannel et)
            {
                return et.FilePath;
            }

            if (channel is PptChannel ppt)
            {
                return ppt.FilePath;
            }

            if (channel is WppChannel wpp)
            {
                return wpp.FilePath;
            }

            return null;
        }

        private static bool HasOfficeLockFile(string fullPath)
        {
            string dir = Path.GetDirectoryName(fullPath);
            string name = Path.GetFileName(fullPath);
            if (string.IsNullOrEmpty(dir) || string.IsNullOrEmpty(name))
            {
                return false;
            }

            if (File.Exists(Path.Combine(dir, "~$" + name)))
            {
                return true;
            }

            if (name.Length > 2 && File.Exists(Path.Combine(dir, "~$" + name.Substring(2))))
            {
                return true;
            }

            return false;
        }

        private sealed class LocalEntry
        {
            public string LocalPath;
            public long Size;
            public DateTime MtimeUtc;
        }

        private static Dictionary<string, LocalEntry> WalkLocal(string conversationId)
        {
            var map = new Dictionary<string, LocalEntry>(StringComparer.OrdinalIgnoreCase);
            string sessionDir;
            try
            {
                sessionDir = WorkspacePathResolver.GetSessionDirectory(conversationId);
            }
            catch
            {
                return map;
            }

            if (!Directory.Exists(sessionDir))
            {
                return map;
            }

            string prefix = "sessions/" + NormalizeConversationId(conversationId) + "/";
            foreach (string file in Directory.EnumerateFiles(sessionDir, "*", SearchOption.AllDirectories))
            {
                string name = Path.GetFileName(file);
                if (IsTempName(name))
                {
                    continue;
                }

                string relUnder = ToRelative(file, sessionDir);
                if (string.IsNullOrEmpty(relUnder))
                {
                    continue;
                }

                var info = new FileInfo(file);
                map[prefix + relUnder] = new LocalEntry
                {
                    LocalPath = file,
                    Size = info.Length,
                    MtimeUtc = info.LastWriteTimeUtc
                };
            }

            return map;
        }

        private static string ToLocalPath(string conversationId, string apiRelative)
        {
            if (string.IsNullOrEmpty(apiRelative))
            {
                return null;
            }

            string normalized = apiRelative.Replace('\\', '/');
            string prefix = "sessions/" + NormalizeConversationId(conversationId) + "/";
            if (!normalized.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
            {
                return null;
            }

            string rest = normalized.Substring(prefix.Length);
            string sessionDir = WorkspacePathResolver.GetSessionDirectory(conversationId);
            string full = sessionDir;
            foreach (string part in rest.Split(new[] { '/' }, StringSplitOptions.RemoveEmptyEntries))
            {
                if (part == ".." || part == ".")
                {
                    return null;
                }

                full = Path.Combine(full, part);
            }

            string parent = Path.GetDirectoryName(full);
            if (!string.IsNullOrEmpty(parent))
            {
                Directory.CreateDirectory(parent);
            }

            return full;
        }

        private static string ToRelative(string full, string root)
        {
            string fullNorm = Path.GetFullPath(full);
            string rootNorm = Path.GetFullPath(root);
            if (!fullNorm.StartsWith(rootNorm, StringComparison.OrdinalIgnoreCase))
            {
                return null;
            }

            string rest = fullNorm.Substring(rootNorm.Length).TrimStart(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
            return rest.Replace('\\', '/');
        }

        private static bool SameFile(LocalEntry local, BackendApiClient.WorkspaceFileStat remote)
        {
            if (local == null || remote == null)
            {
                return false;
            }

            if (local.Size != remote.Size)
            {
                return false;
            }

            if (remote.MtimeUtc == default(DateTime))
            {
                return true;
            }

            return Math.Abs((local.MtimeUtc - remote.MtimeUtc).TotalSeconds) <= MtimeSkewSeconds;
        }

        private static void AlignMtime(string localPath, DateTime utc)
        {
            try
            {
                if (utc == default(DateTime))
                {
                    return;
                }

                DateTime value = utc.Kind == DateTimeKind.Utc ? utc : utc.ToUniversalTime();
                File.SetLastWriteTimeUtc(localPath, value);
            }
            catch
            {
            }
        }
    }
}
